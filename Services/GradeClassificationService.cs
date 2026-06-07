// Services/GradeClassificationService.cs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Dossier.Models;

namespace Dossier.Services
{
    public class GradeClassificationService : IGradeClassificationService
    {
        private readonly IEquivalencyService _equivalencyService;

        private readonly Regex _fractionRegex = new Regex(@"(\d+(?:\.\d+)?)\s*(?:/|out of|of)\s*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);
        private readonly Regex _percentRegex = new Regex(@"(\d+(?:\.\d+)?)%", RegexOptions.IgnoreCase);
        private readonly Regex _priorityKeywordRegex = new Regex(@"(?:average|final|cgpa|overall|gpa)\s*[:\-\s]*\s*(\d{1,3}(?:\.\d+)?)", RegexOptions.IgnoreCase);
        private readonly Regex _generalGradeRegex = new Regex(@"\b(?!(?:1\.1|2\.1|2\.2))\d{1,3}(?:\.\d+)?\b");

        // Matches "@"-style threshold lines: "UK 2.1: Bachelors @ 7.5", "1st: Bachelors @ 90%", "2.2: Bachelors @ CGPA of 3.0/4.0"
        // After "@", skips optional non-numeric prefix words (e.g. "CGPA of", "GPA of") before capturing the number.
        private readonly Regex _thresholdLineRegex = new Regex(
            @"(?:UK\s+)?(?:(1st|first)|(2\.1|2:1)|(2\.2|2:2)|(3rd|third))[^@\n]{0,60}@[^0-9\n]{0,30}(\d+(?:[.,]\d+)?)",
            RegexOptions.IgnoreCase);

        // Matches range-style threshold lines: "2.2: 12 - 14", "2.1: 15 - 17", "1st: 18 - 20"
        // Used by institutions like Aberdeen that publish a banded scale rather than a single cut-off.
        private readonly Regex _rangeThresholdRegex = new Regex(
            @"(?:UK\s+)?(?:(1st|first)|(2\.1|2:1)|(2\.2|2:2)|(3rd|third))\s*:\s*(\d+(?:[.,]\d+)?)\s*[-–]\s*(\d+(?:[.,]\d+)?)",
            RegexOptions.IgnoreCase);

        // Extracts the student's grade from the narrative portion of the note
        // Handles: "grade of X", "average is X", "average = X", "current average = X%", "cgpa of X"
        private readonly Regex _studentGradeNarrativeRegex = new Regex(
            @"\bgrade\s+(?:of|is)\s*(\d+(?:[.,]\d+)?)|\baverage\s*(?:of\s+|is\s+|=\s*)(\d+(?:[.,]\d+)?)|\bcgpa\s*(?:of|is|:)\s*(\d+(?:[.,]\d+)?)",
            RegexOptions.IgnoreCase);

        // For dual-system notes: finds every "grade of X" mention and denominators from threshold lines
        private readonly Regex _allGradeOfRegex = new Regex(
            @"\bgrade\s+of\s*(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);
        // Extracts the denominator from threshold lines like "@ 3.7/4.0" or "@ CGPA of 3.7/4.0"
        private readonly Regex _thresholdDenomRegex = new Regex(
            @"@[^0-9\n]{0,30}\d+(?:[.,]\d+)?/(\d+(?:\.\d+)?)", RegexOptions.IgnoreCase);

        public GradeClassificationService(IEquivalencyService equivalencyService)
        {
            _equivalencyService = equivalencyService ?? throw new ArgumentNullException(nameof(equivalencyService));
        }

        public string DetermineUKClassification(string overallGradeGPA, string equivalencyNote, string countryOfStudy, string qualificationName)
        {
            if (string.IsNullOrWhiteSpace(equivalencyNote)) return "??";

            string cleanNote = equivalencyNote.Replace("__", " ").Replace("_", " ");
            string noteLower = cleanNote.ToLowerInvariant();

            // 0a. Note-based PhD/Masters shield
            //     If the note says the applicant already holds a PhD or Masters, reflect that directly.
            //     Checked before everything else — no rubric comparison makes sense for a PhD holder.
            if (Regex.IsMatch(cleanNote,
                @"\bgraduated\s+(?:in\s+\d{4}\s+)?with\s+a\s+ph\.?d\b|awarded\s+a\s+ph\.?d\b|has\s+a\s+ph\.?d\b|completed\s+a\s+ph\.?d\b",
                RegexOptions.IgnoreCase))
                return "PhD";

            if (Regex.IsMatch(cleanNote,
                @"\bgraduated\s+(?:in\s+\d{4}\s+)?with\s+a\s+masters?\b|awarded\s+a\s+masters?\s+degree|has\s+a\s+masters?\s+degree|completed\s+a\s+masters?\s+degree",
                RegexOptions.IgnoreCase))
                return "Masters";

            // 0b. QualificationName Masters Shield
            if (qualificationName?.Contains("Masters", StringComparison.OrdinalIgnoreCase) == true)
            {
                if (!DoesNoteLookLikeUndergrad(cleanNote)) return "Masters";
            }

            bool isLancasterGlasgow = noteLower.Contains("lancaster") || noteLower.Contains("glasgow");

            // 1. UK INSTITUTION EARLY PATH
            //    Fires when: (a) CountryOfStudy is UK/Ireland, (b) the note explicitly states
            //    the grade is on the UK scale ("current UK grade of X%"), or (c) the note names
            //    a specific UK institution as the degree-awarding partner (joint programmes like
            //    XJTLU/Liverpool, Nottingham Ningbo, etc.).
            //    "ucl"/"kcl" alone excluded — they appear in "non-prestigious UCL" comparisons.
            //    "university of" excluded — too broad (University of Florida, etc.).
            //    Negative lookbehind (?<!non[-\s]) prevents "non-Liverpool" etc. firing.
            //    Lancaster/Glasgow are excluded — they use a special point scale (step 4).
            bool noteStatesUKGrade = Regex.IsMatch(noteLower, @"\buk\s+grade\b");
            var ukPartnerKeywords = new[] {
                "liverpool", "nottingham", "ningbo",   // "ningbo" = Univ. of Nottingham Ningbo China (UNNC)
                "coventry", "warwick", "loughborough",
                "manchester", "birmingham", "sheffield", "leeds", "bristol", "exeter",
                "southampton", "leicester", "cardiff", "edinburgh", "dundee", "aberdeen",
                "strathclyde", "heriot-watt", "goldsmiths", "college london", "stirling",
                "st andrews", "newcastle", "durham", "aston", "reading", "sussex",
                "essex", "kent", "bath", "surrey", "york"
            };
            bool noteIndicatesUKPartner = ukPartnerKeywords.Any(k =>
                Regex.IsMatch(noteLower, @"(?<!non[-\s])\b" + Regex.Escape(k) + @"\b"));
            if ((IsUK(countryOfStudy) || noteStatesUKGrade || noteIndicatesUKPartner) && !isLancasterGlasgow)
            {
                double? ukGrade = ExtractStudentGradeFromNarrative(cleanNote);
                if (ukGrade.HasValue && ukGrade.Value >= 35)
                {
                    // Percentage-scale grade at a UK institution → apply UK classification directly
                    double v = ukGrade.Value;
                    if (v >= 70) return "1.0";
                    if (v >= 60) return "2.1";
                    if (v >= 50) return "2.2";
                    if (v >= 40) return "3.0";
                    return "Fail";
                }
                // Grade is on a small scale (GPA/points) or not found — do NOT call keywords here.
                // Threshold labels in the note ("1st: Bachelors @ X") would fire as false keywords.
                // Fall through: the custom threshold step handles GPA-scale grades correctly.
            }

            // 2. CUSTOM THRESHOLD DETECTION (for non-UK institutions)
            //    Must run before generic keyword matching because threshold lines contain the same
            //    words ("1st", "2.1") that would otherwise fire as keywords.
            bool skipKeywords = false;
            double? narrativeGradeForThresholds = ExtractStudentGradeFromNarrative(cleanNote);

            // 2a. Range-format thresholds first ("2.2: 12 - 14", "2.1: 15 - 17", "1st: 18 - 20")
            //     These are scale-specific (e.g. Aberdeen 20-pt) and take priority over "@" thresholds
            //     in the same note so the student's grade is compared on the correct scale.
            var rangeThresholds = ParseRangeThresholdsFromNote(cleanNote);
            if (rangeThresholds.Count >= 2 && narrativeGradeForThresholds.HasValue)
            {
                double maxRange = rangeThresholds.Values.Max();
                // Only use these if the student's grade is plausibly on the same scale
                if (narrativeGradeForThresholds.Value <= maxRange * 1.5)
                    return ApplyCustomThresholds(narrativeGradeForThresholds.Value, rangeThresholds);
            }

            // 2b. "@"-format thresholds ("2.2: Bachelors @ 80%", "1st: Bachelors @ CGPA of 3.7/4.0")
            var customThresholds = ParseCustomThresholdsFromNote(cleanNote);
            if (customThresholds.Count >= 2 && narrativeGradeForThresholds.HasValue)
            {
                // Scale mismatch: thresholds are percentages (>30) but grade is on a small scale (<10)
                bool scaleMismatch = customThresholds.Values.All(v => v > 30) && narrativeGradeForThresholds.Value < 10;
                if (!scaleMismatch)
                    return ApplyCustomThresholds(narrativeGradeForThresholds.Value, customThresholds);
                skipKeywords = true;
            }

            // 3. KEYWORD PRIORITY
            if (!skipKeywords)
            {
                string keywordResult = DetermineClassificationFromTextKeywords(cleanNote);
                if (keywordResult != "??") return keywordResult;
            }

            double? studentGrade = ExtractGradeValue(cleanNote);

            // 4. HARD EXCEPTION: Lancaster / Glasgow point scale (17–20 points)
            if (isLancasterGlasgow)
            {
                double val = studentGrade ?? 0;
                if (val >= 17.5) return "1.0";
                if (val >= 14.5) return "2.1";
                if (val >= 11.5) return "2.2";
                if (val >= 8.5)  return "3.0";
                return "Fail";
            }

            // 5. HARD EXCEPTION: Italy /110 scale
            if (countryOfStudy?.Contains("Italy", StringComparison.OrdinalIgnoreCase) == true || noteLower.Contains("110"))
            {
                var m110 = Regex.Matches(cleanNote, @"\b(10[0-9]|110)\b");
                if (m110.Count > 0)
                {
                    double maxVal = m110.Cast<Match>().Select(m => double.Parse(m.Value)).Max();
                    if (maxVal >= 108) return "1.0";
                    if (maxVal >= 106) return "2.1";
                    if (maxVal >= 101) return "2.2";
                }
                if (studentGrade.HasValue && studentGrade.Value <= 30)
                {
                    double v = studentGrade.Value;
                    if (v >= 28.5) return "1.0";
                    if (v >= 26.5) return "2.1";
                    if (v >= 24.0) return "2.2";
                    return "3.0";
                }
            }

            // 6. GENERAL GRADE EXTRACTION FALLBACK
            if (studentGrade.HasValue)
            {
                double val = studentGrade.Value;

                // UK percentage scale — only when country is definitively UK
                if (IsUK(countryOfStudy) && val >= 35)
                {
                    if (val >= 70) return "1.0";
                    if (val >= 60) return "2.1";
                    if (val >= 50) return "2.2";
                    if (val >= 40) return "3.0";
                    return "Fail";
                }

                var equiv = _equivalencyService.GetEquivalency(countryOfStudy);
                if (equiv != null) return ApplySmartEquivalency(val, equiv);

                return ApplyStandardThresholds(GuessScaleAndNormalize(val));
            }

            return "??";
        }

        private double? ExtractGradeValue(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            text = text.Trim().Replace("'", "").Replace("<", "").Replace(">", "");

            var priorityMatch = _priorityKeywordRegex.Match(text);
            if (priorityMatch.Success && double.TryParse(priorityMatch.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double pExtracted))
                return pExtracted;

            var fraction = _fractionRegex.Match(text);
            if (fraction.Success && double.TryParse(fraction.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
                return num; 

            var percent = _percentRegex.Match(text);
            if (percent.Success && double.TryParse(percent.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double pVal))
                return pVal;

            var generalMatches = _generalGradeRegex.Matches(text);
            if (generalMatches.Count > 0)
            {
                return generalMatches.Cast<Match>()
                    .Select(m => double.TryParse(m.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0)
                    .Where(d => d < 500) // SHIELD: Ignore years
                    .DefaultIfEmpty(0)
                    .Max();
            }
            return null;
        }

        private bool DoesNoteLookLikeUndergrad(string note)
        {
            if (string.IsNullOrWhiteSpace(note)) return false;
            string lower = note.ToLowerInvariant();
            return lower.Contains("2:1") || lower.Contains("2.1") || lower.Contains("2:2") || lower.Contains("2.2") || lower.Contains("1st") || lower.Contains("first class");
        }

        private bool IsUK(string country)
        {
            if (string.IsNullOrWhiteSpace(country)) return false;
            return country.Contains("United Kingdom", StringComparison.OrdinalIgnoreCase) ||
                   country.Contains("England", StringComparison.OrdinalIgnoreCase) ||
                   country.Contains("Scotland", StringComparison.OrdinalIgnoreCase) ||
                   country.Contains("Wales", StringComparison.OrdinalIgnoreCase) ||
                   country.Contains("Ireland", StringComparison.OrdinalIgnoreCase); // same classification scale as UK
        }

        private string DetermineClassificationFromTextKeywords(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "??";
            // REMOVE REQUIRES 2_1: Remove the course requirements so we don't grade the student against the job description
            string lower = text.ToLowerInvariant();
            lower = lower.Replace("requires 2.1", "")
                             .Replace("requires a 2.1", "")
                             .Replace("requirement is 2.1", "")
                             .Replace("requires 1st", "")
                             .Replace("2.1 required", "")
                             .Replace("2.1 req", "")
                             .Replace("borderline 2.1", "")
                             .Replace("borderline 2:1", "");
            // Now look for the ACTUAL grade the student has
            if (Regex.IsMatch(lower, @"\b(1st|1\.0|first class)\b")) return "1.0";
            if (Regex.IsMatch(lower, @"\b(2\.1|2:1|upper second)\b")) return "2.1";
            if (Regex.IsMatch(lower, @"\b(2\.2|2:2|lower second)\b")) return "2.2";
            if (Regex.IsMatch(lower, @"\b(3\.0|3rd|third class)\b")) return "3.0";

            if (lower.Contains("summa cum laude") || lower.Contains("high distinction")) return "1.0";
            if (lower.Contains("magna cum laude") || lower.Contains("distinction")) return "1.0"; 
            if (lower.Contains("cum laude") || lower.Contains("merit")) return "2.1";
            
            return "??";
        }

        private Dictionary<string, double> ParseCustomThresholdsFromNote(string cleanNote)
        {
            var thresholds = new Dictionary<string, double>();
            if (string.IsNullOrWhiteSpace(cleanNote)) return thresholds;

            foreach (Match m in _thresholdLineRegex.Matches(cleanNote))
            {
                string raw = m.Groups[5].Value.Replace(',', '.');
                if (!double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                    continue;
                if      (m.Groups[1].Success) thresholds["1.0"] = val;
                else if (m.Groups[2].Success) thresholds["2.1"] = val;
                else if (m.Groups[3].Success) thresholds["2.2"] = val;
                else if (m.Groups[4].Success) thresholds["3.0"] = val;
            }
            return thresholds;
        }

        // Parses banded-range thresholds: "2.2: 12 - 14", "2.1: 15 - 17", "1st: 18 - 20"
        // Returns the lower bound of each band as the threshold value.
        private Dictionary<string, double> ParseRangeThresholdsFromNote(string cleanNote)
        {
            var thresholds = new Dictionary<string, double>();
            if (string.IsNullOrWhiteSpace(cleanNote)) return thresholds;

            foreach (Match m in _rangeThresholdRegex.Matches(cleanNote))
            {
                // Group 5 = lower bound, Group 6 = upper bound — use lower bound as the cut-off
                string raw = m.Groups[5].Value.Replace(',', '.');
                if (!double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                    continue;
                if      (m.Groups[1].Success) thresholds["1.0"] = val;
                else if (m.Groups[2].Success) thresholds["2.1"] = val;
                else if (m.Groups[3].Success) thresholds["2.2"] = val;
                else if (m.Groups[4].Success) thresholds["3.0"] = val;
            }
            return thresholds;
        }

        private double? ExtractStudentGradeFromNarrative(string cleanNote)
        {
            var m = _studentGradeNarrativeRegex.Match(cleanNote);
            if (!m.Success) return null;
            string raw = (m.Groups[1].Success ? m.Groups[1].Value
                        : m.Groups[2].Success ? m.Groups[2].Value
                        : m.Groups[3].Value).Replace(',', '.');
            return double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ? val : (double?)null;
        }

        private string ApplyCustomThresholds(double grade, Dictionary<string, double> thresholds)
        {
            // Detect descending scale (lower number = better, e.g. German 1–5 system)
            bool descending = thresholds.ContainsKey("1.0") && thresholds.ContainsKey("2.2")
                              && thresholds["1.0"] < thresholds["2.2"];

            if (descending)
            {
                if (thresholds.ContainsKey("1.0") && grade <= thresholds["1.0"]) return "1.0";
                if (thresholds.ContainsKey("2.1") && grade <= thresholds["2.1"]) return "2.1";
                if (thresholds.ContainsKey("2.2") && grade <= thresholds["2.2"]) return "2.2";
                return "3.0";
            }
            else
            {
                if (thresholds.ContainsKey("1.0") && grade >= thresholds["1.0"]) return "1.0";
                if (thresholds.ContainsKey("2.1") && grade >= thresholds["2.1"]) return "2.1";
                if (thresholds.ContainsKey("2.2") && grade >= thresholds["2.2"]) return "2.2";
                return "3.0";
            }
        }

        private string ApplySmartEquivalency(double grade, DegreeEquivalency equiv)
        {
            double? t1st = ExtractGradeValue(equiv.First);       
            double? t21  = ExtractGradeValue(equiv.SecondUpper); 
            double? t22  = ExtractGradeValue(equiv.SecondLower); 

            if (!t1st.HasValue || !t22.HasValue) return "??";
            if (t1st.Value < t22.Value)
            {
                if (grade <= t1st.Value) return "1.0";
                if (t21.HasValue && grade <= t21.Value) return "2.1";
                return grade <= t22.Value ? "2.2" : "3.0";
            }
            else
            {
                if (grade >= t1st.Value) return "1.0";
                if (t21.HasValue && grade >= t21.Value) return "2.1";
                return grade >= t22.Value ? "2.2" : "3.0";
            }
        }

        private double GuessScaleAndNormalize(double grade)
        {
            if (grade > 20) return grade;
            if (grade <= 4.0) return (grade / 4.0) * 100.0;
            if (grade <= 5.0) return (grade / 5.0) * 100.0;
            if (grade <= 10.0) return (grade / 10.0) * 100.0;
            return grade * 5.0; 
        }

        private string ApplyStandardThresholds(double percent)
        {
            if (percent >= 70) return "1.0";
            if (percent >= 60) return "2.1";
            return percent >= 50 ? "2.2" : "3.0";
        }

        public string ParseUKGradeText(string t) => DetermineClassificationFromTextKeywords(t);
        public double? ParseGradeValue(string s) => ExtractGradeValue(s);

        // Returns the best grade string for display in OverallGradeGPA.
        // Priority: note is authoritative; raw used only when note yields nothing.
        public string GetPreferredGradeDisplay(string rawGPA, string equivalencyNote)
        {
            if (string.IsNullOrWhiteSpace(equivalencyNote))
                return rawGPA ?? "";

            string cleanNote = equivalencyNote.Replace("__", " ").Replace("_", " ");

            // Extract every student grade mention: "grade of X", "average is/of/= X", "cgpa of X"
            // Group 1 = the number, Group 2 = optional trailing %
            var gradePattern = new Regex(
                @"(?:\bgrade\s+(?:of|is)\s*|\baverage\s*(?:of\s+|is\s+|=\s*)|\bcgpa\s*(?:of|is|:)\s*)(\d+(?:\.\d+)?)(\s*%)?",
                RegexOptions.IgnoreCase);

            var matches = gradePattern.Matches(cleanNote);
            // Note has no extractable grade → fall back to raw (preserves "X/Y" format if present)
            if (matches.Count == 0) return rawGPA ?? "";

            // ── Dual system: multiple grade mentions, last one is GPA-scale (<10) ──────────────
            if (matches.Count >= 2)
            {
                var last = matches[matches.Count - 1];
                if (double.TryParse(last.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double lastVal)
                    && lastVal < 10)
                {
                    var denomMs = _thresholdDenomRegex.Matches(cleanNote);
                    if (denomMs.Count > 0)
                        return $"{last.Groups[1].Value}/{FormatDenomValue(denomMs[denomMs.Count - 1].Groups[1].Value)}";
                    return last.Groups[1].Value;
                }
            }

            // ── Single system: use the last (or only) grade mention from the note ──────────────
            var chosen = matches[matches.Count - 1];
            string gradeStr = chosen.Groups[1].Value;
            bool hasPercent = chosen.Groups[2].Success && chosen.Groups[2].Value.Contains('%');

            if (!double.TryParse(gradeStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double gradeVal))
                return rawGPA ?? "";

            // GPA-scale grade with fraction thresholds → add denominator
            if (gradeVal < 10)
            {
                var denomMs = _thresholdDenomRegex.Matches(cleanNote);
                if (denomMs.Count > 0)
                    return $"{gradeStr}/{FormatDenomValue(denomMs[denomMs.Count - 1].Groups[1].Value)}";
            }

            // Note explicitly included % (e.g. "average is 72.25%") → preserve it
            if (hasPercent) return $"{gradeStr}%";

            // Note has a plain number on percentage scale — return it as-is (no false % appended)
            return string.IsNullOrWhiteSpace(rawGPA) ? gradeStr : gradeStr;
        }

        private string FormatDenomValue(string raw)
        {
            return double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double d) && d == Math.Floor(d)
                ? ((int)d).ToString()
                : raw;
        }
    }
}
