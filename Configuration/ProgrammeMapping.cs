// Configuration/ProgrammeMapping.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Dossier.Configuration
{
    // Single source of truth for programme code mapping across every tab (ADMerger,
    // PAT, Portico automation, …). The app is CLOSED for modification but OPEN for
    // extension: to add a programme or fix a changed Portico code, a user edits the
    // repo's data/progcodes.csv — no code change, no rebuild.
    //
    // data/progcodes.csv ships with the repo and is copied next to the executable on
    // build (see the <Content Include="data/progcodes.csv"> entry in Dossier.csproj).
    // It is found at runtime first beside the exe, then relative to the project root
    // when running from source. Columns: ProgCode,RouteCode,ProgShort,Prog.
    //   ProgCode  — the Portico programme/route code matched during automation (e.g. TMSENG002401)
    //   RouteCode — Portico route code (usually identical to ProgCode)
    //   ProgShort — ADMerger short code (e.g. CSML)
    //   Prog      — full programme name
    //
    // The BuiltIn table below is only a last-resort fallback for when the CSV is
    // missing or unreadable, so the app never breaks.
    public static class ProgrammeMapping
    {
        // Built-in fallback. (Name, Short, ProgCode) — mirrors data/progcodes.csv.
        private static readonly IReadOnlyList<(string Name, string Short, string Portico)> BuiltIn = new[]
        {
            ("MSc Artificial Intelligence for Biomedicine and Healthcare", "AIBH",  "TMSARTSINT03"),
            ("MSc Artificial Intelligence for Sustainable Development",    "AISD",  "TMSARTSSUD01"),
            ("MSc Artificial Intelligence and Data Engineering",          "AIDE",  "TMSCOMSSAD18"),
            ("MSc Information Security",                                   "ISEC",  "TMSCOMSINF01"),
            ("MSc Computational Finance",                                 "CF",    "TMSCOMSCFI01"),
            ("MSc Financial Risk Management",                             "FRM",   "TMSCOMSFRM01"),
            ("MSc Financial Technology",                                  "FT",    "TMSFINSTEC01"),
            ("MSc Emerging Digital Technologies",                         "EDT",   "TMSCOMSEDT01"),
            ("MSc Machine Learning",                                      "ML",    "TMSENG002101"),
            ("MSc Data Science and Machine Learning",                     "DSML",  "TMSENG003001"),
            ("MSc Computational Statistics and Machine Learning",         "CSML",  "TMSENG002401"),
            ("MSc Robotics and Artificial Intelligence",                  "RAI",   "TMSROBAARI01"),
            ("MSc Systems Engineering for the Internet of Things",        "SEIOT", "TMSCOMSEIT01"),
            ("MSc Disability, Design and Innovation",                     "DDI",   "TMSCOMSDDI19"),
            ("MSc Computer Science",                                      "CS",    "TMSCOMSING01"),
            ("MSc Software Systems Engineering",                          "SSE",   "TMSCOMSSSE01"),
            ("MSc Computer Graphics, Vision and Imaging",                 "CGVI",  "TMSCOMSCGV01"),
        };

        // Candidate locations for the editable mapping file: beside the executable
        // (published app / after build) first, then the project's data/ folder when
        // running from source. Same pattern AdMergerView uses.
        public static IEnumerable<string> SearchPaths => new[]
        {
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "progcodes.csv"),
            Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "data", "progcodes.csv")),
        };

        // Effective table: loaded from data/progcodes.csv, or the built-in fallback.
        public static readonly IReadOnlyList<(string Name, string Short, string Portico)> Programmes = LoadTable();

        // Full programme name → ADMerger short code (kept for existing callers; same keys as before).
        public static readonly Dictionary<string, string> Mappings =
            Programmes.Where(p => !string.IsNullOrWhiteSpace(p.Name))
                      .GroupBy(p => p.Name)
                      .ToDictionary(g => g.Key, g => g.Last().Short);

        public static string GetCode(string programmeName) =>
            Mappings.TryGetValue(programmeName, out var c) ? c : programmeName;

        // Full programme name (or ADMerger short code) → Portico programme code
        // (ProgCode, e.g. "TMSCOMSDDI19"). Matching tolerates case, "&"/"and" and
        // punctuation/whitespace differences. Returns null when the programme isn't recognised.
        public static string? GetPorticoCode(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            if (PorticoByNormalisedName.TryGetValue(Normalise(value), out var byName)) return byName;
            if (PorticoByShort.TryGetValue(value.Trim().ToLowerInvariant(), out var byShort)) return byShort;
            return null;
        }

        private static readonly Dictionary<string, string> PorticoByNormalisedName =
            Programmes.Where(p => !string.IsNullOrWhiteSpace(p.Name) && !string.IsNullOrWhiteSpace(p.Portico))
                      .GroupBy(p => Normalise(p.Name))
                      .ToDictionary(g => g.Key, g => g.Last().Portico);

        private static readonly Dictionary<string, string> PorticoByShort =
            Programmes.Where(p => !string.IsNullOrWhiteSpace(p.Short) && !string.IsNullOrWhiteSpace(p.Portico))
                      .GroupBy(p => p.Short.ToLowerInvariant())
                      .ToDictionary(g => g.Key, g => g.Last().Portico);

        private static string Normalise(string name) =>
            Regex.Replace((name ?? "").ToLowerInvariant().Replace("&", "and"), "[^a-z0-9]+", " ").Trim();

        // Read data/progcodes.csv (ProgCode,RouteCode,ProgShort,Prog) into the effective
        // table. Any failure (no file, bad rows) falls back to the built-in table so the
        // app keeps working.
        private static IReadOnlyList<(string Name, string Short, string Portico)> LoadTable()
        {
            try
            {
                foreach (var path in SearchPaths)
                {
                    if (!File.Exists(path)) continue;

                    var rows = new List<(string Name, string Short, string Portico)>();
                    foreach (var line in File.ReadAllLines(path).Skip(1))
                    {
                        if (string.IsNullOrWhiteSpace(line)) continue;
                        var parts = SplitCsvRow(line);
                        if (parts.Length < 4) continue;

                        var progCode  = parts[0].Trim();
                        var progShort = parts[2].Trim();
                        var progName  = parts[3].Trim();
                        if (string.IsNullOrWhiteSpace(progShort) && string.IsNullOrWhiteSpace(progName))
                            continue;

                        rows.Add((progName, progShort, progCode));
                    }

                    if (rows.Count > 0) return rows;
                }
            }
            catch
            {
                // Fall through to the built-in table — never let a bad CSV break the app.
            }

            return BuiltIn;
        }

        // Minimal RFC-4180-ish split: honours double-quoted fields so names containing
        // commas (e.g. "MSc Disability, Design and Innovation") parse correctly.
        private static string[] SplitCsvRow(string line)
        {
            var fields = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;
            foreach (char c in line)
            {
                if (c == '"') { inQuotes = !inQuotes; }
                else if (c == ',' && !inQuotes) { fields.Add(current.ToString()); current.Clear(); }
                else { current.Append(c); }
            }
            fields.Add(current.ToString());
            return fields.ToArray();
        }
    }
}
