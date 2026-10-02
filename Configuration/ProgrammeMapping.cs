// Configuration/ProgrammeMapping.cs

using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Dossier.Configuration
{
    public static class ProgrammeMapping
    {
        // Single source of truth for every tab (ADMerger, PAT, Portico, …): one row per
        // UCL CS-department programme. Name = full programme name, Short = ADMerger short
        // code, Portico = Portico programme/route code.
        public static readonly IReadOnlyList<(string Name, string Short, string Portico)> Programmes = new[]
        {
            ("MSc Artificial Intelligence for Biomedicine and Healthcare", "AIBH",  "TMSARTSBAH01"),
            ("MSc Artificial Intelligence for Sustainable Development",    "AISD",  "TMSARTSSUD01"),
            ("MSc Artificial Intelligence and Data Engineering",          "AIDE",  "TMSCOMSSAD18"),
            ("MSc Information Security",                                   "ISEC",  "TMSCOMSINF01"),
            ("MSc Computational Finance",                                 "CF",    "TMSCOMSCFI01"),
            ("MSc Financial Risk Management",                             "FRM",   "TMSCOMSFRM01"),
            ("MSc Financial Technology",                                  "FT",    "TMSFINSTEC01"),
            ("MSc Emerging Digital Technologies",                         "EDT",   "TMSCOMSEDT01"),
            ("MSc Machine Learning",                                      "ML",    "TMSCOMSMCL01"),
            ("MSc Data Science and Machine Learning",                     "DSML",  "TMSDATSMLE01"),
            ("MSc Computational Statistics and Machine Learning",         "CSML",  "TMSCOMSSML01"),
            ("MSc Robotics and Artificial Intelligence",                  "RAI",   "TMSROBAARI01"),
            ("MSc Systems Engineering for the Internet of Things",        "SEIOT", "TMSCOMSEIT01"),
            ("MSc Disability, Design and Innovation",                     "DDI",   "TMSCOMSDDI19"),
            ("MSc Computer Science",                                      "CS",    "TMSCOMSING01"),
            ("MSc Software Systems Engineering",                          "SSE",   "TMSCOMSSSE01"),
            ("MSc Computer Graphics, Vision and Imaging",                 "CGVI",  "TMSCOMSCGV01"),
        };

        // Full programme name → ADMerger short code (kept for existing callers; same keys as before).
        public static readonly Dictionary<string, string> Mappings =
            Programmes.ToDictionary(p => p.Name, p => p.Short);

        public static string GetCode(string programmeName) =>
            Mappings.TryGetValue(programmeName, out var c) ? c : programmeName;

        // Full programme name (or ADMerger short code) → Portico programme/route code
        // (e.g. "TMSCOMSDDI19"). Matching tolerates case, "&"/"and" and punctuation/whitespace
        // differences. Returns null when the programme isn't recognised.
        public static string? GetPorticoCode(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            if (PorticoByNormalisedName.TryGetValue(Normalise(value), out var byName)) return byName;
            if (PorticoByShort.TryGetValue(value.Trim().ToLowerInvariant(), out var byShort)) return byShort;
            return null;
        }

        private static readonly Dictionary<string, string> PorticoByNormalisedName =
            Programmes.ToDictionary(p => Normalise(p.Name), p => p.Portico);

        private static readonly Dictionary<string, string> PorticoByShort =
            Programmes.ToDictionary(p => p.Short.ToLowerInvariant(), p => p.Portico);

        private static string Normalise(string name) =>
            Regex.Replace((name ?? "").ToLowerInvariant().Replace("&", "and"), "[^a-z0-9]+", " ").Trim();
    }
}
