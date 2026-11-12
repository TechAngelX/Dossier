using System;

namespace Dossier.Utilities;

public static class AcademicYearHelper
{
    // Returns the current UCL admissions cycle year as "YYYY/YY" (e.g. "2027/28").
    // The cycle for the following September rolls over around August of the prior year:
    // from Aug onwards the active cycle targets next September's entry.
    public static string Current(DateTime? today = null)
    {
        var d = today ?? DateTime.Today;
        int entryYear = d.Month >= 8 ? d.Year + 1 : d.Year;
        return $"{entryYear}/{(entryYear + 1) % 100:D2}";
    }
}
