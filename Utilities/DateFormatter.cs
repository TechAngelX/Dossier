// Utilities/DateFormatter.cs

using System;
using System.Globalization;

namespace Dossier.Utilities
{
    public static class DateFormatter
    {
        private static readonly string[] DateFormats = {
            "dd-MMM-yy", "dd-MMM-yyyy",
            "dd/MM/yy", "dd/MM/yyyy",
            "yyyy-MM-dd", "MM/dd/yyyy"
        };
        
        public static string FormatDate(string dateStr)
        {
            if (string.IsNullOrWhiteSpace(dateStr)) 
                return "";
            
            foreach (var format in DateFormats)
            {
                if (DateTime.TryParseExact(dateStr, format, CultureInfo.InvariantCulture, 
                    DateTimeStyles.None, out DateTime date))
                {
                    return date.ToString("dd/MM/yyyy");
                }
            }
            
            if (DateTime.TryParse(dateStr, out DateTime generalDate))
            {
                return generalDate.ToString("dd/MM/yyyy");
            }
            
            return dateStr;
        }
        
        public static string CalculateDueDate(string receivedDateStr)
        {
            if (string.IsNullOrWhiteSpace(receivedDateStr)) 
                return "";
            
            foreach (var format in DateFormats)
            {
                if (DateTime.TryParseExact(receivedDateStr, format, CultureInfo.InvariantCulture, 
                    DateTimeStyles.None, out DateTime date))
                {
                    return date.AddDays(56).ToString("dd/MM/yyyy");
                }
            }
            
            if (DateTime.TryParse(receivedDateStr, out DateTime generalDate))
            {
                return generalDate.AddDays(56).ToString("dd/MM/yyyy");
            }
            
            return "";
        }
    }
}

