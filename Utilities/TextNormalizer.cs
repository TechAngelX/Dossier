// Utilities/TextNormalizer.cs

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Dossier.Utilities
{
    public static class TextNormalizer
    {
        public static string NormalizeInstitutionName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "";
            
            name = FixEncodingIssues(name);
            name = RemoveDiacritics(name);
            name = name.ToLower();
            name = name.Replace("university of", "").Replace("the ", "");
            name = Regex.Replace(name, @"[^\w\s]", " ");
            name = Regex.Replace(name, @"\s+", " ");
            
            return name.Trim();
        }
        
        public static string FixEncodingIssues(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            
            try
            {
                var encoding = System.Text.Encoding.GetEncoding("ISO-8859-1");
                var utf8 = System.Text.Encoding.UTF8;
                byte[] bytes = encoding.GetBytes(text);
                return utf8.GetString(bytes);
            }
            catch
            {
                return text;
            }
        }
        
        public static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();
            
            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }
            
            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }
        
        public static List<string> ExtractKeyTerms(string normalizedName)
        {
            var commonWords = new HashSet<string> { "of", "the", "and", "in", "at", "for", "on", "a", "an" };
            
            return normalizedName
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(word => word.Length > 2 && !commonWords.Contains(word))
                .ToList();
        }
    }
}
