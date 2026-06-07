// Services/InstitutionMatchingService.cs

using System;
using System.Collections.Generic;
using System.Linq;
using Dossier.Utilities;

namespace Dossier.Services
{
    public class InstitutionMatchingService : IInstitutionMatchingService
    {
        private const int MinimumMatchThreshold = 70;
        
        public string FindBestMatch(string searchName, List<string> candidateNames)
        {
            if (string.IsNullOrWhiteSpace(searchName))
                return null;
            
            string normalizedSearch = TextNormalizer.NormalizeInstitutionName(searchName);
            
            System.Diagnostics.Debug.WriteLine($"\n=== Matching '{searchName}' ===");
            System.Diagnostics.Debug.WriteLine($"Normalized: '{normalizedSearch}'");
            
            string bestMatch = null;
            int bestScore = 0;
            var topMatches = new List<(string name, int score)>();
            
            foreach (var candidateName in candidateNames)
            {
                string normalizedCandidate = TextNormalizer.NormalizeInstitutionName(candidateName);
                
                int score = CalculateSimilarityScore(normalizedSearch, normalizedCandidate);
                
                if (score >= MinimumMatchThreshold)
                {
                    topMatches.Add((candidateName, score));
                }
                
                if (score > bestScore && score >= MinimumMatchThreshold)
                {
                    bestScore = score;
                    bestMatch = candidateName;
                }
            }
            
            topMatches = topMatches.OrderByDescending(x => x.score).Take(5).ToList();
            System.Diagnostics.Debug.WriteLine($"Top 5 matches:");
            foreach (var match in topMatches)
            {
                System.Diagnostics.Debug.WriteLine($"  {match.score}% - {match.name}");
            }
            System.Diagnostics.Debug.WriteLine($"Best match: {bestMatch} ({bestScore}%)");
            
            return bestMatch;
        }
        
        public int CalculateMatchScore(string search, string candidate, List<string> searchTerms)
        {
            return CalculateSimilarityScore(search, candidate);
        }
        
        private int CalculateSimilarityScore(string search, string candidate)
        {
            if (search == candidate)
                return 100;
            
            if (candidate.Contains(search) || search.Contains(candidate))
                return 95;
            
            var searchWords = search.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(w => w.Length > 2)
                .ToList();
            
            var candidateWords = candidate.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(w => w.Length > 2)
                .ToList();
            
            int exactMatches = 0;
            int fuzzyMatches = 0;
            int totalSignificantWords = 0;
            
            foreach (var searchWord in searchWords)
            {
                if (IsSignificantWord(searchWord))
                {
                    totalSignificantWords++;
                    bool foundMatch = false;
                    
                    foreach (var candidateWord in candidateWords)
                    {
                        if (searchWord == candidateWord)
                        {
                            exactMatches++;
                            foundMatch = true;
                            break;
                        }
                    }
                    
                    if (!foundMatch)
                    {
                        foreach (var candidateWord in candidateWords)
                        {
                            int distance = LevenshteinDistance(searchWord, candidateWord);
                            int maxLen = Math.Max(searchWord.Length, candidateWord.Length);
                            double similarity = 1.0 - ((double)distance / maxLen);
                            
                            if (similarity >= 0.90)
                            {
                                fuzzyMatches++;
                                break;
                            }
                        }
                    }
                }
            }
            
            if (totalSignificantWords == 0)
                return 0;
            
            int exactScore = (exactMatches * 100) / totalSignificantWords;
            int fuzzyScore = (fuzzyMatches * 100) / totalSignificantWords;
            
            return exactScore + (fuzzyScore / 2);
        }
        
        private bool IsSignificantWord(string word)
        {
            var insignificantWords = new HashSet<string>
            {
                "university", "college", "institute", "school", "academy",
                "national", "state", "royal", "imperial", "public",
                "science", "technology", "arts"
            };
            
            return !insignificantWords.Contains(word) && word.Length > 2;
        }
        
        private int LevenshteinDistance(string source, string target)
        {
            if (string.IsNullOrEmpty(source))
                return string.IsNullOrEmpty(target) ? 0 : target.Length;
            
            if (string.IsNullOrEmpty(target))
                return source.Length;
            
            int n = source.Length;
            int m = target.Length;
            int[,] d = new int[n + 1, m + 1];
            
            for (int i = 0; i <= n; i++)
                d[i, 0] = i;
            
            for (int j = 0; j <= m; j++)
                d[0, j] = j;
            
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (target[j - 1] == source[i - 1]) ? 0 : 1;
                    
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            
            return d[n, m];
        }
    }
}

