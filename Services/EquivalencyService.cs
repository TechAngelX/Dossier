// Services/EquivalencyService.cs

using System;
using System.Collections.Generic;
using Dossier.Models;
using Dossier.data;

namespace Dossier.Services
{
    public class EquivalencyService : IEquivalencyService
    {
        public int Count => EquivalencyData.Equivalencies.Count;
        
        public void LoadEquivalencies()
        {
            // Data is static and loaded automatically by the compiler.
            // Method remains for interface compatibility.
        }
        
        public DegreeEquivalency GetEquivalency(string country)
        {
            if (string.IsNullOrWhiteSpace(country))
                return null;

            string key = country.Trim();
            
            if (EquivalencyData.Equivalencies.TryGetValue(key, out var data))
            {
                return new DegreeEquivalency
                {
                    Country = key,
                    Third = data.G30?.TrimStart('<'),
                    SecondLower = data.G22,
                    SecondUpper = data.G21,
                    First = data.G10
                };
            }
            
            return null;
        }
        
        public Dictionary<string, DegreeEquivalency> GetAllEquivalencies()
        {
            var all = new Dictionary<string, DegreeEquivalency>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in EquivalencyData.Equivalencies)
            {
                all[kvp.Key] = new DegreeEquivalency
                {
                    Country = kvp.Key,
                    Third = kvp.Value.G30?.TrimStart('<'),
                    SecondLower = kvp.Value.G22,
                    SecondUpper = kvp.Value.G21,
                    First = kvp.Value.G10
                };
            }
            return all;
        }
    }
}
