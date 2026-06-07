using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using Dossier.data;

namespace Dossier.Services
{
    public class RankingService : IRankingService
    {
        private readonly IInstitutionMatchingService _matchingService;
        private List<string> _institutionNames = new();

        public int Count => RankingData.Rankings.Count;

        public RankingService(IInstitutionMatchingService matchingService)
        {
            _matchingService = matchingService;
        }

        public async Task LoadRankingsAsync()
        {
            await Task.Run(() => {
                _institutionNames = RankingData.Rankings.Keys.ToList();
            });
        }

        public string GetRanking(string institutionName)
        {
            if (string.IsNullOrWhiteSpace(institutionName)) return "NR";

            // 1. Apply Mappings first (e.g., convert "UCL" to "University College London")
            string normalizedName = institutionName.Trim();
            if (MappingData.InstitutionMappings.TryGetValue(normalizedName, out string? mappedName))
            {
                normalizedName = mappedName;
            }
            
            // 2. Check for exact match in the rankings
            if (RankingData.Rankings.TryGetValue(normalizedName, out string? rank))
                return rank;

            // 3. Check for fuzzy match
            string? bestMatch = _matchingService.FindBestMatch(normalizedName, _institutionNames);
            return bestMatch != null ? RankingData.Rankings[bestMatch] : "NR";
        }

        public IReadOnlyList<string> GetAllInstitutionNames() => _institutionNames.AsReadOnly();
    }
}
