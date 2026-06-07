// Services/IRankingService.cs

using System.Collections.Generic;
using System.Threading.Tasks;

namespace Dossier.Services
{
    public interface IRankingService
    {
        Task LoadRankingsAsync();
        string GetRanking(string institutionName);
        int Count { get; }
        IReadOnlyList<string> GetAllInstitutionNames();
    }
}
