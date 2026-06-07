// Services/IInstitutionMatchingService.cs

using System.Collections.Generic;

namespace Dossier.Services
{
    public interface IInstitutionMatchingService
    {
        string FindBestMatch(string searchName, List<string> candidateNames);
        int CalculateMatchScore(string searchName, string candidateName, List<string> searchTerms);
    }
}
