// Services/IEquivalencyService.cs

using System.Collections.Generic;
using Dossier.Models;

namespace Dossier.Services
{
    public interface IEquivalencyService
    {
        void LoadEquivalencies();
        DegreeEquivalency GetEquivalency(string country);
        Dictionary<string, DegreeEquivalency> GetAllEquivalencies();
        int Count { get; }
    }
}
