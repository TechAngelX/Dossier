// Services/ICsvService.cs

using System.Collections.Generic;
using Dossier.Models;

namespace Dossier.Services
{
    public interface ICsvService
    {
        List<InTrayRecord> LoadInTrayRecords(string filePath);
        List<ApplicationRecord> LoadApplicationRecords(string filePath);
List<string> GenerateOutputFiles(List<OutputRecord> data, string outputFolderPath, OutputSettings? settings = null);
    }
}
