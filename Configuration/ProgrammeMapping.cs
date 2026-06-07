// Configuration/ProgrammeMapping.cs

using System.Collections.Generic;

namespace Dossier.Configuration
{
    public static class ProgrammeMapping
    {
        public static readonly Dictionary<string, string> Mappings = new Dictionary<string, string>
        {
            {"MSc Artificial Intelligence for Biomedicine and Healthcare", "AIBH"},
            {"MSc Artificial Intelligence for Sustainable Development", "AISD"},
            {"MSc Artificial Intelligence and Data Engineering", "AIDE"},
            {"MSc Information Security", "ISEC"},
            {"MSc Computational Finance", "CF"},
            {"MSc Financial Risk Management", "FRM"},
            {"MSc Financial Technology", "FT"},
            {"MSc Emerging Digital Technologies", "EDT"},
            {"MSc Machine Learning", "ML"},
            {"MSc Data Science and Machine Learning", "DSML"},
            {"MSc Computational Statistics and Machine Learning", "CSML"},
            {"MSc Robotics and Artificial Intelligence", "RAI"},
            {"MSc Systems Engineering for the Internet of Things", "SEIOT"},
            {"MSc Disability, Design and Innovation", "DDI"},
            {"MSc Computer Science", "CS"},
            {"MSc Software Systems Engineering", "SSE"},
            {"MSc Computer Graphics, Vision and Imaging", "CGVI"}
        };
        
        public static string GetCode(string programmeName)
        {
            return Mappings.ContainsKey(programmeName) ? Mappings[programmeName] : programmeName;
        }
    }
}
