// Models/ApplicationRecord.cs

using CsvHelper.Configuration.Attributes;

namespace Dossier.Models
{
    public class ApplicationRecord
    {
        [Name("Applicant ID")]
        public string? ApplicantID { get; set; }
       
        public string? Programme { get; set; }
        public string? Forename { get; set; }
        public string? Surname { get; set; }
        
        [Name("Fee Status")]
        public string? FeeStatus { get; set; }
        
        [Name("Qualification name")]
        public string? QualificationName { get; set; }
        
        [Name("Degree subject")]
        public string? DegreeSubject { get; set; }
        
        [Name("Institution name")]
        public string? InstitutionName { get; set; }
        
        [Name("Country of study")]
        public string? CountryOfStudy { get; set; }
        
        [Name("Overall  grade/GPA")]
        public string? OverallGradeGPA { get; set; }
       
        [Name("Equivalency note")]
        public string? EquivalencyNote { get; set; }
        
        [Name("Grade Achieved/Pending")]
        public string? GradeAchievedPending { get; set; }

        // Optional fields (user-selectable via checkboxes)
        [Name("Gender")]
        public string? Gender { get; set; }

        [Name("Country of Nationality")]
        public string? Nationality { get; set; }

        [Name("Date of Birth")]
        public string? DateOfBirth { get; set; }

        [Name("Email address")]
        public string? Email { get; set; }

        [Name("Paid")]
        public string? Paid { get; set; }
    }
}
