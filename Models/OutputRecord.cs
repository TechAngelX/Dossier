// Models/OutputRecord.cs

namespace Dossier.Models
{
    public class OutputRecord
    {
        public string? ReceivedDate { get; set; }
        public string? DueDate { get; set; }
        public string? StudentNo { get; set; }
        public string? Programme { get; set; }
        public string? Forename { get; set; }
        public string? Surname { get; set; }
        public string? FeeStatus { get; set; }
        public string? CountryOfStudy { get; set; }
        public string? QualificationName { get; set; }
        public string? DegreeSubject { get; set; }
        public string? InstitutionName { get; set; }
        public string? THERanking { get; set; }
        public string? OverallGradeGPA { get; set; }
        public string? DegreeStatus { get; set; } 
        public string? EquivalencyNote { get; set; }
        public string? UKGrade { get; set; }

        // Optional fields (user-selectable via checkboxes)
        public string? Gender { get; set; }
        public string? Nationality { get; set; }
        public string? DateOfBirth { get; set; }
        public string? Email { get; set; }
        public string? Paid { get; set; }
    }
}
