using System.ComponentModel.DataAnnotations;

namespace ECSGDocumentGenerator
{
    public class Hit
    {
        public string PolicyAreas { get; set; }
        public string Id { get; set; }
        public string SensitiveSummaryDate { get; set; }
        public string Knowledge { get; set; }
        public string CaseSource { get; set; }
        public string ContactPerson { get; set; }
        public string Outline { get; set; }
        public string LastAdoptedProposalDecision { get; set; }
        public string RefId { get; set; }
        public string ReasonForSensitivity { get; set; }
        public string LineToTake { get; set; }
        public bool EcFunding { get; set; }
        public string AuthorOfTheSensitiveSummary { get; set; }
        public string LastAdoptedDecisionDate { get; set; }
        //[Key]
        public string MemberState { get; set; }
        public string PolicyContext { get; set; }
        public string LegalAssessment { get; set; }
        public LeadDG LeadDg { get; set; }
        public string[] MemberStates { get; set; }
        public string CaseTitle { get; set; }
        public string Reason { get; set; }
        public string CaseType { get; set; }
        public bool CaseSensitivity { get; set; }
        public bool CaseSensitivitySg { get; set; }
        public bool CaseSensitivityLs { get; set; }
        public LegalBasis[] LegalBasis { get; set; }
        public string CreationDate { get; set; }
        public bool Article259 { get; set; }
        public bool Disabled { get; set; }
        public string IncriminatedFact { get; set; }
    }

}
