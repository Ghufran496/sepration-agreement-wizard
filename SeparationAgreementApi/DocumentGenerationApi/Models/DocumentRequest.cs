using System.Collections.Generic;

namespace DocumentGenerationApi.Models
{
    public class DocumentRequest
    {
        public PartyInfo PartyInfo { get; set; }
        public List<Clause> SelectedClauses { get; set; } = new List<Clause>();
    }
} 