using System;
using System.Collections.Generic;

namespace DocumentGenerationApi.Models
{
    public class PartyInfo
    {
        public string Party1FirstName { get; set; }
        public string Party1MiddleName { get; set; }
        public string Party1LastName { get; set; }
        public string Party2FirstName { get; set; }
        public string Party2MiddleName { get; set; }
        public string Party2LastName { get; set; }
        
        public string SpousalSupportPayor { get; set; }
        public string SpousalSupportRecipient { get; set; }
        public string ChildSupportPayor { get; set; }
        public string ChildSupportRecipient { get; set; }
        
        public string Party1Role { get; set; }
        public string Party2Role { get; set; }
        
        public DateTime? MarriedDate { get; set; }
        public DateTime? CohabitationDate { get; set; }
        public DateTime? SeparationDate { get; set; }
        
        public List<Child> Children { get; set; } = new List<Child>();
    }

    public class Child
    {
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public DateTime? Birthdate { get; set; }
    }
} 