namespace DocumentGenerationApi.Models
{
    public class Clause
    {
        public int Id { get; set; }
        public string Category { get; set; }
        public string Text { get; set; }
        public bool Selected { get; set; }

        public string label { get; set; }
    }
} 