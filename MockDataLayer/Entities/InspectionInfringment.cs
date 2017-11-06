using System;

namespace MockDataLayer.Entities
{
    public class InspectionInfringment
    {
        public DateTime DueDate { get; set; }
        public DateTime? CompletionDate { get; set; }
        public string Location { get; set; }
        public string CorrectiveAction { get; set; }
    }
}