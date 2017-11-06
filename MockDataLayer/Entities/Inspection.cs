using System;
using System.Collections.Generic;

namespace MockDataLayer.Entities
{
    public class Inspection
    {
        public DateTime DueDate { get; set; }
        public DateTime? CompletionDate { get; set; }
        public string WaivedBy { get; set; }
        public DateTime? WaivedDate { get; set; }
        public int? WSCCaseReference { get; set; }
        public string PropertyReference { get; set; }
        public  ICollection<InspectionAppointment> Appointments { get; set; }
        public string InspectorName { get; set; }
        public string InspectorContact { get; set; }

    }
}