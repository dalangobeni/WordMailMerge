using System;
using System.Collections.Generic;

namespace MockDataLayer.Entities
{
    public class InspectionAppointment
    {
        public DateTime ScheduledDate { get; set; }

        public virtual ICollection<InspectionInfringment> Infringments { get; set; }
    }
}