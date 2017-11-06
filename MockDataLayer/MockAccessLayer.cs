using System;
using System.Collections.Generic;
using MockDataLayer.Entities;

namespace MockDataLayer
{
    public static class MockAccessLayer
    {
        public static Inspection GetInspectionForProperty()
        {
            return new Inspection
            {
                DueDate = DateTime.Now,
                PropertyReference = "abcde",
                CompletionDate = DateTime.Now,
                InspectorContact = "000",
                InspectorName = "John Rudd",
                Appointments = new List<InspectionAppointment>
                {
                    new InspectionAppointment
                    {
                        ScheduledDate = DateTime.Now.AddDays(1),
                        Infringments = new List<InspectionInfringment>
                        {
                            new InspectionInfringment
                            {
                                CorrectiveAction = "Corrective Action goes here",
                                Location = "Bathroom",
                                DueDate = DateTime.Now.AddDays(4)
                            },
                            new InspectionInfringment
                            {
                                CorrectiveAction = "Corrective Action goes here",
                                Location = "Kitchen",
                                DueDate = DateTime.Now.AddDays(4)
                            }
                        }
                    }
                }
            };
        }
    }
}