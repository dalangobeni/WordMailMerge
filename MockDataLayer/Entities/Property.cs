using System.Collections.Generic;

namespace MockDataLayer.Entities
{
    public class Property : BaseAddressesEntity
    {
        public long PropRef { get; set; }
        
        public virtual ICollection<Inspection> Inspections { get; set; }
      
    }
}