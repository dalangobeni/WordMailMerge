namespace MockDataLayer.Entities
{
    public abstract class BaseAddressesEntity
    {
        public string SiteNumber { get; set; }
        public string SiteStreetName { get; set; }
        public string SiteLocality { get; set; }
        public string SiteCity { get; set; }
        public long? SiteCountyId { get; set; }
        public string SitePostCode { get; set; }
        public string MailNumber { get; set; }
        public string MailStreetName { get; set; }
        public string MailLocality { get; set; }
        public string MailCity { get; set; }
        public long? MailCountyId { get; set; }
        public string MailPostCode { get; set; }

    }
}