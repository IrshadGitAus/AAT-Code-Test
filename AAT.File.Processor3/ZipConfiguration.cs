namespace AAT.File.Processor3
{
    public class ZipConfiguration
    {
        public string InboundFolder { get; set; }

        public string ExtractFolder { get; set; }

        public string AdministratorEmail { get; set; }

        public string[] AllowedFileTypes { get; set; }

        public string[] AllowedImageFileTypes { get; set; }

        public string[] AllowedXMLFiles { get; set; }

        public string[] XMLPartyAttributes { get; set; }

        public string XMLPartyAttributeToCreateFolder { get; set; }

    }
}
