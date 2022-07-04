using DocumentFormat.OpenXml.Wordprocessing;

namespace SharedLibrary
{
    public class Request
    {
        public string WordOpenXML { get; set; }
        public string? InnerXml { get; set; }
        public string? OuterXml { get; set; }
    }
}
