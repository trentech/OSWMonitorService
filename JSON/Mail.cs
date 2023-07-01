using Newtonsoft.Json;

namespace OSWMonitorService.JSON
{
    public class Mail
    {
        [JsonProperty]
        public string STMP { get; set; }
        [JsonProperty]
        public int Port { get; set; }
        [JsonProperty]
        public bool SSL { get; set; }
        [JsonProperty]
        public string From { get; set; }

        public Mail()
        {
            STMP = "smtp.gmail.com";
            Port = 25;
            SSL = false;
            From = "example@email.com";
        }
    }
}
