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
        public string Username { get; set; }
        [JsonProperty]
        public string Password { get; set; }
        [JsonProperty]
        public string From { get; set; }

        public Mail()
        {
            STMP = "smtp.gmail.com";
            Port = 587;
            SSL = true;
            Username = "";
            Password = "";
            From = "example@email.com";
        }
    }
}
