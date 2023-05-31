using Newtonsoft.Json;

namespace OSWMonitorService
{
    public class Sensor
    {
        [JsonProperty]
        public string Name { get; }
        [JsonProperty]
        public string IP { get; set; }
        [JsonProperty]
        public bool Skip { get; set; } = false;
        [JsonIgnore]
        public DateTime DateTime { get; set; } = DateTime.Now;
        [JsonIgnore]
        public double Temperature { get; set; } = 0;
        [JsonIgnore]
        public double Humidity { get; set; } = 0;
        [JsonIgnore]
        public double DewPoint { get; set; } = 0;
        [JsonIgnore]
        public bool IsRecording { get; set; } = false;
        [JsonIgnore]
        public bool IsOffline { get; set; } = false;
        public Sensor(string name, string ip)
        {
            Name = name;
            IP = ip;
        }
    }
}
