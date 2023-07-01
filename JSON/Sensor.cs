using Newtonsoft.Json;

namespace OSWMonitorService.JSON
{
    public class Sensor
    {
        [JsonProperty]
        public string Name { get; }
        [JsonProperty]
        public string IP { get; set; }
        [JsonProperty]
        public bool Skip { get; set; } = false;
        [JsonProperty]
        public double TemperatureLimit { get; set; } = 0;
        [JsonProperty]
        public double HumidityLimit { get; set; } = 0;
        [JsonProperty]
        public List<string> Recipients { get; set; }
        [JsonIgnore]
        public DateTime DateTime { get; set; } = DateTime.Now;
        [JsonIgnore]
        public double Temperature { get; set; } = 0;
        [JsonIgnore]
        public double Humidity { get; set; } = 0;
        [JsonIgnore]
        public double DewPoint { get; set; } = 0;
        [JsonIgnore]
        public bool IsOnline { get; set; } = true;
        public Sensor(string name, string ip)
        {
            Name = name;
            IP = ip;
            Recipients = new List<string>();
        }
    }
}
