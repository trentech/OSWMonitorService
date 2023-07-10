using Newtonsoft.Json;
using static System.Environment;

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
        public double HumidityWarning { get; set; } = 0;
        [JsonProperty]
        public double HumidityLimit { get; set; } = 0;
        [JsonProperty]
        public List<string> Recipients { get; set; }
        [JsonProperty]
        public bool IsOnline { get; set; } = true;
        [JsonIgnore]
        public DateTime DateTime { get; set; } = DateTime.Now;
        [JsonIgnore]
        public double Temperature { get; set; } = 0;
        [JsonIgnore]
        public double Humidity { get; set; } = 0;
        [JsonIgnore]
        public double DewPoint { get; set; } = 0;

        public Sensor(string name, string ip)
        {
            Name = name;
            IP = ip;
            Recipients = new List<string>();
        }

        public void Save()
        {
            File.WriteAllText(Path.Combine(Config.PATH, @"Sensors\" + IP + ".json"), JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }
}
