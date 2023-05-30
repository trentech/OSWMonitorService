using Newtonsoft.Json;
using System.Reflection.PortableExecutable;
using static System.Environment;

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
        public double Temperature { get; set; }
        [JsonIgnore]
        public double Humidity { get; set; }
        [JsonIgnore]
        public double DewPoint { get; set; }
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
