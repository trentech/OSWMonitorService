using Newtonsoft.Json;
using System.Reflection.PortableExecutable;
using static System.Environment;

namespace OSWMontiorService
{
    public class Sensor
    {
        [JsonProperty]
        public string Name { get; }
        [JsonProperty]
        public string IP { get; set; }
        [JsonProperty]
        public bool Skip { get; set; } = true;
        [JsonIgnore]
        public TimeOnly Time { get; set; }
        [JsonIgnore]
        public DateOnly Date { get; set; }
        [JsonIgnore]
        public double Temperature { get; set; }
        [JsonIgnore]
        public double Humidity { get; set; }
        [JsonIgnore]
        public double DewPoint { get; set; }
        [JsonIgnore]
        public bool IsRecording { get; set; }
        [JsonIgnore]
        public bool IsOffline { get; set; } = false;
        public Sensor(string name, string ip)
        {
            Name = name;
            IP = ip;
        }
    }
}
