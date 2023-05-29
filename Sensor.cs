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
        public bool Skip { get; set; } = false;
        [JsonIgnore]
        public TimeOnly Time { get; set; } = TimeOnly.FromDateTime(DateTime.Now);
        [JsonIgnore]
        public DateOnly Date { get; set; } = DateOnly.FromDateTime(DateTime.Now);
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
