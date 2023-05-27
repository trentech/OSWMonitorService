using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Environment;

namespace OSWMontiorService
{

    public class Config
    {
        [JsonProperty("Sensors")]
        public List<Sensor> Sensors { get; set; }

        [JsonProperty("Path")]
        public string Destination { get; set; }
        [JsonProperty("Delay")]
        public int Delay { get; set; }
        [JsonIgnore]
        public static string PATH = Path.Combine(GetFolderPath(SpecialFolder.CommonApplicationData), "OSW Monitoring");
        [JsonIgnore]
        private static string CONFIG = Path.Combine(PATH, "config.json");

        public Config()
        {
            Sensors = new List<Sensor>();
            Destination = PATH;
            Delay = 10;
        }

        public static Config Get()
        {
            Config config;

            if (!Directory.Exists(PATH))
            {
                Directory.CreateDirectory(PATH);
            }

            if (!File.Exists(CONFIG))
            {
                Sensor example = new Sensor("Example Sensor", "192.168.1.1");
                example.Skip = true;

                config = new Config();
                config.Sensors.Add(example);
                config.Destination = PATH;

                File.WriteAllText(CONFIG, JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            else
            {
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                config = JsonConvert.DeserializeObject<Config>(File.ReadAllText(CONFIG));
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
            }

#pragma warning disable CS8603 // Possible null reference return.
            return config;
#pragma warning restore CS8603 // Possible null reference return.
        }

        public void Save()
        {
            File.WriteAllText(CONFIG, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }
}
