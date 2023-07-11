using Newtonsoft.Json;
using static System.Environment;

namespace OSWMonitorService.JSON
{
    public class Config
    {
        [JsonProperty]
        public DataType DataType { get; set; }
        [JsonProperty]
        public int Delay { get; set; }
        [JsonProperty]
        public Mail Email { get; set; }

        [JsonIgnore]
        public static string PATH = Path.Combine(GetFolderPath(SpecialFolder.CommonApplicationData), "OSW Monitoring");
        [JsonIgnore]
        private static string CONFIG = Path.Combine(PATH, "config.json");

        public Config()
        {
            DataType = new DataType();
            Delay = 10;
            Email = new Mail();
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
                config = new Config();

                File.WriteAllText(CONFIG, JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            else
            {
                config = JsonConvert.DeserializeObject<Config>(File.ReadAllText(CONFIG));
            }

            if(!Directory.Exists(Path.Combine(PATH, "Sensors")))
            {
                Directory.CreateDirectory(Path.Combine(PATH, "Sensors"));

                List<string> recipients = new List<string>
                {
                    "example@mail.com"
                };

                Sensor sensor = new Sensor("Example", "192.168.1.1");
                sensor.Recipients = recipients;
                sensor.Skip = true;

                sensor.Save();
            }

            return config;
        }

        public void Save()
        {
            File.WriteAllText(CONFIG, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }
}
