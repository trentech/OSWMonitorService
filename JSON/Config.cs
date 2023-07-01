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
        [JsonProperty]
        public List<Sensor> Sensors { get; set; }
        [JsonProperty]
        public bool DevMode { get; set; }

        [JsonIgnore]
        public static string PATH = Path.Combine(GetFolderPath(SpecialFolder.CommonApplicationData), "OSW Monitoring");
        [JsonIgnore]
        private static string CONFIG = Path.Combine(PATH, "config.json");

        public Config()
        {
            Sensors = new List<Sensor>();
            DataType = new DataType();
            Delay = 10;
            Email = new Mail();
            DevMode = false;
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
                example.Recipients.Add("test1@example.com");
                example.Recipients.Add("test2@example.com");

                config = new Config();
                config.Sensors.Add(example);

                File.WriteAllText(CONFIG, JsonConvert.SerializeObject(config, Formatting.Indented));
            }
            else
            {
                config = JsonConvert.DeserializeObject<Config>(File.ReadAllText(CONFIG));
            }

            return config;
        }

        public void Save()
        {
            File.WriteAllText(CONFIG, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }
}
