using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using static System.Environment;

namespace OSWMontiorService
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

    public class DataType
    {
        public enum DataTypes
        {
            MYSQL, EXCEL, ACCESS
        }

        [JsonConverter(typeof(StringEnumConverter))]
        public DataTypes Type { get; set; }
        [JsonProperty]
        public string Path { get; set; }
        [JsonProperty]
        public string Name { get; set; }
        [JsonProperty]
        public string Username { get; set; }
        [JsonProperty]
        public string Password { get; set; }

        public DataType()
        {
            this.Type = DataTypes.EXCEL;
            this.Path = Config.PATH; 
            this.Name = "database"; // EXCLUDE EXTENSION
            this.Username = "root";
            this.Password = "password";
        }
    }

    public class Mail
    {
        [JsonProperty]
        public string STMP { get; set; }
        [JsonProperty]
        public int Port { get; set; }
        [JsonProperty]
        public bool SSL { get; set; }
        [JsonProperty]
        public string From { get; set; }
        [JsonProperty]
        public List<string> Recipients { get; set; }

        public Mail()
        {
            STMP = "smtp.gmail.com";
            Port = 25;
            SSL = false;
            From = "example@email.com";
            Recipients = new List<string>() { "test1@example.com", "test2@example.com" };
        }
    }
}
