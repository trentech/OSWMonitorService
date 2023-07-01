using Newtonsoft.Json.Converters;
using Newtonsoft.Json;

namespace OSWMonitorService.JSON
{
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
            Type = DataTypes.EXCEL;
            Path = Config.PATH;
            Name = "database"; // EXCLUDE EXTENSION
            Username = "root";
            Password = "password";
        }
    }
}
