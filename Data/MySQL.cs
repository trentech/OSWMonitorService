using MySql.Data.MySqlClient;
using OSWMonitorService.JSON;
using Serilog;

namespace OSWMonitorService.DataTypes
{
    internal class MySQL
    {
        Config config;

        public MySQL(Config config)
        {
            this.config = config;
        }

        public void AddEntry(Sensor sensor)
        {
            if (!TableExists(sensor))
            {
                CreateTable(sensor);
            }

            string tableName = sensor.IP.Replace(".", "");
            DataType dataType = config.DataType;
            string connection = string.Format("Server={0}; database={1}; UID={2}; password={3}", dataType.Path, dataType.Name, dataType.Username, dataType.Password);

            using (MySqlConnection db = new MySqlConnection(connection))
            {
                try
                {
                    db.Open();
                }
                catch (Exception ex)
                {
                    Log.Error("Failed to open Database file.");
                    Log.Error(ex.Message);
                    return;
                }

                MySqlCommand command = new MySqlCommand("INSERT INTO " + tableName + " ([Temperature], [Humidity], [Dew], [Online], [DateTime]) VALUES (?,?,?,?,?)", db);

                command.Parameters.AddWithValue("@Temperature", sensor.Temperature);
                command.Parameters.AddWithValue("@Humidity", sensor.Humidity);
                command.Parameters.AddWithValue("@Dew", sensor.DewPoint);
                command.Parameters.AddWithValue("@Online", sensor.IsOnline);
                command.Parameters.AddWithValue("@DateTime", GetDateTime(sensor.DateTime));

                command.ExecuteNonQuery();
                db.Close();
            }

        }

        public bool TableExists(Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");
            DataType dataType = config.DataType;
            string connection = string.Format("Server={0}; database={1}; UID={2}; password={3}", dataType.Path, dataType.Name, dataType.Username, dataType.Password);

            using (MySqlConnection db = new MySqlConnection(connection))
            {
                try
                {
                    db.Open();
                }
                catch (Exception ex)
                {
                    Log.Error("Failed to open Database file.");
                    Log.Error(ex.Message);
                    return false;
                }

                MySqlCommand command = new MySqlCommand("SHOW TABLES LIKE \'" + tableName + "\'", db);

                command.Prepare();

                if (command.ExecuteReader().HasRows)
                {
                    return true;
                }
            }

            return false;
        }

        public void CreateTable(Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");
            DataType dataType = config.DataType;
            string connection = string.Format("Server={0}; database={1}; UID={2}; password={3}", dataType.Path, dataType.Name, dataType.Username, dataType.Password);

            using (MySqlConnection db = new MySqlConnection(connection))
            {
                try
                {
                    db.Open();
                }
                catch (Exception ex)
                {
                    Log.Error("Failed to open Database file.");
                    Log.Error(ex.Message);
                    return;
                }

                MySqlCommand command = new MySqlCommand("CREATE TABLE " + tableName + " ([Temperature] DOUBLE, [Humidity] DOUBLE, [Dew] DOUBLE, [Online] BIT, [DateTime] DateTime)", db);

                command.ExecuteNonQuery();
                db.Close();
            }
        }

        private DateTime GetDateTime(DateTime d)
        {
            // REMOVES MILLISECONDS. WORKAROUND FOR ERROR 'Data type mismatch in criteria expression'
            return new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second);
        }
    }
}
