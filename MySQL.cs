using MySql.Data.MySqlClient;
using Serilog;

namespace OSWMontiorService
{
    public class MySQL
    {
        Config config;

        string server;
        string db;
        string username;
        string password;

        public MySQL(Config config, string server, string db, string username, string password)
        {
            this.config = config;
            this.server = server;
            this.db = db;
            this.username = username;
            this.password = password;
        }

        public void AddAll()
        {
            foreach (Sensor sensor in config.Sensors)
            {
                if (sensor.Skip)
                {
                    continue;
                }

                AddEntry(sensor);
            }
        }

        public void AddEntry(Sensor sensor)
        {
            if (!TableExists(sensor))
            {
                CreateTable(sensor);
            }

            string tableName = sensor.IP.Replace(".", "");

            string connection = string.Format("Server={0}; database={1}; UID={2}; password={3}", server, db, username, password);

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

                MySqlCommand command = new MySqlCommand("INSERT INTO " + tableName + " ([Temperature], [Humidity], [Dew], [Recording], [DateTime]) VALUES (?,?,?,?,?)", db);

                command.Parameters.AddWithValue("@Temperature" ,sensor.Temperature);
                command.Parameters.AddWithValue("@Humidity", sensor.Humidity);
                command.Parameters.AddWithValue("@Dew", sensor.DewPoint);
                command.Parameters.AddWithValue("@Recording", sensor.IsRecording);
                command.Parameters.AddWithValue("@DateTime", GetDateTime(sensor.Date.ToDateTime(sensor.Time)));

                command.ExecuteNonQuery();
                db.Close();
            }

        }

        private bool TableExists(Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");

            string connection = string.Format("Server={0}; database={1}; UID={2}; password={3}", server, db, username, password);

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

        private void CreateTable(Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");

            string connection = string.Format("Server={0}; database={1}; UID={2}; password={3}", server, db, username, password);

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

                MySqlCommand command = new MySqlCommand("CREATE TABLE " + tableName + " ([Temperature] DOUBLE, [Humidity] DOUBLE, [Dew] DOUBLE, [Recording] BIT, [DateTime] DateTime)", db);

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
