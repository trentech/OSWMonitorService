using MySql.Data.MySqlClient;
using OSWMonitorService.JSON;
using Serilog;
using System.Data;

namespace OSWMonitorService.DataTypes
{
    internal class MySQL
    {
        MySqlConnection db;

        public MySQL(Config config)
        {
            DataType dataType = config.DataType;
            string connection = string.Format("Server={0}; database={1}; UID={2}; password={3}", dataType.Path, dataType.Name, dataType.Username, dataType.Password);
            db = new MySqlConnection(connection);
        }

        public void Close()
        {
            db.Close();
        }

        public void AddEntry(Sensor sensor)
        {
            if (!TableExists(sensor))
            {
                CreateTable(sensor);
            }

            if (db.State != ConnectionState.Open)
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
            }

            MySqlCommand command = new MySqlCommand("INSERT INTO " + sensor.IP.Replace(".", "") + " ([Temperature], [Humidity], [Dew], [Online], [DateTime]) VALUES (?,?,?,?,?)", db);

            command.Parameters.AddWithValue("@Temperature", sensor.Temperature);
            command.Parameters.AddWithValue("@Humidity", sensor.Humidity);
            command.Parameters.AddWithValue("@Dew", sensor.DewPoint);
            command.Parameters.AddWithValue("@Online", sensor.IsOnline);
            command.Parameters.AddWithValue("@DateTime", GetDateTime(sensor.DateTime));

            command.ExecuteNonQuery();

            db.Close();
        }

        private bool TableExists(Sensor sensor)
        {
            if (db.State != ConnectionState.Open)
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
            }

            MySqlCommand command = new MySqlCommand("SHOW TABLES LIKE \'" + sensor.IP.Replace(".", "") + "\'", db);

            command.Prepare();

            if (command.ExecuteReader().HasRows)
            {
                return true;
            }

            return false;
        }

        private void CreateTable(Sensor sensor)
        {
            if (db.State != ConnectionState.Open)
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
            }

            MySqlCommand command = new MySqlCommand("CREATE TABLE " + sensor.IP.Replace(".", "") + " ([Temperature] DOUBLE, [Humidity] DOUBLE, [Dew] DOUBLE, [Online] BIT, [DateTime] DateTime)", db);

            command.ExecuteNonQuery();
        }

        private DateTime GetDateTime(DateTime d)
        {
            // REMOVES MILLISECONDS. WORKAROUND FOR ERROR 'Data type mismatch in criteria expression'
            return new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second);
        }
    }
}
