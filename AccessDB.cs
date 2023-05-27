using System.Data;
using System.Data.OleDb;

namespace OSWMontiorService
{
    public class AccessDB
    {
        public static void AddAll(ILogger<Worker> logger)
        {
            Config config = Config.Get();

            foreach (Sensor sensor in config.Sensors)
            {
                if (sensor.Skip)
                {
                    continue;
                }

                AddEntry(logger, sensor);
            }
        }

        public static void AddEntry(ILogger<Worker> logger, Sensor sensor)
        {
            if (!TableExists(logger, sensor))
            {
                CreateTable(logger, sensor);
            }

            string tableName = sensor.IP.Replace(".", "");

            using (OleDbConnection db = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0; Data Source = C:\Users\monroett\Desktop\OSWSensors.accdb"))
            {
                try
                {
                    db.Open();
                }
                catch (Exception ex)
                {
                    logger.LogError("Failed to open Database file: {time}", DateTime.Now);
                    logger.LogError(ex.Message, DateTime.Now);
                    return;
                }

                OleDbCommand command = new OleDbCommand("INSERT INTO " + tableName + " ([Temperature], [Humidity], [Dew], [Recording], [DateTime]) VALUES (?,?,?,?,?)", db);

                command.Parameters.AddWithValue("@Temperature" ,sensor.Temperature);
                command.Parameters.AddWithValue("@Humidity", sensor.Humidity);
                command.Parameters.AddWithValue("@Dew", sensor.DewPoint);
                command.Parameters.AddWithValue("@Recording", sensor.IsRecording);
                command.Parameters.AddWithValue("@DateTime", GetDateTime(sensor.Date.ToDateTime(sensor.Time)));

                command.ExecuteNonQuery();
                db.Close();
            }

        }

        private static bool TableExists(ILogger<Worker> logger, Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");

            using (OleDbConnection db = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0; Data Source = C:\Users\monroett\Desktop\OSWSensors.accdb"))
            {
                try
                {
                    db.Open();
                }
                catch (Exception ex)
                {
                    logger.LogError("Failed to open Database file: {time}", DateTime.Now);
                    logger.LogError(ex.Message, DateTime.Now);
                    return false;
                }

                DataTable schema = db.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                if (schema.Rows.OfType<DataRow>().Any(r => r.ItemArray[2].ToString().ToLower() == tableName.ToLower()))
                {
                    return true;
                }
            }

            return false;
        }

        private static void CreateTable(ILogger<Worker> logger, Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");

            using (OleDbConnection db = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0; Data Source = C:\Users\monroett\Desktop\OSWSensors.accdb"))
            {
                try
                {
                    db.Open();
                }
                catch (Exception ex)
                {
                    logger.LogError("Failed to open Database file: {time}", DateTime.Now);
                    logger.LogError(ex.Message, DateTime.Now);
                    return;
                }

                OleDbCommand command = new OleDbCommand("CREATE TABLE " + tableName + " ([Temperature] DOUBLE, [Humidity] DOUBLE, [Dew] DOUBLE, [Recording] BIT, [DateTime] DateTime)", db);

                command.ExecuteNonQuery();
                db.Close();
            }
        }

        private static DateTime GetDateTime(DateTime d)
        {
            // REMOVES MILLISECONDS. WORKAROUND FOR ERROR 'Data type mismatch in criteria expression'
            return new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second);
        }
    }
}
