using OSWMontiorService.Properties;
using System.Data;
using System.Data.OleDb;

namespace OSWMontiorService
{
    public class Access
    {
        Config config;
        ILogger<Worker> logger;

        public Access(ILogger<Worker> logger, Config config) 
        {
            this.logger = logger;
            this.config = config;
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
            string dbFile = Path.Combine(config.DataType.Datebase, "OSWSensors.accdb");

            if (!File.Exists(dbFile))
            {
                File.WriteAllBytes(dbFile, Resources.OSWSensors);
            }

            if (!TableExists(sensor))
            {
                CreateTable(sensor);
            }

            string tableName = sensor.IP.Replace(".", "");

            using (OleDbConnection db = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0; Data Source = " + Path.Combine(config.DataType.Datebase, "OSWSensors.accdb")))
            {
                try
                {
                    db.Open();
                }
                catch (Exception ex)
                {
                    logger.LogError("[{time}]: Failed to open Database file.", DateTime.Now);
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

                // BELOW IS THROWING SYNTAX ERROR. NO CLUE WHY 
                //command = new OleDbCommand("INSERT INTO Names ([IP], [Name]) VALUES (?,?)", db);

                //command.Parameters.AddWithValue("@IP", tableName);
                //command.Parameters.AddWithValue("@Name", sensor.Name);

                //command.ExecuteNonQuery();

                db.Close();
            }
        }

        private bool TableExists(Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");

            using (OleDbConnection db = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0; Data Source = " + Path.Combine(config.DataType.Datebase, "OSWSensors.accdb")))
            {
                try
                {
                    db.Open();
                }
                catch (Exception ex)
                {
                    logger.LogError("[{time}]: Failed to open Database file.", DateTime.Now);
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

        private void CreateTable(Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");

            using (OleDbConnection db = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0; Data Source = " + Path.Combine(config.DataType.Datebase, "OSWSensors.accdb")))
            {
                try
                {
                    db.Open();
                }
                catch (Exception ex)
                {
                    logger.LogError("[{time}]: Failed to open Database file", DateTime.Now);
                    logger.LogError(ex.Message, DateTime.Now);
                    return;
                }

                OleDbCommand command = new OleDbCommand("CREATE TABLE " + tableName + " ([Temperature] DOUBLE, [Humidity] DOUBLE, [Dew] DOUBLE, [Recording] BIT, [DateTime] DateTime)", db);

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
