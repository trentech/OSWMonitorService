using Microsoft.Extensions.Logging;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.Util;
using System.Data;
using System.Data.OleDb;

namespace OSWMontiorService
{
    public class DBTest : BackgroundService
    {
        private readonly ILogger<DBTest> logger;

        public DBTest(ILogger<DBTest> logger)
        {
            this.logger = logger;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "<Pending>")]
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            logger.LogInformation("Starting Dev Service: {time}", DateTime.Now);

            while (!stoppingToken.IsCancellationRequested)
            {
                logger.LogInformation("Running Task: {time}", DateTime.Now);

                Sensor sensor = new Sensor("TEST", "192.168.1.1");
                sensor.Temperature = 55.5;
                sensor.Humidity = 30.45;
                sensor.DewPoint = 35;
                sensor.IsRecording = true;
                sensor.Date = DateOnly.FromDateTime(DateTime.Now);
                sensor.Time = TimeOnly.FromDateTime(DateTime.Now);

                if(!TableExists(sensor))
                {
                    CreateTable(sensor);
                }

                AddSensorData(sensor);

                logger.LogInformation("Waiting 60 seconds: {time}", DateTime.Now);
                await Task.Delay(1000 * 60, stoppingToken);
            }
        }

        public bool TableExists(Sensor sensor)
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

        public void AddSensorData(Sensor sensor)
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

        private DateTime GetDateTime(DateTime d)
        {
            // REMOVES MILLISECONDS. WORKAROUND FOR ERROR 'Data type mismatch in criteria expression'
            return new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second);
        }

        public void CreateTable(Sensor sensor)
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
    }
}
