using Microsoft.Extensions.Logging;
using NPOI.POIFS.Crypt.Dsig;
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

                OleDbConnection NamesDB = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0; Data Source = C:\Users\monroett\Desktop\OSWSensors.accdb");

                try
                {
                    NamesDB.Open();
                }
                catch (Exception ex)
                {
                    logger.LogError("Failed to open Database file: {time}", DateTime.Now);
                    logger.LogError(ex.Message, DateTime.Now);
                    return;
                }

                OleDbCommand NamesCommand = new OleDbCommand("SELECT * FROM 11;", NamesDB);
                OleDbDataReader dr = NamesCommand.ExecuteReader();

                string theColumns = "";
                for (int column = 0; column < dr.FieldCount; column++)
                {
                    theColumns += dr.GetName(column) + " | ";
                }
                logger.LogInformation(theColumns +": {time}", DateTime.Now);

                NamesDB.Close();

                logger.LogInformation("Waiting 60 seconds: {time}", DateTime.Now);
                await Task.Delay(1000 * 60, stoppingToken);
            }
        }
    }
}
