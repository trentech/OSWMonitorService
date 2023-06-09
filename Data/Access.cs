﻿using OSWMonitorService.JSON;
using Serilog;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;

namespace OSWMonitorService.DataTypes
{
    internal class Access
    {
        OleDbConnection db;
        private readonly object _lock = new object();

        public Access(Config config)
        {
            string dbFile = Path.Combine(config.DataType.Path, config.DataType.Name) + ".accdb";
            db = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0; Data Source = " + dbFile);
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

            string tableName = sensor.IP.Replace(".", "");

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

            try
            {
                lock (_lock)
                {
                    OleDbCommand command = new OleDbCommand("INSERT INTO " + tableName + " ([Temperature], [Humidity], [Dew], [Online], [DateTime]) VALUES (?,?,?,?,?)", db);

                    command.Parameters.AddWithValue("@Temperature", sensor.Temperature);
                    command.Parameters.AddWithValue("@Humidity", sensor.Humidity);
                    command.Parameters.AddWithValue("@Dew", sensor.DewPoint);
                    command.Parameters.AddWithValue("@Online", sensor.IsOnline);
                    command.Parameters.AddWithValue("@DateTime", GetDateTime(sensor.DateTime));

                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Log.Error("Failed to write to Database file.");
                Log.Error(ex.Message);
            }

            db.Close();
        }

        private bool TableExists(Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");

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

            DataTable schema = db.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

            if (schema.Rows.OfType<DataRow>().Any(r => r.ItemArray[2].ToString().ToLower() == tableName.ToLower()))
            {
                return true;
            }

            return false;
        }

        private void CreateTable(Sensor sensor)
        {
            string tableName = sensor.IP.Replace(".", "");

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

            OleDbCommand command = new OleDbCommand("CREATE TABLE " + tableName + " ([Temperature] DOUBLE, [Humidity] DOUBLE, [Dew] DOUBLE, [Online] BIT, [DateTime] DateTime)", db);

            command.ExecuteNonQuery();

            command = new OleDbCommand("INSERT INTO [Names] ([IP], [Name]) VALUES (?,?)", db);

            command.Parameters.AddWithValue("@IP", tableName);
            command.Parameters.AddWithValue("@SensorName", sensor.Name);

            command.ExecuteNonQuery();
        }

        private DateTime GetDateTime(DateTime d)
        {
            // REMOVES MILLISECONDS. WORKAROUND FOR ERROR 'Data type mismatch in criteria expression'
            return new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second);
        }
    }
}
