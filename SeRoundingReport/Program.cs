using System;
using System.Collections.Generic;
using System.Data;

using NLog;
using SeRoundingReport.Services;

namespace SeRoundingReport
{
    class Program
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("<< SE Rounding Report v1.0.0 | Schneider Electric HWD 2020 >>");
            Console.ForegroundColor = ConsoleColor.Gray;

            Console.WriteLine("Initializing...");
            logger.Info("Initializing...");

            var connectionString = AppConfigService.GetConnectionString();
            var supervisors = AppConfigService.GetSupervisors();
            var jobTitle = AppConfigService.GetPSOOracleJobTile();
            var customEndDate = AppConfigService.GetCustomEndDate();
            var reportPath = AppConfigService.GetReportPath();

            if (string.IsNullOrEmpty(connectionString))
            {
                Console.WriteLine("SQL ConnectionString not found.  Please make sure to set the ConnectionString in the app.config file.");
                logger.Warn("SQL ConnectionString not found.  Please make sure to set the ConnectionString in the app.config file.");
                System.Threading.Thread.Sleep(2000);
                return;
            }
            var sql = new SqlService(connectionString);

            Console.WriteLine("Running SQL queries to gather rounding data.");
            logger.Info("Running SQL queries to gather rounding data.");

            // Get Supervisor Rounds
            List<DataTable> supResultsToReport = new List<DataTable>();
            foreach (var s in supervisors)
            {
                var r = sql.GetSupervisorRounds(s.Value, customEndDate);
                if (r != null)
                {
                    supResultsToReport.Add(EditSupervisorTableForExcel(r));
                }
            }

            // Get Officer Rounds
            List<DataTable> offResults = new List<DataTable>();
            var r1 = sql.GetOfficerRounds(7, 14, jobTitle, customEndDate);
            var r2 = sql.GetOfficerRounds(15, 22, jobTitle, customEndDate);
            var r3 = sql.GetOfficerRounds(23, 6, jobTitle, customEndDate);
            if (r1 != null) offResults.Add(r1);
            if (r2 != null) offResults.Add(r2);
            if (r3 != null) offResults.Add(r3);

            List<DataTable> offResultsToReport = new List<DataTable>();
            foreach (var t in offResults)
            {
                var r = EditOfficerTableForExcel(t);
                if (r != null)
                    offResultsToReport.Add(r);
            }

            Console.WriteLine("Generating Excel report.");
            logger.Info("Generating Excel report.");

            var reportDT = string.IsNullOrEmpty(customEndDate) ? DateTime.Now : DateTime.Parse(customEndDate);
            ExcelService.GenerateReport($@"{reportPath}\UCM Rounding Report", "Public Safety - Weekly Round Report",reportDT, offResultsToReport.ToArray(), supResultsToReport.ToArray());

            Console.WriteLine("Emailing report to recipients.");
            logger.Info("Emailing report to recipients.");

            EmailSingletonService.Instance.Initialize();
            var b = EmailSingletonService.Instance.SendEmail($@"{reportPath}\UCM Rounding Report {DateTime.Now.ToString("yyyyMMdd")}.xlsx");

            if (b)
            {
                logger.Info("Email sent to recipients.");
            }
            else
            {
                logger.Warn("Failed to email report.");
            }

            Console.WriteLine("Done processing this program will terminate.");
            Environment.Exit(0);
        }

        #region Private Methods
        private static DataTable EditSupervisorTableForExcel(DataTable dt)
        {
            try
            {
                var t = new DataTable("ExcelTable");
                t.Columns.Add("Post", typeof(string));
                t.Columns.Add("Priority Taps", typeof(Int32));
                t.Columns.Add("Required Rounds", typeof(Int32));
                t.Columns.Add("Total Required", typeof(Int32));
                t.Columns.Add("Total Completed", typeof(Int32));
                t.Columns.Add("Compliance", typeof(Int32));
                t.Columns.Add("Target", typeof(Int32));

                foreach (DataRow dr in dt.Rows)
                {
                    var post = dr[0].ToString().Substring(dr[0].ToString().IndexOf("-") + 1);

                    t.Rows.Add(post, 1, 2, 0, dr[1], 0, 90);
                }
                return t;
            }
            catch (Exception) { return null; }
        }

        private static DataTable EditOfficerTableForExcel(DataTable dt)
        {
            try
            {
                var t = new DataTable("ExcelTable");
                t.Columns.Add("Clinical Building", typeof(string));
                t.Columns.Add("Priority Taps", typeof(Int32));
                t.Columns.Add("Required Rounds", typeof(Int32));
                t.Columns.Add("Total Required", typeof(Int32));
                t.Columns.Add("Total Completed", typeof(Int32));
                t.Columns.Add("Compliance", typeof(Int32));
                t.Columns.Add("Target", typeof(Int32));

                foreach (DataRow dr in dt.Rows)
                {
                    t.Rows.Add((dr[0].ToString().Equals("F", StringComparison.OrdinalIgnoreCase) ? "Comer 2 (F)" : dr[0]), 40, 4, 0, dr[1], 0, 90);
                }
                return t;
            }
            catch (Exception) { return null; }
        }
        #endregion

    }
}
