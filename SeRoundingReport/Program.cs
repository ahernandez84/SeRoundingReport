using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SeRoundingReport.Models;
using SeRoundingReport.Services;

namespace SeRoundingReport
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("<< SE Rounding Report v1.0.0 | Schneider Electric HWD 2020 >>");
            Console.ForegroundColor = ConsoleColor.Gray;

            Console.WriteLine("Initializing...");

            var connectionString = AppConfigService.GetConnectionString();
            var supervisors = AppConfigService.GetSupervisors();
            var jobTitle = AppConfigService.GetPSOOracleJobTile();
            var customEndDate = AppConfigService.GetCustomEndDate();
            var reportPath = AppConfigService.GetReportPath();

            if (string.IsNullOrEmpty(connectionString))
            {
                Console.WriteLine("SQL ConnectionString not found.  Please make sure to set the ConnectionString in the app.config file.");
                System.Threading.Thread.Sleep(2000);
                return;
            }
            var sql = new SqlService(connectionString);

            Console.WriteLine("Running SQL queries to gather rounding data.");

            // Get Supervisor Rounds
            List<DataTable> sResults = new List<DataTable>();
            foreach (var s in supervisors)
            {
                var r = sql.GetSupervisorRounds(s.Value);
                if (r != null)
                {
                    sResults.Add(r);
                }
            }

            // Get Officer Rounds
            List<DataTable> oResults = new List<DataTable>();

            var r1 = sql.GetOfficerRounds(7, 14, jobTitle, customEndDate);
            var r2 = sql.GetOfficerRounds(15, 22, jobTitle, customEndDate);
            var r3 = sql.GetOfficerRounds(23, 6, jobTitle, customEndDate);
            if (r1 != null) oResults.Add(r1);
            if (r2 != null) oResults.Add(r2);
            if (r3 != null) oResults.Add(r3);

            List<DataTable> oResultsToReport = new List<DataTable>();
            foreach (var t in oResults)
            {
                var r = EditOfficerTableForExcel(t);
                if (r != null)
                    oResultsToReport.Add(r);
            }

            Console.WriteLine("Generating Excel report.");
            ExcelService.GenerateReport($@"{reportPath}\UCM Rounding Report", "Public Safety - Weekly Round Report", oResultsToReport.ToArray());

            Console.WriteLine("Emailing report to recipients");
            // TODO: EMAIL THE REPORT

            Console.WriteLine("Done processing, this program will terminate.");
            Environment.Exit(0);
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
                    t.Rows.Add(dr[0], 40, 4, 0, dr[1], 0, 90);
                }
                return t;
            }
            catch (Exception) { return null; }
        }

    }
}
