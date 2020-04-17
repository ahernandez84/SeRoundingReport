using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SeRoundingReport.Models;

namespace SeRoundingReport.Services
{
    class AppConfigService
    {
        public static string GetConnectionString()
        {
            try
            {
                return ConfigurationManager.AppSettings["ConnectionString"];
            }
            catch (Exception) { return string.Empty; }
        }

        public static Dictionary<string, Supervisor> GetSupervisors()
        {
            Dictionary<string, Supervisor> supShifts = new Dictionary<string, Supervisor>();

            try
            {
                var first = ConfigurationManager.AppSettings["FirstShift"].Split(',');
                var second = ConfigurationManager.AppSettings["SecondShift"].Split(',');
                var third = ConfigurationManager.AppSettings["ThirdShift"].Split(',');

                supShifts.Add("First", new Supervisor { StartOfShift=first[0], EndOfShift=first[1], CardNumber=first[2] });
                supShifts.Add("Second", new Supervisor { StartOfShift = second[0], EndOfShift = second[1], CardNumber = second[2] });
                supShifts.Add("Third", new Supervisor { StartOfShift = third[0], EndOfShift = third[1], CardNumber = third[2] });

                return supShifts;
            }
            catch (Exception) { return supShifts; }
        }

        public static string GetPSOOracleJobTile()
        {
            try
            {
                return ConfigurationManager.AppSettings["PSOOracleJobTitle"];
            }
            catch (Exception) { return string.Empty; }
        }

        public static string GetCustomEndDate()
        {
            try
            {
                return ConfigurationManager.AppSettings["CustomEndDate"];
            }
            catch (Exception) { return string.Empty; }
        }

        public static string GetReportPath()
        {
            try
            {
                return ConfigurationManager.AppSettings["ReportPath"];
            }
            catch (Exception) { return string.Empty; }
        }
    }
}
