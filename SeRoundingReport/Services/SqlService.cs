using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SeRoundingReport.Models;

namespace SeRoundingReport.Services
{
    class SqlService
    {
        private string connectionString;
        public SqlService(string connectionString)
        {
            this.connectionString = connectionString;
        }

        public DataTable GetSupervisorRounds(Supervisor sup)
        {
            try
            {
                var results = new DataTable("SupervisorRounds");

                string query = @";with sup as (
	                                select [timestamp],dooridhi,dooridlo,personidlo,nonabacardnumber
	                                from accessevent ae
	                                where 
		                                [timestamp] >= @weekstart and [timestamp] <= @weekend and EventClass in (0,4,46)
                                ), groupedEvents as (
	                                select 
		                                dc.textcol1 as 'Post'
		                                ,case when [timestamp] is null then 0 else 1 end as 'Count'
	                                from sup
	                                right join doorcustom dc on (
		                                dc.objectidhi = sup.DoorIdHi and dc.objectidlo = sup.DoorIdLo 
		                                and sup.NonAbaCardNumber = @cardnumber
		                                and datepart(hour, sup.[timestamp]) between @start and @end
	                                )
	                                join door d on (dc.ObjectIdHi = d.objectidhi and dc.objectidlo = d.objectidlo)
	                                where
		                                dc.textcol1 is not null and dc.textcol1 like '%supervisor%'
                                )
                                select 
	                                Post,sum([Count]) as 'TotalCount' 
                                from groupedEvents
                                group by Post";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@weekstart", DateTime.Now.AddDays(-7).ToShortDateString());
                        command.Parameters.AddWithValue("@weekend", new DateTime(DateTime.Now.Year,DateTime.Now.Month,DateTime.Now.Day,11,59,59));
                        command.Parameters.AddWithValue("@cardnumber", sup.CardNumber);
                        command.Parameters.AddWithValue("@start", sup.StartOfShift);
                        command.Parameters.AddWithValue("@end", sup.EndOfShift);

                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(results);
                        }
                    }
                }

                return results;
            }
            catch (Exception) { return null; }
        }

        public DataTable GetOfficerRounds(int start, int end, string jobTitle, string customEndDate = "")
        {
            try
            {
                var results = new DataTable("OfficerRounds");

                string query = @";with pso as (
	                                select [timestamp],dooridhi,dooridlo,personidlo,cardnumber
	                                from accessevent ae
	                                where 
		                                [timestamp] >= @weekstart and [timestamp] <= @weekend and EventClass in (0,4,46)
                                ), groupedEvents as (
	                                select 
		                                d.uiname as 'DoorName'
		                                ,dc.textcol1 as 'Post'
		                                ,case when [timestamp] is null then 0 else 1 end as 'Count'
	                                from pso
	                                join (
		                                select cardnumber 
		                                from personnel p 
		                                join personnelcustom pc on (pc.objectidlo=p.objectidlo) 
		                                where 
		                                p.[state]=1 and p.cardnumber is not null and p.cardnumber > 0x00 and p.sitecode=3107
		                                and pc.textcol5 like @jobtitle
	                                ) p on (p.CardNumber = pso.CardNumber)
	                                right join doorcustom dc on (
		                                dc.objectidhi = pso.DoorIdHi and dc.objectidlo = pso.DoorIdLo 
		                                and datepart(hour, pso.[timestamp]) between @start and @end
	                                )
	                                join door d on (dc.ObjectIdHi = d.objectidhi and dc.objectidlo = d.objectidlo)
	                                where
		                                dc.textcol1 is not null and dc.textcol1 like '%officer%'
                                )
                                select 
	                                substring(DoorName,1,abs(charindex('.',doorname)-1)) as 'ClinicalBldg',sum([Count]) as 'TotalCount' 
                                from groupedEvents
                                group by substring(DoorName,1,abs(charindex('.',doorname)-1))";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        DateTime nowDT;
                        if (string.IsNullOrEmpty(customEndDate))
                            nowDT = DateTime.Now;
                        else
                            nowDT = DateTime.Parse(customEndDate);

                        command.Parameters.AddWithValue("@weekstart", nowDT.AddDays(-7).ToShortDateString());
                        command.Parameters.AddWithValue("@weekend", new DateTime(nowDT.Year, nowDT.Month, nowDT.Day, 11, 59, 59));
                        command.Parameters.AddWithValue("@jobtitle", jobTitle + "%");
                        command.Parameters.AddWithValue("@start", start);
                        command.Parameters.AddWithValue("@end", end);

                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(results);
                        }
                    }
                }

                return results;
            }
            catch (Exception) { return null; }
        }

    }
}
