using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

using NLog;
using SeRoundingReport.Models;

namespace SeRoundingReport.Services
{
    class SqlService
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private string connectionString;
        public SqlService(string connectionString)
        {
            this.connectionString = connectionString;
        }

        #region Get Door Counts
        public Dictionary<string, int> GetSupervisorDoorCounts()
        {
            try
            {
                var results = new Dictionary<string, int>();

                string query = @"select dc.textcol1 as 'Post', count(d.uiname) as 'TotalDoors'
                                from door d
                                join doorcustom dc on (dc.objectidhi = d.objectidhi and dc.objectidlo = d.objectidlo)
                                where dc.textcol1 is not null
                                and (dc.textcol1 like '%supervisor%')
                                group by dc.textcol1
                                order by dc.textcol1";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                results.Add(reader.GetString(0), reader.GetInt32(1));
                            }
                        }
                    }
                }

                return results;
            }
            catch (Exception ex) { logger.Error(ex, "SqlService <GetSupervisorDoorCounts> method."); return null; }
        }

        public Dictionary<string, int> GetOfficerDoorCounts()
        {
            try
            {
                var results = new Dictionary<string, int>();

                string query = @"select substring(d.uiname,1,abs(charindex('.',d.uiname)-1)) as 'ClinicalBldg',count(d.uiname) as 'TotalCount'
                                from door d
                                join doorcustom dc on (dc.objectidhi = d.objectidhi and dc.objectidlo = d.objectidlo)
                                where dc.textcol1 is not null
                                and (dc.textcol1 like '%officer%')
                                group by substring(d.uiname,1,abs(charindex('.', d.uiname)-1))
                                order by ClinicalBldg";

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(query, connection))
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                results.Add(reader.GetString(0), reader.GetInt32(1));
                            }
                        }
                    }
                }

                return results;
            }
            catch (Exception ex) { logger.Error(ex, "SqlService <GetOfficerDoorCount> method."); return null; }
        }
        #endregion

        #region Get Rounding Data
        public DataTable GetSupervisorRounds(Supervisor sup, string customEndDate = "", bool getRawData = false)
        {
            try
            {
                var results = new DataTable("SupervisorRounds");

                #region Query
                string query = @";with sup as (
	                                select distinct 
                                        datepart(d, [timestamp]) as 'TimeStamp'
                                        ,datepart(hh, [timestamp]) as 'TimeStampHour'
                                        ,dooridhi,dooridlo,personidlo,nonabacardnumber
	                                from accessevent ae
	                                where 
		                                [timestamp] >= @weekstart and [timestamp] <= @weekend 
		                                and (datepart(hour, [timestamp]) between @start and @end or datepart(hour, [timestamp]) between @start2 and @end2)
                                        and EventClass in (0,4,46)
                                ), groupedEvents as (
	                                select 
		                                dc.textcol1 as 'Post'
		                                ,case when [timestamp] is null then 0 else 1 end as 'Count'
	                                from sup
	                                right join doorcustom dc on (
		                                dc.objectidhi = sup.DoorIdHi and dc.objectidlo = sup.DoorIdLo 
		                                and sup.NonAbaCardNumber = @cardnumber
	                                )
	                                join door d on (dc.ObjectIdHi = d.objectidhi and dc.objectidlo = d.objectidlo)
	                                where
		                                dc.textcol1 is not null and dc.textcol1 like '%supervisor%'
                                )
                                select 
	                                Post,sum([Count]) as 'TotalCount' 
                                from groupedEvents
                                group by Post";
                #endregion

                #region Raw Data Query
                string queryRawData = @";with sup as (
	                                select distinct 
                                        datepart(d, [timestamp]) as 'TimeStamp'
                                        ,datepart(hh, [timestamp]) as 'TimeStampHour'
                                        ,dooridhi,dooridlo,personidlo,nonabacardnumber
	                                from accessevent ae
	                                where 
		                                [timestamp] >= @weekstart and [timestamp] <= @weekend 
		                                and (datepart(hour, [timestamp]) between @start and @end or datepart(hour, [timestamp]) between @start2 and @end2)
                                        and EventClass in (0,4,46)
                                ), groupedEvents as (
	                                select 
		                                dc.textcol1 as 'Post'
		                                ,d.uiname as 'DoorName'
		                                ,coalesce([timestamp], 0) as 'Day'
		                                ,coalesce(timestamphour, 0) as 'Hour'
		                                ,case when [timestamp] is null then 0 else 1 end as 'Count'
	                                from sup
	                                right join doorcustom dc on (
		                                dc.objectidhi = sup.DoorIdHi and dc.objectidlo = sup.DoorIdLo 
		                                and sup.NonAbaCardNumber = @cardnumber
	                                )
	                                join door d on (dc.ObjectIdHi = d.objectidhi and dc.objectidlo = d.objectidlo)
	                                where
		                                dc.textcol1 is not null and dc.textcol1 like '%supervisor%'
                                )
                                select * from groupedEvents
                                order by Post, DoorName";
                #endregion

                var queryToExecute = getRawData ? queryRawData : query;

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(queryToExecute, connection))
                    {
                        DateTime nowDT;
                        if (string.IsNullOrEmpty(customEndDate))
                            nowDT = DateTime.Now.AddDays(-1);
                        else
                            nowDT = DateTime.Parse(customEndDate);

                        int start, end, start2, end2;
                        if (sup.StartOfShift < sup.EndOfShift)
                        {
                            start = sup.StartOfShift;
                            end = sup.EndOfShift;
                            start2 = 0;
                            end2 = 0;
                        }
                        else
                        {
                            start = sup.StartOfShift;
                            end = 24;
                            start2 = 0;
                            end2 = sup.EndOfShift;
                        }

                        command.Parameters.AddWithValue("@weekstart", nowDT.AddDays(-6).ToShortDateString());
                        command.Parameters.AddWithValue("@weekend", new DateTime(nowDT.Year, nowDT.Month, nowDT.Day,23,59,59));
                        command.Parameters.AddWithValue("@cardnumber", sup.CardNumber);
                        command.Parameters.AddWithValue("@start", start);
                        command.Parameters.AddWithValue("@end", end);
                        command.Parameters.AddWithValue("@start2", start2);
                        command.Parameters.AddWithValue("@end2", end2);

                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(results);
                        }
                    }
                }

                return results;
            }
            catch (Exception ex) { logger.Error(ex, "SqlService <GetSupervisorRounds> method."); return null; }
        }

        public DataTable GetOfficerRounds(int start, int end, string jobTitle, string customEndDate = "", bool getRawData = false)
        {
            try
            {
                var results = new DataTable("OfficerRounds");

                #region Query
                string query = @";with pso as (
	                                select distinct 
                                        datepart(d, [timestamp]) as 'TimeStamp'
                                        ,datepart(hh, [timestamp]) as 'TimeStampHour'
                                        ,dooridhi,dooridlo,personidlo,cardnumber
	                                from accessevent ae
	                                where 
		                                [timestamp] >= @weekstart and [timestamp] <= @weekend 
                                        and (datepart(hour, [timestamp]) between @start and @end or datepart(hour, [timestamp]) between @start2 and @end2)
                                        and EventClass in (0,4,46)
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
		                                --and datepart(hour, pso.[timestamp]) between @start and @end
	                                )
	                                join door d on (dc.ObjectIdHi = d.objectidhi and dc.objectidlo = d.objectidlo)
	                                where
		                                dc.textcol1 is not null and dc.textcol1 like '%officer%'
                                )
                                select 
	                                substring(DoorName,1,abs(charindex('.',doorname)-1)) as 'ClinicalBldg',sum([Count]) as 'TotalCount' 
                                from groupedEvents
                                group by substring(DoorName,1,abs(charindex('.',doorname)-1))";
                #endregion

                #region Raw Data Query
                string queryRawData = @";with pso as (
	                                        select distinct 
                                                datepart(d, [timestamp]) as 'TimeStamp'
                                                ,datepart(hh, [timestamp]) as 'TimeStampHour'
                                                ,dooridhi,dooridlo,personidlo,cardnumber
	                                        from accessevent ae
	                                        where 
		                                        [timestamp] >= @weekstart and [timestamp] <= @weekend 
                                                and (datepart(hour, [timestamp]) between @start and @end or datepart(hour, [timestamp]) between @start2 and @end2)
                                                and EventClass in (0,4,46)
                                        ), groupedEvents as (
	                                        select 
		                                        d.uiname as 'DoorName'
		                                        ,dc.textcol1 as 'Post'
		                                        ,coalesce([timestamp], 0) as 'Day'
		                                        ,coalesce(timestamphour, 0) as 'Hour'
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
		                                        --and datepart(hour, pso.[timestamp]) between @start and @end
	                                        )
	                                        join door d on (dc.ObjectIdHi = d.objectidhi and dc.objectidlo = d.objectidlo)
	                                        where
		                                        dc.textcol1 is not null and dc.textcol1 like '%officer%'
                                        )
                                        select 
	                                        DoorName as 'ClinicalBldg',[Day],[Hour],[Count]
                                        from groupedEvents
                                        order by DoorName,[day],[hour]";
                #endregion

                var queryToExecute = getRawData ? queryRawData : query;

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(queryToExecute, connection))
                    {
                        DateTime nowDT;
                        if (string.IsNullOrEmpty(customEndDate))
                            nowDT = DateTime.Now.AddDays(-1);
                        else
                            nowDT = DateTime.Parse(customEndDate);

                        int start1, end1, start2, end2;
                        if (start < end)
                        {
                            start1 = start;
                            end1 = end;
                            start2 = 0;
                            end2 = 0;
                        }
                        else
                        {
                            start1 = start;
                            end1 = 24;
                            start2 = 0;
                            end2 = end;
                        }

                        command.Parameters.AddWithValue("@weekstart", nowDT.AddDays(-6).ToShortDateString());
                        command.Parameters.AddWithValue("@weekend", new DateTime(nowDT.Year, nowDT.Month, nowDT.Day, 23, 59, 59));
                        command.Parameters.AddWithValue("@jobtitle", jobTitle + "%");
                        command.Parameters.AddWithValue("@start", start1);
                        command.Parameters.AddWithValue("@end", end1);
                        command.Parameters.AddWithValue("@start2", start2);
                        command.Parameters.AddWithValue("@end2", end2);

                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(results);
                        }
                    }
                }

                return results;
            }
            catch (Exception ex) { logger.Error(ex, "SqlService <GetOfficerRounds> method."); return null; }
        }
        #endregion

    }
}
