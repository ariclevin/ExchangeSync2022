using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchangeSync
{
    public static class Trace
    {
        private static StringBuilder sb;
        private static List<LogRow> list;
        private static List<TimeLogRow> timeList;

        public static bool Null
        {
            get
            {
                if (list != null)
                {
                    return false;
                }
                else
                    return true;
            }
        }

        public static bool Empty
        {
            get
            {
                if (list != null)
                {
                    if (list.Count > 0)
                        return false;
                    else
                        return true;
                }
                else
                    return true;
            }
        }

        public static int Count
        {
            get
            {
                return list.Count;
            }
        }

        public static int Criticals { get; set; }
        public static int Errors { get; set; }
        public static int Warnings { get; set; }

        public static void InitializeLog()
        {
            list = new List<LogRow>();
            // sb = new StringBuilder();
            Errors = 0;
        }

        public static void Reset()
        {
            if (list == null)
            {
                list = new List<LogRow>();
                Errors = 0;
                Criticals = 0;
                Warnings = 0;
            }
            else
            {
                if (!Empty)
                    ClearLog();
            }
        }

        public static void ClearLog()
        {
            list.Clear();
            Errors = 0;
            Criticals = 0;
            Warnings = 0;
        }

        private static string AddLogHeader()
        {
            return "\"Level\",\"Date and Time\",\"Source\",\"Category\",\"Message\",\"Details\"";
        }

        public static void AddLog(EventLevel level, DateTime dateTime, string source, string category, string message, string details = "")
        {
            if (CanLog(level))
            {
                LogRow row = new LogRow(level, dateTime, source, category, message, details);
                list.Add(row);
            }
        }

        public static void AddLog(EventLevel level, DateTime dateTime, string category, string methodName, string commandName, string message, string details, string parameters)
        {
            if (CanLog(level))
            {
                LogRow row = new LogRow(level, dateTime, category, methodName, commandName, message, details, parameters);
                list.Add(row);
            }
        }


        private static bool CanLog(EventLevel level)
        {
            bool rc = false;
            if (level != EventLevel.Verbose)
            {
                if (!string.IsNullOrEmpty(AppSetting.LoggingCriteria.Value))
                {
                    string loggingCriteria = AppSetting.LoggingCriteria.Value;
                    switch (level)
                    {
                        case EventLevel.Information:
                            if (loggingCriteria.Substring(0, 1) == "1")
                                rc = true;
                            break;
                        case EventLevel.Warning:
                            if (loggingCriteria.Substring(1, 1) == "1")
                                rc = true; Warnings++;
                            break;
                        case EventLevel.Error:
                            if (loggingCriteria.Substring(2, 1) == "1")
                                rc = true; Errors++;
                            break;
                        case EventLevel.Critical:
                            if (loggingCriteria.Substring(3, 1) == "1")
                                rc = true; Criticals++;
                            break;
                    }
                }
            }
            else
            {
                // If Verbose, Check for AppSetting LoggingCriteria Verbose
                if (!string.IsNullOrEmpty(AppSetting.LoggingCriteriaVerbose.Value))
                {
                    if (AppSetting.LoggingCriteriaVerbose.Value.ToLower() == "yes")
                        rc = true;
                }
            }
            return rc;
        }

        public static void AddLog(string line)
        {
            sb.AppendLine(line);
        }

        public static List<LogRow> RetrieveLog()
        {
            return list;
        }

        public static string RetrieveLogAsText(bool includeHeader = false)
        {
            StringBuilder sb = new StringBuilder();
            if (includeHeader)
                sb.AppendLine(AddLogHeader());

            foreach (LogRow row in list)
            {
                sb.Append("\"" + row.Level.ToString() + "\"");
                sb.Append(",");
                sb.Append("\"" + row.LogDateTime.ToShortDateString() + "\"");
                sb.Append(",");
                sb.Append("\"" + row.Category + "\"");
                sb.Append(",");
                sb.Append("\"" + row.MethodName + "\"");
                sb.Append(",");
                sb.Append("\"" + row.CommandName + "\"");
                sb.Append(",");
                sb.Append("\"" + row.Message + "\"");
                sb.Append(",");
                sb.AppendLine("\"" + row.Details + "\"");
                sb.Append(",");
                sb.AppendLine("\"" + row.Parameters + "\"");

            }
            return sb.ToString();
        }

        public static void TerminateLog()
        {
            sb = null;
        }

    }

    public class LogRow
    {
        public EventLevel Level { get; set; }
        public DateTime LogDateTime { get; set; }
        public string Category { get; set; }
        public string MethodName { get; set; }
        public string CommandName { get; set; }
        public string Message { get; set; }
        public string Details { get; set; }
        public string Parameters { get; set; }


        public LogRow(EventLevel level, DateTime logDateTime, string source, string category, string message, string details)
        {
            Level = level;
            LogDateTime = logDateTime;
            MethodName = source;
            Category = category;
            Message = message;
            Details = details;
        }

        public LogRow(EventLevel level, DateTime logDateTime, string category, string methodname, string commandName, string message, string details, string parameters)
        {
            Level = level;
            LogDateTime = logDateTime;
            Category = category;
            MethodName = methodname;
            CommandName = commandName;
            Message = message;
            Details = details;
            Parameters = parameters;

        }
    }

    public class TimeLogRow
    {
        public DateTime StartDateTime { get; set; }
        public DateTime EndDateTime { get; set; }
        public string TotalTime { get; set; }
        public string ActionType { get; set; }
        public string Result { get; set; }
        public string Details { get; set; }

        public TimeLogRow(DateTime startDateTime, DateTime endDateTime, string totalTime, string action, string result, string details)
        {
            StartDateTime = startDateTime;
            EndDateTime = endDateTime;
            TotalTime = totalTime;
            ActionType = action;
            Result = result;
            Details = details;
        }
    }
}
