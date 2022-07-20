using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Infragistics.Documents.Excel;

namespace ExchangeSync
{
    public class ExcelLog
    {
        public int CurrentRow { get; set; }

        public Workbook CreateBook()
        {
            Workbook book = new Workbook(WorkbookFormat.Excel2007);
            book.WindowOptions.ScrollBars = Infragistics.Documents.Excel.ScrollBars.Vertical;

            Worksheet sheet = book.Worksheets.Add("ExchangeSync Log");
            CurrentRow = 0;

            sheet.Rows[0].Cells[0].Value = "Level"; sheet.Columns[0].Width = 3000;
            sheet.Rows[0].Cells[1].Value = "Date/Time"; sheet.Columns[1].Width = 3000;
            sheet.Rows[0].Cells[2].Value = "Category"; sheet.Columns[2].Width = 3000;
            sheet.Rows[0].Cells[3].Value = "Method Name"; sheet.Columns[3].Width = 3000;
            sheet.Rows[0].Cells[4].Value = "Command Name"; sheet.Columns[4].Width = 3000;
            sheet.Rows[0].Cells[5].Value = "Message"; sheet.Columns[5].Width = 10000;
            sheet.Rows[0].Cells[6].Value = "Details"; sheet.Columns[6].Width = 10000;
            sheet.Rows[0].Cells[7].Value = "Parameters"; sheet.Columns[7].Width = 10000;

            return book;
        }

        public void AddRow(Workbook book, EventLevel level, DateTime eventDateTime, string category, string methodName, string commandName, string message, string details, string parameters)
        {
            CurrentRow++;

            Worksheet sheet = book.Worksheets[0];
            sheet.Rows[CurrentRow].Cells[0].Value = level.ToString();
            sheet.Rows[CurrentRow].Cells[1].Value = eventDateTime.ToShortDateString();
            sheet.Rows[CurrentRow].Cells[2].Value = category;
            sheet.Rows[CurrentRow].Cells[3].Value = methodName;
            sheet.Rows[CurrentRow].Cells[4].Value = commandName;
            sheet.Rows[CurrentRow].Cells[5].Value = message;
            sheet.Rows[CurrentRow].Cells[6].Value = details;
            sheet.Rows[CurrentRow].Cells[7].Value = parameters;

            if (!string.IsNullOrEmpty(details))
                sheet.Rows[CurrentRow].Cells[6].CellFormat.WrapText = ExcelDefaultableBoolean.True;

            if (!string.IsNullOrEmpty(parameters))
                sheet.Rows[CurrentRow].Cells[7].CellFormat.WrapText = ExcelDefaultableBoolean.True;

            // Set Colors for Warning, Error and Critical
            WorkbookColorInfo colorInfo;
            switch (level)
            {
                case EventLevel.Error:
                case EventLevel.Critical:
                    colorInfo = new WorkbookColorInfo(Color.Red);
                    break;
                case EventLevel.Warning:
                    colorInfo = new WorkbookColorInfo(Color.Orange);
                    break;
                default:
                    colorInfo = new WorkbookColorInfo(Color.Black);
                    break;
            }

            for (int i = 0; i<=7; i++)
            {
                sheet.Rows[CurrentRow].Cells[i].CellFormat.Font.ColorInfo = colorInfo;
            }
        }

        public string SaveBook(Workbook book, bool isAutoSync = false)
        {
            Worksheet sheet = book.Worksheets[0];
            WorksheetRegion region = new WorksheetRegion(sheet, 0, 0, CurrentRow + 2, 6);
            WorksheetTable table = region.FormatAsTable(true);
            table.Name = "Log";

            string fileName = @"C:\logs\ExchangeSync_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx";
            if (isAutoSync)
            {
                if (ProgramSetting.SourceApp == SourceApplication.Console)
                        fileName = @"C:\logs\ExchangeConsoleSync_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx";
                else
                    fileName = @"C:\logs\ExchangeAutoSync_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx";
            }
                

            book.Save(fileName);
            return fileName;

        }

        public string SaveConfigBook(Workbook book)
        {
            Worksheet sheet = book.Worksheets[0];
            WorksheetRegion region = new WorksheetRegion(sheet, 0, 0, CurrentRow + 2, 7);
            WorksheetTable table = region.FormatAsTable(true);
            table.Name = "Log";

            string fileName = @"C:\logs\ExchangeSyncConfig_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx";
            book.Save(fileName);
            return fileName;

        }


    }
}
