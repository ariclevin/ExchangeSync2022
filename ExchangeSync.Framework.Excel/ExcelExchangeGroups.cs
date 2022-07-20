using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Infragistics.Documents.Excel;


namespace ExchangeSync
{
    public class ExcelExchangeGroups
    {

        public int CurrentRow { get; set; }

        public Workbook CreateBook()
        {
            Workbook book = new Workbook(WorkbookFormat.Excel2007);
            book.WindowOptions.ScrollBars = Infragistics.Documents.Excel.ScrollBars.Vertical;

            Worksheet sheet = book.Worksheets.Add("Exchage Server Groups");
            CurrentRow = 0;

            sheet.Rows[0].Cells[0].Value = "Group Name"; sheet.Columns[0].Width = 3000;
            sheet.Rows[0].Cells[1].Value = "Contact Name"; sheet.Columns[1].Width = 3000;
            sheet.Rows[0].Cells[2].Value = "Email Address"; sheet.Columns[2].Width = 5000;
            sheet.Rows[0].Cells[3].Value = "Alias"; sheet.Columns[3].Width = 3000;
            sheet.Rows[0].Cells[4].Value = "Identity"; sheet.Columns[4].Width = 3000;
            sheet.Rows[0].Cells[5].Value = "Organizational Unit"; sheet.Columns[5].Width = 10000;

            return book;
        }

        public void AddRow(Workbook book, string groupName, string contactName, string emailAddress, string alias, string identity, string ou)
        {
            CurrentRow++;

            Worksheet sheet = book.Worksheets[0];
            sheet.Rows[CurrentRow].Cells[0].Value = groupName;
            sheet.Rows[CurrentRow].Cells[1].Value = contactName;
            sheet.Rows[CurrentRow].Cells[2].Value = emailAddress;
            sheet.Rows[CurrentRow].Cells[3].Value = alias;
            sheet.Rows[CurrentRow].Cells[4].Value = identity;
            sheet.Rows[CurrentRow].Cells[5].Value = ou;
        }

        public string SaveBook(Workbook book)
        {
            Worksheet sheet = book.Worksheets[0];
            WorksheetRegion region = new WorksheetRegion(sheet, 0, 0, CurrentRow + 2, 6);
            WorksheetTable table = region.FormatAsTable(true);
            table.Name = "ExchangeGroups";

            string fileName = @"C:\logs\ExchangeGroups_" + DateTime.Now.ToString("yyMMddHHmmss") + ".xlsx";

            book.Save(fileName);
            return fileName;
        }

    }
}
