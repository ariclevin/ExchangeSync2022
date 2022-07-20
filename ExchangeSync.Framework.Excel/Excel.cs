using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Infragistics.Documents.Excel;

namespace ExchangeSync
{
    public class Excel
    {
        public int CurrentRow { get; set; }

        public Workbook CreateExchangeGroupsBook()
        {
            Workbook book = new Workbook(WorkbookFormat.Excel2007);
            book.WindowOptions.ScrollBars = Infragistics.Documents.Excel.ScrollBars.Vertical;

            Worksheet sheet = book.Worksheets.Add("Exchage Server Groups");
            CurrentRow = 0;

            sheet.Rows[0].Cells[0].Value = "Group Name"; sheet.Columns[0].Width = 3000;
            sheet.Rows[0].Cells[1].Value = "Email Address"; sheet.Columns[1].Width = 3000;
            sheet.Rows[0].Cells[2].Value = "Alias"; sheet.Columns[2].Width = 3000;
            sheet.Rows[0].Cells[3].Value = "Identity"; sheet.Columns[3].Width = 3000;
            sheet.Rows[0].Cells[4].Value = "Organizational Unit"; sheet.Columns[4].Width = 10000;

            return book;
        }


        public static List<ApplicationSetting> OpenApplicationSettings(string path)
        {
            Workbook book = Workbook.Load(path);
            Worksheet sheet = book.Worksheets[0];

            List<ApplicationSetting> list = new List<ApplicationSetting>();
            foreach (WorksheetRow row in sheet.Rows)
            {
                if (row.Index > 0)
                {
                    string key = row.Cells[0].Value != null ? row.Cells[0].Value.ToString() : string.Empty; ;
                    string value = row.Cells[1].Value != null ? row.Cells[1].Value.ToString() : string.Empty;
                    string description = row.Cells[2].Value != null ? row.Cells[2].Value.ToString() : string.Empty;
                    list.Add(new ApplicationSetting(key, value, description));
                }
            }
            
            return list;
        }

        public static List<FieldMapping> OpenFieldMappings(string path)
        {
            Workbook book = Workbook.Load(path);
            Worksheet sheet = book.Worksheets[0];

            List<FieldMapping> list = new List<FieldMapping>();
            foreach (WorksheetRow row in sheet.Rows)
            {
                if (row.Index > 0)
                {
                    string crmFieldName = row.Cells[0].Value.ToString();
                    string exchangeFieldName = row.Cells[1].Value.ToString();
                    int fieldType = Convert.ToInt32(row.Cells[2].Value.ToString());
                    int dependencyType = Convert.ToInt32(row.Cells[3].Value.ToString());
                    list.Add(new FieldMapping(crmFieldName, exchangeFieldName, fieldType, dependencyType));
                }
            }

            return list;

        }
    }


}
