using Syncfusion.XlsIO;
using System.IO;
using System.Data;
using System;

namespace ServerSideApplication.Data
{
    public class ExcelService
    {
        public MemoryStream CreateExcel()
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Create a workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                

                //Initialize DataTable and data get from SampleDataTable method
                DataTable table = SampleDataTable();

                //Import data from DataTable
                worksheet.ImportDataTable(table, true, 1, 1);

                worksheet.UsedRange.AutofitColumns();

                //Save the document as a stream and return the stream.
                using (MemoryStream stream = new MemoryStream())
                {
                    //Save the created Excel document to MemoryStream
                    workbook.SaveAs(stream);
                    return stream;
                }
            }
            return null; 
        }
        private DataTable SampleDataTable()
        {
            DataTable reports = new DataTable();
            reports.Columns.Add("SalesPerson");
            reports.Columns.Add("Age", typeof(int));
            reports.Columns.Add("Salary", typeof(int));
            reports.Rows.Add("Andy Bernard", 21, 30000);
            reports.Rows.Add("Jim Halpert",25, 40000);
            reports.Rows.Add("Karen Fillippelli", 30, 50000);
            reports.Rows.Add("Phyllis Lapin", 34, 39000);
            reports.Rows.Add("Stanley Hudson", 45, 58000);

            return reports;
        }

    }
}
