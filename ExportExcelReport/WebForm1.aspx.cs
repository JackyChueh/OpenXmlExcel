using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;


namespace ExportExcelReport
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            //Excel.InsertText(@"C:\1\SheetX.xlsx", "Inserted Text");

            string sql = "SELECT ''";
            DataTable dt = DataAccessLayer.SelectDataTable(sql, System.Configuration.ConfigurationManager.ConnectionStrings["TII"].ToString());
            //CreateExcelFile.CreateExcelDocument(dt, "test.xlsx", this.Response);

            ExportExcelReport.ExportToExcelLight.CreateExcelDocument(dt, "light.xlsx", this.Response);

            //MemoryStream memoryStream = new MemoryStream();
            //outStream.CopyTo(memoryStream);

            //Excel.CreateSpreadsheetWorkbook(@"c:\1\SheetZZZ.xlsx");
            //using (StreamWriter writer = new StreamWriter(Response.OutputStream, Encoding.UTF8))
            //{
            //    //writer.Write("This is the content");

            //}

        }


        //private void ResponseExcel()
        //{
        //    SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        //    // Add a WorkbookPart to the document.
        //    WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        //    workbookpart.Workbook = new Workbook();

        //    // Add a WorksheetPart to the WorkbookPart.
        //    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        //    worksheetPart.Worksheet = new Worksheet(new SheetData());

        //    // Add Sheets to the Workbook.
        //    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
        //        AppendChild<Sheets>(new Sheets());

        //    // Append a new worksheet and associate it with the workbook.
        //    Sheet sheet = new Sheet()
        //    {
        //        Id = spreadsheetDocument.WorkbookPart.
        //            GetIdOfPart(worksheetPart),
        //        SheetId = 1,
        //        Name = "mySheet"
        //    };
        //    sheets.Append(sheet);

        //    //workbookpart.Workbook.Save();
        //    workbookpart.Workbook.Save(stream);

        //    Stream outStream = new MemoryStream();
        //    Excel.ResponseSpreadsheetWorkbook(outStream);
        //    MemoryStream memoryStream = new MemoryStream();
        //    outStream.CopyTo(memoryStream);

        //    Response.Clear();
        //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //    Response.AddHeader("Content-Disposition", "attachment; filename=test.xlsx");
        //    //Response.BinaryWrite(myMemoryStream.ToArray());
        //    memoryStream.WriteTo(Response.OutputStream); //works too
        //    Response.Flush();
        //    Response.Close();
        //    Response.End();
        //}
    }
}