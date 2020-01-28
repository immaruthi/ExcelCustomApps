using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GMBRDLC_Application
{
    public partial class Home : System.Web.UI.Page
    {

        public static Byte[] PdfSharpConvert(String html)
        {
            Byte[] res = null;
            using (MemoryStream ms = new MemoryStream())
            {
                var pdf = TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerator.GeneratePdf(html, PdfSharp.PageSize.A4);
                pdf.Save(ms);
                res = ms.ToArray();
            }
            return res;
        }

        private MemoryStream createPDF(string html)
        {
            MemoryStream msOutput = new MemoryStream();
            TextReader reader = new StringReader(html);

            // step 1: creation of a document-object
            Document document = new Document(PageSize.A4, 30, 30, 30, 30);

            // step 2:
            // we create a writer that listens to the document
            // and directs a XML-stream to a file
            PdfWriter writer = PdfWriter.GetInstance(document, msOutput);

            // step 3: we create a worker parse the document
            HTMLWorker worker = new HTMLWorker(document);

            // step 4: we open document and start the worker on the document
            document.Open();
            worker.StartDocument();

            // step 5: parse the html into the document
            worker.Parse(reader);

            // step 6: close the document and the worker
            worker.EndDocument();
            worker.Close();
            document.Close();

            return msOutput;
        }

        protected void Page_Load(object sender, EventArgs e)
        {
           // string fileInformation = System.IO.File.ReadAllText(@"C:\Users\v-mapall\source\repos\ExcelCustomApps\GMBRDLC Application\Template.html");
           // //PdfSharpConvert(fileInformation);

           //// System.IO.File.WriteAllBytes(@"C:\Users\v-mapall\source\repos\ExcelCustomApps\GMBRDLC Application\hello.pdf", PdfSharpConvert(fileInformation));

           // System.IO.File.WriteAllBytes(@"C:\Users\v-mapall\source\repos\ExcelCustomApps\GMBRDLC Application\hello1.pdf", createPDF(fileInformation).ToArray());

            if (!IsPostBack)
            {


                

                //string strQuery = "SELECT * FROM tblStudent";
                //SqlDataAdapter da = newSqlDataAdapter(strQuery, con);
                //DataTable dt = newDataTable();
                //da.Fill(dt);
                //RDLC ds = newRDLC();
                //ds.Tables["tblStudent"].Merge(dt);
                DataSet ds = new DataSet();

                DataTable mergeTable = new DataTable();

                mergeTable.Clear();

                mergeTable.Columns.Add("Name", typeof(System.String));
                mergeTable.Columns.Add("NoViews", typeof(System.Int64));
                mergeTable.Columns.Add("DailyViews", typeof(System.Int64));

                DataRow workRow = mergeTable.NewRow();
                mergeTable.Rows.Add(workRow);

                mergeTable.Rows[0][0] = "maruthi";
                mergeTable.Rows[0][1] = 20;
                mergeTable.Rows[0][2] = 30;

                ds.Tables.Add(mergeTable);

                //ReportDataModels.GMBInformationDataTable gMBInformationRows = ds;

                ReportViewer1.ProcessingMode = ProcessingMode.Local;
                ReportViewer1.LocalReport.ReportPath = Server.MapPath("Report1.rdlc");
                ReportDataSource datasource = new ReportDataSource("DataSet1", ds.Tables[0]);
                ReportViewer1.LocalReport.DataSources.Clear();
                ReportViewer1.LocalReport.DataSources.Add(datasource);
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Warning[] warnings;
            string[] streamids;
            string mimeType;
            string encoding;
            string filenameExtension;

            byte[] bytes = ReportViewer1.LocalReport.Render(
                "PDF", null, out mimeType, out encoding, out filenameExtension,
                out streamids, out warnings);

            using (FileStream fs = new FileStream(@"C:\Users\v-mapall\source\repos\ExcelCustomApps\GMBRDLC Application\output.pdf", FileMode.Create))
            {
                fs.Write(bytes, 0, bytes.Length);
            }
        }
    }
}