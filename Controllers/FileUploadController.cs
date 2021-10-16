using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BulkUpload.Controllers
{
    public class FileUploadController : Controller
    {               
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["POSX_ConnectionString"].ConnectionString);
        OleDbConnection Econ;
        // GET: Employee
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase files)
        {
            Byte[] imgbyte = null;
            if (files.ContentType== "image/jpeg")
            {
              
            }
            else { 
            string filename = Guid.NewGuid() + Path.GetExtension(files.FileName);
            string filepath = "/excelfolder/" + filename;
                files.SaveAs(Path.Combine(Server.MapPath("/excelfolder"), filename));
            InsertExceldata(filepath, filename);
            }

            return View("~/Views/FileUpload/Index.cshtml");
        }
        private void ExcelConn(string filepath)
        {
            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", filepath);
            Econ = new OleDbConnection(constr);

        }
        private void InsertExceldata(string fileepath, string filename)
        {
            try
            {
                string fullpath = Server.MapPath("/excelfolder/") + filename;
                ExcelConn(fullpath);
                string query = string.Format("Select * from [{0}]", "Sheet1$");
                OleDbCommand Ecom = new OleDbCommand(query, Econ);
                Econ.Open();

                DataSet ds = new DataSet();
                OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);
                Econ.Close();
                oda.Fill(ds);

                DataTable dt = ds.Tables[1];

                SqlBulkCopy objbulk = new SqlBulkCopy(con);
                objbulk.DestinationTableName = "TARGET_ACH_DETAILS_TEST";
                objbulk.ColumnMappings.Add("TARGETID", "TARGETID");
                objbulk.ColumnMappings.Add("PERIODID", "PERIODID");
                objbulk.ColumnMappings.Add("AREA_ID", "AREA_ID");
                objbulk.ColumnMappings.Add("PARTNER_ID", "PARTNER_ID");
                objbulk.ColumnMappings.Add("TARGET_AMT", "TARGET_AMT");
                objbulk.ColumnMappings.Add("ACH_AMT", "ACH_AMT");
                objbulk.ColumnMappings.Add("ACH_PARC", "ACH_PARC");
                objbulk.ColumnMappings.Add("CREATED_BY", "CREATED_BY");
                objbulk.ColumnMappings.Add("CREATED_AT", "CREATED_AT");
                con.Open();
                objbulk.WriteToServer(dt);
                con.Close();
            }
            catch (Exception ex)
            {

                throw ex;
            }


            
        }

    }
}