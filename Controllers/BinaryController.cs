using BulkUpload.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BulkUpload.Controllers
{
    public class BinaryController : Controller
    {
        // GET: Binary
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]

        public ActionResult Index(File_Upload file)

        {

            try

            {

                Byte[] bytes = null;

                if (file.Filepic.FileName != null)

                {

                    Stream fs = file.Filepic.InputStream;

                    BinaryReader br = new BinaryReader(fs);

                    bytes = br.ReadBytes((Int32)fs.Length);

                    string connectionstring = Convert.ToString(ConfigurationManager.ConnectionStrings["POSX_ConnectionString"]);

                    SqlConnection con = new SqlConnection(connectionstring);



                    SqlCommand cmd = new SqlCommand("Insert into fileUpload_del(FileNames,Filepic,UploadDate) values(@FileNames,@Filepic,@UploadDate)", con);

                    cmd.CommandType = CommandType.Text;

                    cmd.Parameters.AddWithValue("@FileNames", file.Name);

                    cmd.Parameters.AddWithValue("@Filepic", bytes);

                    cmd.Parameters.AddWithValue("@UploadDate", DateTime.Now);

                    con.Open();

                    cmd.ExecuteNonQuery();

                    con.Close();

                    ViewBag.Image = ViewImage(bytes);

                }

            }

            catch (Exception ex)

            {

                throw ex;

            }

            return View();

        }

        private string ViewImage(byte[] arrayImage)

        {

            string base64String = Convert.ToBase64String(arrayImage, 0, arrayImage.Length);

            return "data:image/png;base64," + base64String;

        }
    }
}