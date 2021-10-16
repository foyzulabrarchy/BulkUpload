using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BulkUpload.Models
{
    public class File_Upload
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public HttpPostedFileBase Filepic { get; set; }

        public DateTime UploadDate { get; set; }
    }
}