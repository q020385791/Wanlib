using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WanWebLib.Controllers.DOC
{
    public class WordController : Controller
    {
        //Word控制工具
        public ActionResult Index()
        {
            return View();
        }
        //直接下載
        public ActionResult DownLoadWordFile()
        {
            string Filepath = @"C:\Users\WAN\source\repos\Wanlib\WanWebLib\SampleFile\測試下載用檔案.docx";
            System.IO.FileStream fs = System.IO.File.OpenRead(Filepath);
            byte[] data = new byte[fs.Length];
           
            return File(data, System.Net.Mime.MediaTypeNames.Application.Octet, "下載名稱.docx");
            
        }
        //瀏覽器開啟
        public ActionResult OpenAtChrome()
        {
          
            string Filepath = Server.MapPath("~/SampleFile/測試下載用檔案.docx");
            string Folder = Server.UrlDecode("~/SampleFile");
            string FileName = "測試下載用檔案.docx";
            System.IO.FileStream fs = System.IO.File.OpenRead(Filepath);
            byte[] data = new byte[fs.Length];
        


            return Redirect(Folder+"//"+ FileName);
        }

        //瀏覽器開啟
        [HttpPost]
        public ActionResult PostOpenAtChrome()
        {
            //string Filepath = @"C:\Users\WAN\source\repos\Wanlib\WanWebLib\SampleFile\測試下載用檔案.docx";
            string Filepath = Server.MapPath("~/SampleFile/測試下載用檔案.docx");
            string Folder = Server.UrlDecode("~/SampleFile");
            string FileName = "測試下載用檔案.docx";
            System.IO.FileStream fs = System.IO.File.OpenRead(Filepath);
            byte[] data = new byte[fs.Length];

            int br = fs.Read(data, 0, data.Length);


            //return Redirect(Folder + "//" + FileName);

            return File(data, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "阿捏.docx");
        }
    }
}