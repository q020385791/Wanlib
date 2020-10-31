using System;
using System.Collections.Generic;
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
        //修改並直接下載
        [HttpPost]
        public ActionResult DownLoadWordFile()
        {
            

            return View();
        }
    }
}