using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using LinqToExcel;
using System.IO;
using System.Web.Configuration;
using UploadExcelSimple.Models;

namespace UploadExcelSimple.Controllers
{
    public class HomeController : Controller
    {
        private string fileSavedPath = WebConfigurationManager.AppSettings["UploadPath"];

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        public ActionResult UpLoad()
        {
            var file = Request.Files as HttpFileCollectionBase;

            var uploadResult = this.FileUploadHandler(file[0]);

            if (uploadResult != null)
            {
                var DataList =  ExcelPreview(uploadResult);
                ViewData.Model = DataList;
            }

            return PartialView("_PreveiwData");
        }

  
        public IEnumerable<UserList> ExcelPreview(string filename)
        {
            var targetFile = new FileInfo(filename);
            var excelFile = new ExcelQueryFactory(filename);

            var targetSheetName = "";
            var workSheetNames = excelFile.GetWorksheetNames();
            foreach (var item in workSheetNames)
            {
                targetSheetName = item;
                break;
            }
            
            var List = excelFile.Worksheet<UserList>(targetSheetName).ToList();
            
            return List;
        }


        private string FileUploadHandler(HttpPostedFileBase file)
        {
            string result;

            try
            {
                string virtualBaseFilePath = Url.Content(fileSavedPath);
                string filePath = HttpContext.Server.MapPath(virtualBaseFilePath);

                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }

                string newFileName = string.Concat(
                    DateTime.Now.ToString("yyyyMMddHHmmssfff"),
                    Path.GetExtension(file.FileName).ToLower());

                string fullFilePath = Path.Combine(Server.MapPath(fileSavedPath), newFileName);
                file.SaveAs(fullFilePath);

                result = fullFilePath;
            }
            catch (Exception ex)
            {
                throw;
            }
            return result;
        }



    }
}