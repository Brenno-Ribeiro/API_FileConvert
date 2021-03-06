using Aspose.Cells;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;


namespace FileConvert.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelDocumentController : ControllerBase
    {
        private string _fullNameDocumentWithExtension;

        [HttpPost]
        [Route("ExcelToPdf")]
        public IActionResult ConvertFileTo(IFormFile file, string nameDocument)
        {
            DeleteDirectory();

            if (nameDocument == string.Empty)
            {
                return BadRequest("Insire um nome para seu arquivo!");
            }

            if(file == null)
            {
                return BadRequest("Adicione um arquivo!!");
            }

            var doc = new Workbook(file.OpenReadStream());

            _fullNameDocumentWithExtension = $"{nameDocument}.pdf";

            string path = Path.Combine(CreatDirectory(), _fullNameDocumentWithExtension);

          
            doc.Save(path, SaveFormat.Pdf);


            return DownLoad(path);
        }

        private FileContentResult DownLoad(string filePath)
        {
            var data = System.IO.File.ReadAllBytes(filePath);
            var result = new FileContentResult(data, "application/octet-stream")
            {
                FileDownloadName = filePath
            };

            return result;
        }


        private void DeleteDirectory()
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Files");

            if (Directory.Exists(path))
            {
                Directory.Delete(path,true);
            }
        }


        private string CreatDirectory()
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Files");

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            return path;
        }

    }
}
