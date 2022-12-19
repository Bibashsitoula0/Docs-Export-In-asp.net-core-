
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using System;

namespace Api.Controllers
{

    [Route("api/[controller]")]
    [ApiController]
    // [Authorize]
    public class StudentController : ControllerBase
    {
        private IStudentService _studentService;
        private IWebHostEnvironment _hostingEnvironment;

        public StudentController(IStudentService studentService, IWebHostEnvironment hostingEnvironment)
        {
            _studentService = studentService;
            _hostingEnvironment = hostingEnvironment;
        }      

        [HttpGet]
        public IActionResult CreateDocument()
        {
            
                MemoryStream mem = new MemoryStream();                         
               using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
                    {
                        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                        mainPart.Document = new Document();
                        Body body = mainPart.Document.AppendChild(new Body());
                        Paragraph para = body.AppendChild(new Paragraph());
                        Run run = para.AppendChild(new Run());
                        run.AppendChild(new Text("Hello, World!"));
                        mainPart.Document.Save();
                        wordDocument.Close();
                        mem.Position = 0;               
                           
                    }
               return File(mem, "application/msword", "Sample.docx");
        }
      }
    }

