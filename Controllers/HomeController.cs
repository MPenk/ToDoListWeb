using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Mvc.ViewEngines;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using PuppeteerSharp;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ToDoListWeb.Models;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace ToDoListWeb.Controllers
{
    public class HomeController : Controller
    {
        ToDoListModel toDoListModel;
        private ICompositeViewEngine _viewEngine;

        public HomeController(ICompositeViewEngine viewEngine)
        {
            this.toDoListModel = new ToDoListModel();
            _viewEngine = viewEngine;
        }

        async public Task<IActionResult> Index()
        {
            ToDoListModel model = this.toDoListModel;
            await model.init();
            var tasks = model.GetTasks()?? new List<TaskModel>();

            return View(tasks);
        }

        [HttpPost]
        public async Task<ActionResult> Index(string button, List<TaskModel> model)
        {
            if (ModelState.IsValid && model.Count != 0)
            {
                IList<string> TasksToExport = new List<string>();
                foreach (var task in model)
                {
                    if (task.Checked)
                        TasksToExport.Add(task.TaskName);
                }
                string site = await RenderViewToString("Render", TasksToExport);

                byte[] content;
                string mime = "pdf";

                if (button == "Eksport do DOCX")
                {
                    content = HtmlToDoc(site);
                    mime = "vnd.openxmlformats-officedocument.wordprocessingml.document";
                }
                else
                {
                    var pdfStream = HtmlToPdfAsync(site);
                    content = StreamToByteArray(await pdfStream);
                }
                return new FileContentResult(content, "application/" + mime);
            }
            ToDoListModel m = this.toDoListModel;
            await m.init();
            var tasks = m.GetTasks() ?? new List<TaskModel>();
            return  View(tasks);
        }


        private byte[] HtmlToDoc(string html)
        {
            using (MemoryStream generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(
                       generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }

                    HtmlConverter converter = new HtmlConverter(mainPart);
                    Body body = mainPart.Document.Body;

                    var paragraphs = converter.Parse(html);
                    for (int i = 0; i < paragraphs.Count; i++)
                    {
                        body.Append(paragraphs[i]);
                    }

                    mainPart.Document.Save();
                }

                return generatedDocument.ToArray();
            }
        }

        private async Task<Stream> HtmlToPdfAsync(string html)
        {
            var browserFetcher = new BrowserFetcher();
            await browserFetcher.DownloadAsync();
            await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions { Headless = true });
            await using var page = await browser.NewPageAsync();
            await page.SetContentAsync(html);
            var result = await page.GetContentAsync();
            var pdfContent = await page.PdfStreamAsync();
            return pdfContent;
        }

        private async Task<string> RenderViewToString(string viewName, object model)
        {
            if (string.IsNullOrEmpty(viewName))
                viewName = ControllerContext.ActionDescriptor.ActionName;

            ViewData.Model = model;

            using (var writer = new StringWriter())
            {
                ViewEngineResult viewResult =
                    _viewEngine.FindView(ControllerContext, viewName, false);

                ViewContext viewContext = new ViewContext(
                    ControllerContext,
                    viewResult.View,
                    ViewData,
                    TempData,
                    writer,
                    new HtmlHelperOptions()
                );

                await viewResult.View.RenderAsync(viewContext);

                return writer.GetStringBuilder().ToString();
            }
        }
        public static byte[] StreamToByteArray(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }

    }
}
