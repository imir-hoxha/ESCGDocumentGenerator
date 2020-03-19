using ECSGDocumentGenerator;
using ECSGDocumentGenerator.Model;
using Microsoft.AspNet.OData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Http;
using System.IO;
using Newtonsoft.Json;
using ESCGDocumentGenertor.API.Models;
using Microsoft.AspNet.OData.Routing;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web.Http.Cors;

namespace ESCGDocumentGenertor.API.Controllers
{
   // [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class DocumentsController : ODataController
    {
        List<Document> docs = new List<Document>();

        public DocumentsController()
        {
            docs.Add(new Document() { Id = 1, Url = "http://s2013" });
            docs.Add(new Document() { Id = 2, Url = "http://s2016" });
            docs.Add(new Document() { Id = 3, Url = "http://s2019" });
        }


        [HttpGet]
        public HttpResponseMessage Get()
        {
            byte[] myByteArray = new byte[10];
            MemoryStream stream = new MemoryStream();
            stream.Write(myByteArray, 0, myByteArray.Length);
            HttpResponseMessage result = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = "testing.docx";
            //string responseUri = result.RequestMessage.RequestUri.AbsoluteUri.ToString();
            //string redirectedUrl = null;
            //if (result.StatusCode == HttpStatusCode.OK)
            //{
            //    HttpResponseHeaders headers = result.Headers;
            //    if (headers != null && headers.Location != null)
            //    {
            //        redirectedUrl = headers.Location.AbsoluteUri;
            //    }
            //}
            return result;
        }

        //[HttpGet]
        //public IHttpActionResult Get()
        //{
        //    //Content c = new Content();
        //    //c.Hit = new Hit() { Id = content.Id, RefId = content.RefId, Article259 = content.Article259 };
        //    //Document d = new Document() { Id = "abc", Url = "http://sp2016" };
        //    return Ok(docs);
        //}

        [HttpGet]
        public IHttpActionResult Get([FromODataUri] int key)
        {
            //Content c = new Content();
            //c.Hit = new Hit() { Id = content.Id, RefId = content.RefId, Article259 = content.Article259 };
            //Document d = new Document() { Id = "abc", Url = "http://sp2016" };

            var _docs = docs.FirstOrDefault(p => p.Id == key);
            return Ok(_docs);
        }

        [HttpGet]
        public HttpResponseMessage MergeDocuments([FromODataUri] string country, [FromODataUri] string listId, 
            [FromODataUri] string uniqueId, [FromODataUri] IEnumerable<string> docs)
        {
            byte[] myByteArray = new byte[10];
            MemoryStream stream = new MemoryStream();
            stream.Write(myByteArray, 0, myByteArray.Length);
            HttpResponseMessage result = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
            result.Content = new StreamContent(stream);
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            result.Content.Headers.ContentDisposition.FileName = "testing.docx";
            return result;
        }


        [HttpGet]
        public IHttpActionResult SearchTopics([FromODataUri] string dg)
        {
            //if (!ModelState.IsValid)
            //{
            //    return BadRequest(ModelState);
            //}
            //Content c = new Content();
            //c.Hit = new Hit() { Id = content.Id, RefId = content.RefId, Article259 = content.Article259};
            // Document d = new Document() { Id = content.Id, Url = content.Url };
            //Hit hit = new Hit() { Id = content.Id, RefId = content.RefId, Article259 = content.Article259 };
            //var invites = parameters["ArrayHere"] as IEnumerable<string>;
            return Ok("OK");
            //if (!ModelState.IsValid)
            //{
            //    return BadRequest(ModelState);
            //}

            //string jsonBody = "{\"memberState\":\"SE\"}";
            //HttpWebResponse resp = ThemesInfringementService.GetSensitiveJsonData($"{ThemesHeadersConfiguration.ThemesServiceUrl}?size=10&page=0&sort=ASC",
            //    jsonBody,
            //    ThemesHeadersConfiguration.ThemesAuthenticationToken,
            //    ThemesHeadersConfiguration.ThemesApplicationHeader,
            //    ThemesHeadersConfiguration.ThemesHost);

            //if (resp.StatusCode == HttpStatusCode.OK)
            //{
            //    WebResponse response = resp;
            //    Stream responseStream = response.GetResponseStream();
            //    StreamReader streamReader = new StreamReader(responseStream);
            //    string responseData = streamReader.ReadToEnd();
            //    response.Close();

            //    //GenerateDocument(responseData, "header.docx");

            //}
            //else
            //{
            //    //Console.WriteLine(resp.StatusCode);
            //}


        }

        //private static void GenerateDocument(string jsonData, string fileTemplate)
        //{
        //    DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SomeDocDocumentGenerator", "1.0", GetDataContext(jsonData), fileTemplate, false);
        //    SensitiveDocumentGenerator myDocGen = new SensitiveDocumentGenerator(generationInfo);

        //    string bodyDoc = Path.Combine(@"Docs\templates", "DG Sensitive report-BodyTemplate.docx");
        //    //byte[] result = myDocGen.MergeAndGenerateTemplate(bodyDoc);
        //    byte[] result = myDocGen.GenerateAndMergeTemplates(bodyDoc, GetAllDataContext(jsonData));
        //    WriteOutputToFile("NonSensitiveGeneratedDocument-v2.docx", "DG Sensitive report-BodyTemplate.docx", result);
        //}

        //private static void WriteOutputToFile(string fileName, string templateName, byte[] fileContents)
        //{
        //    ConsoleColor consoleColor = Console.ForegroundColor;

        //    if (fileContents != null)
        //    {
        //        File.WriteAllBytes(Path.Combine("Docs", fileName), fileContents);
        //        Console.ForegroundColor = ConsoleColor.Green;
        //        Console.WriteLine($"Generation succeeded for template({templateName}) --> {fileName}");
        //        Console.WriteLine();
        //    }
        //    else
        //    {
        //        Console.ForegroundColor = ConsoleColor.Red;
        //        Console.WriteLine($"Generation failed for template({templateName}) --> {fileName}");
        //    }

        //    Console.ForegroundColor = consoleColor;
        //}

        //private static Content GetDataContext(string jsonData)
        //{
        //    if (string.IsNullOrEmpty(jsonData))
        //    {
        //        throw new ArgumentNullException("jsonData");
        //    }

        //    var data = JsonConvert.DeserializeObject<SensitiveReport>(jsonData).Content.ToArray().FirstOrDefault();

        //    return data;
        //}

        //private static Content[] GetAllDataContext(string jsonData)
        //{
        //    if (string.IsNullOrEmpty(jsonData))
        //    {
        //        throw new ArgumentNullException("jsonData");
        //    }

        //    var data = JsonConvert.DeserializeObject<SensitiveReport>(jsonData).Content.ToArray();

        //    return data;
        //}


        //private static DocumentGenerationInfo GetDocumentGenerationInfo(string docType, string docVersion, Content dataContext, string wordTemplateFile, bool useDataBoundControls)
        //{
        //    DocumentGenerationInfo generationInfo = new DocumentGenerationInfo();
        //    generationInfo.Metadata = new DocumentMetadata() { DocumentType = docType, DocumentVersion = docVersion };
        //    generationInfo.DataContext = dataContext;
        //    generationInfo.TemplateData = File.ReadAllBytes(Path.Combine(@"Docs\templates", wordTemplateFile));
        //    generationInfo.IsDataBoundControls = useDataBoundControls;

        //    return generationInfo;
        //}
    }
}