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

namespace ESCGDocumentGenertor.API.Controllers
{
    public class HitsController : ODataController
    {

        public IHttpActionResult Get()
        {
            //Content c = new Content();
            //c.Hit = new Hit() { Id = content.Id, RefId = content.RefId, Article259 = content.Article259 };
            Document d = new Document() { Id = 1, Url = "http://sp2016" };
            return Ok("OK");
        }

        public IHttpActionResult Post(Document content)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }
            //Content c = new Content();
            //c.Hit = new Hit() { Id = content.Id, RefId = content.RefId, Article259 = content.Article259};
            Document d = new Document() { Id = content.Id, Url = content.Url };
            //Hit hit = new Hit() { Id = content.Id, RefId = content.RefId, Article259 = content.Article259 };
            return Ok(d);
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