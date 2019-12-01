using ECSGDocumentGenerator;
using ECSGDocumentGenerator.Model;
using Newtonsoft.Json;
using System;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {

        static void Main(string[] args)
        {
            var items = GetDataContext();
            foreach (var item in items)
            {
                Console.WriteLine(item.Hit.Article259);
            }
            //HttpWebResponse resp = CallWebService.GetSensitiveJsonData("http://s-themis-acc.net1.cec.eu.int:8044/doSearch?size=10&page=0&sort=ASC");

            //if (resp.StatusCode == HttpStatusCode.OK)
            //{
            //    WebResponse response = resp;
            //    Stream responseStream = response.GetResponseStream();
            //    StreamReader streamReader = new StreamReader(responseStream);
            //    string responseData = streamReader.ReadToEnd();
            //    response.Close();

            //    //SensitiveReport report = GetDataContext(responseData);
            //    //Content[] content = report.content;
            //    //Console.WriteLine(content[0].hit.reasonForSensitivity);
            //    //for (int i = 0; i < content.Length; i++)
            //    //{
            //    //    Console.WriteLine(i + ") " + content[i].hit.reasonForSensitivity);
            //    //}
            //    //Console.WriteLine("---------------------------");

            //    ////var ct = content.Select(new )

            //    //foreach (Content item in content)
            //    //{
            //    //    Console.WriteLine(item.hit.authorOfTheSensitiveSummary + " " + item.hit.caseSensitivity + " " + item.hit.reasonForSensitivity);

            //    //}

            //    GenerateDocument(responseData, "header.docx");
            //}
            //else
            //{
            //    Console.WriteLine(resp.StatusCode);
            //}


            //string d = CallWebService.makePostRequestUsingWebClient("http://s-themis-acc.net1.cec.eu.int:8044/doSearch?size=10&page=0&sort=ASC");
            //GenerateDocumentUsingDocGenerator();

            //GeneratedClassA cls = new GeneratedClassA();
            //cls.CreatePackage(@"C:\Dev\Doc3.docx");
            //MyDocGenerator.GetPlaceHolderTagToTypeCollection();

            //string jsonFilePath = Path.Combine(@"Docs\templates", "non-sensitive-data.json");
            //string jsonSensitiveFilePath = Path.Combine(@"Docs\templates", "sensitive-data.json");

            //List<SensitiveReport> sensitiveRepData = GetDataContextSensitiveData(jsonSensitiveFilePath);
            //foreach (var item in sensitiveRepData[0].content)
            //{
            //    NewMethod(item.hit);
            //}


        }

        private static void GenerateDocument(string jsonData, string fileTemplate)
        {
            //string jsonFilePath = Path.Combine(@"Docs\templates", jsonData);
            string jsonFilePath = jsonData;
            var dataArray = GetDataContext();
            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SomeDocDocumentGenerator", "1.0", GetDataContext(jsonData), dataArray, fileTemplate, false);
            SensitiveDocumentGenerator myDocGen = new SensitiveDocumentGenerator(generationInfo);
            //(string headerTemplateFile, string bodyTemplateFile, List<Report> dataContext)
            string bodyDoc = Path.Combine(@"Docs\templates", "body.docx");
            byte[] result = myDocGen.MergeAndGenerateTemplate(bodyDoc);
            WriteOutputToFile("NonSensitiveGeneratedDocument-v2.docx", "body.docx", result);
            ////GenerateSensitiveDocumentUsingDocGenerator();
        }

        //private static void NewMethod(Hit sensitiveRepData)
        //{
        //    string tagPlaceHolderValue = "PlaceholderNonRecursiveA";
        //    switch (tagPlaceHolderValue)
        //    {
        //        case "PlaceholderNonRecursiveA":
        //            Console.WriteLine(sensitiveRepData.Id.ToString());
        //            //Console.WriteLine(((sensitiveRepData[0].content[0].hit) as Hit).leadDg.text);
        //            break;
        //        default:
        //            break;
        //    }
        //}

        //private static RefreshableDocumentGenerator GenerateDocumentUsingDocGenerator()
        //{
        //    string jsonFilePath = Path.Combine(@"Docs\templates", "non-sensitive-data.json");
        //    DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SomeDocDocumentGenerator", "1.0", GetDataContext(jsonFilePath), "body.docx", false);

        //    RefreshableDocumentGenerator refreshableDocumentGenerator = new RefreshableDocumentGenerator(generationInfo);
        //    byte[] result = refreshableDocumentGenerator.GenerateDocument();
        //    WriteOutputToFile("NonSensitiveGeneratedDocument.docx", "DG Non sensitive report-Template.docx", result);
        //    return refreshableDocumentGenerator;
        //}

        //private static SensitiveDocumentGenerator GenerateSensitiveDocumentUsingDocGenerator()
        //{
        //    //string jsonFilePath = Path.Combine(@"Docs\templates", "sensitive-data.json");
        //    DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SensitiveDocument", "1.0", GetDataContextSensitiveData(""), "DG Sensitive report-Template.docx", false);

        //    SensitiveDocumentGenerator refreshableDocumentGenerator = new SensitiveDocumentGenerator(generationInfo);
        //    byte[] result = refreshableDocumentGenerator.GenerateDocument();
        //    WriteOutputToFile("NewSensitiveDocument.docx", "DG Sensitive report-Template.docx", result);
        //    return refreshableDocumentGenerator;
        //}

        private static void WriteOutputToFile(string fileName, string templateName, byte[] fileContents)
        {
            ConsoleColor consoleColor = Console.ForegroundColor;

            if (fileContents != null)
            {
                File.WriteAllBytes(Path.Combine("Docs", fileName), fileContents);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Generation succeeded for template({templateName}) --> {fileName}");
                Console.WriteLine();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Generation failed for template({templateName}) --> {fileName}");
            }

            Console.ForegroundColor = consoleColor;
        }

        //private static List<SensitiveReport> GetDataContextSensitiveData(string jsonFilePath)
        //{
        //    List<SensitiveReport> reps = new List<SensitiveReport>();
        //    using (StreamReader r = new StreamReader(jsonFilePath))
        //    {
        //        string json = r.ReadToEnd();
        //        var item = JsonConvert.DeserializeObject<SensitiveReport>(json);

        //        for (int i = 0; i < item.content.Length; i++)
        //        {
        //            reps.Add(new SensitiveReport()
        //            {
        //                content = item.content,
        //                pageable = item.pageable,
        //                totalElements = item.totalElements,
        //                last = item.last,
        //                size = item.size,
        //                number = item.number,
        //                sort = item.sort,
        //                first = item.first,
        //                numberOfElements = item.numberOfElements

        //            });
        //        }

        //    }

        //    return reps;
        //}

        private static Content GetDataContext(string jsonData)
        {
            if (string.IsNullOrEmpty(jsonData))
            {
                throw new ArgumentNullException("jsonData");
            }

            var data = JsonConvert.DeserializeObject<Content>(jsonData);

            return data;
        }

        private static Content[] GetDataContext()
        {
            var filePath = Path.Combine(@"Docs\templates", "sensitive-data.json");
            Content[] repo = null;
            using (StreamReader r = new StreamReader(filePath))
            {
                string json = r.ReadToEnd();
                repo = JsonConvert.DeserializeObject<SensitiveReport>(json).Content;
            }

            return repo;
        }

        private static DocumentGenerationInfo GetDocumentGenerationInfo(string docType, string docVersion, Content dataContext, Content[] data, string wordTemplateFile, bool useDataBoundControls)
        {
            DocumentGenerationInfo generationInfo = new DocumentGenerationInfo();
            generationInfo.Metadata = new DocumentMetadata() { DocumentType = docType, DocumentVersion = docVersion };
            generationInfo.DataContext = dataContext;
            generationInfo.Data = data;
            //generationInfo.Contents = new Content[2];
            generationInfo.TemplateData = File.ReadAllBytes(Path.Combine(@"Docs\templates", wordTemplateFile));
            generationInfo.IsDataBoundControls = useDataBoundControls;

            return generationInfo;
        }

    }

}
//string fileName1 = @"C:\Dev\Destination.docx";
//string fileName2 = @"C:\Dev\body2.docx";
//string testFile = @"C:\Dev\body.docx";
////File.Delete(fileName1);
////File.Copy(testFile, fileName1);
//using (WordprocessingDocument myDoc = WordprocessingDocument.Open(fileName1, true))
//{
//    string altChunkId = "AltChunkId1";
//    MainDocumentPart mainPart = myDoc.MainDocumentPart;
//    AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);
//    using (FileStream fileStream = File.Open(fileName2, FileMode.Open))
//        chunk.FeedData(fileStream);
//    AltChunk altChunk = new AltChunk();
//    altChunk.Id = altChunkId;
//    mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>().Last());
//    mainPart.Document.Save();
//}

//string altChunkId = "AltChunkId" + 0;
//    AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);
//    using (FileStream fileStream = File.Open(@filepaths[0], FileMode.Open))
//    {
//        chunk.FeedData(fileStream);
//    }
//    AltChunk altChunk = new AltChunk();
//    altChunk.Id = altChunkId;

//Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00BE27E7", RsidRunAdditionDefault = "00BE27E7" };

//Run run2 = new Run();
//Break break1 = new Break() { Type = BreakValues.Page };

//run2.Append(break1);
//paragraph2.Append(run2);
//mainPart.Document.Body.Append(paragraph2);
//mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>().Last());
//    mainPart.Document.Save();
//    myDoc.Close();

//foreach (var bodyDescendands in mainDocPart.Document.Descendants<Body>())
//{


//    foreach (var bodyChildElms in bodyDescendands.ChildElements)
//    {
//        if (IsContentControl(bodyChildElms))
//        {
//            Console.WriteLine("ContentControl: " + bodyChildElms.LocalName + " " + bodyChildElms.GetType());
//            if (bodyChildElms is OpenXmlCompositeElement && bodyChildElms.HasChildren)
//            {
//                List<OpenXmlElement> elements = bodyChildElms.Elements().ToList();

//                foreach (var element in elements)
//                {

//                    if (element is OpenXmlCompositeElement)
//                    {
//                        SdtElement el = element as SdtElement;
//                        string templateTagPart = string.Empty;
//                        string tagGuidPart = string.Empty;

//                        GetTagValue(el, out templateTagPart, out tagGuidPart);
//                        //this.SetContentInPlaceholders(new OpenXmlElementDataContext()
//                        //{
//                        //    Element = element,
//                        //    DataContext = openXmlElementDataContext.DataContext
//                        //});
//                    }
//                }
//            }

//        }


//        //Console.WriteLine(bodyChildElms.LocalName + "    -->  " + bodyChildElms.GetType() + "    > " + bodyChildElms.HasChildren);
//        //foreach (var el2 in bodyChildElms.ChildElements)
//        //{
//        //    Console.WriteLine(" - " + el2.LocalName + "   " + el2.GetType());
//        //}
//    }
//}
//foreach(var xp in mainDocPart.CustomXmlParts)
//{
//    Console.WriteLine(xp.RootElement);
//}