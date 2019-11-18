using ConsoleApp1.Domain;
using ConsoleApp1.Model;
using Newtonsoft.Json;
using System;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            GenerateDocumentUsingDocGenerator();
        }

        private static RefreshableDocumentGenerator GenerateDocumentUsingDocGenerator()
        {

            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SomeDocDocumentGenerator", "1.0", GetDataContext(), "DG Non sensitive report-Template.docx", false);

            RefreshableDocumentGenerator refreshableDocumentGenerator =
                new RefreshableDocumentGenerator(generationInfo);
            byte[] result = refreshableDocumentGenerator.GenerateDocument();
            WriteOutputToFile("NonSensitiveGeneratedDocument.docx", "DG Non sensitive report-Template.docx", result);
            return refreshableDocumentGenerator;
        }

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

        private static Report GetDataContext()
        {
            var filePath = Path.Combine(@"Docs\templates", "non-sensitive-data.json");
            Report repo = null;
            using (StreamReader r = new StreamReader(filePath))
            {
                string json = r.ReadToEnd();
                Report item = JsonConvert.DeserializeObject<Report>(json);

                repo = new Report()
                {
                    C1 = item.C1,
                    C24 = item.C24,
                    C2 = item.C2,
                    C18 = item.C18,
                    C14 = item.C14,
                    P1 = item.P1,
                    P22 = item.P22,
                    ED1 = item.ED1,
                    C9 = item.C9,
                    C28 = item.C28
                };
            }

            return repo;
        }

        private static DocumentGenerationInfo GetDocumentGenerationInfo(string docType, string docVersion, object dataContext, string fileName, bool useDataBoundControls)
        {
            DocumentGenerationInfo generationInfo = new DocumentGenerationInfo();
            generationInfo.Metadata = new DocumentMetadata() { DocumentType = docType, DocumentVersion = docVersion };
            generationInfo.DataContext = dataContext;
            generationInfo.TemplateData = File.ReadAllBytes(Path.Combine(@"Docs\templates", fileName));
            generationInfo.IsDataBoundControls = useDataBoundControls;

            return generationInfo;
        }

    }

}
