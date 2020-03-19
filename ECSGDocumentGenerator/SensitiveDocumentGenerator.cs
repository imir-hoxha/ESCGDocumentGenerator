using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ECSGDocumentGenerator.Domain;
using ECSGDocumentGenerator.Model;
using System;
using System.Collections.Generic;
using System.IO;

namespace ECSGDocumentGenerator
{
    public class SensitiveDocumentGenerator : DocumentGenerator
    {
        public static Dictionary<string, PlaceHolderType> PlaceHolderTagToTypeCollection { get; set; }
        private DocumentGenerationInfo generationInfo;
        public SensitiveDocumentGenerator(DocumentGenerationInfo generationInfo) : base(generationInfo)
        {

            this.generationInfo = generationInfo;
        }
        protected override Dictionary<string, PlaceHolderType> GetPlaceHolderTagToTypeCollection()
        {
            Dictionary<string, PlaceHolderType> placeHolderTagToTypeCollection = new Dictionary<string, PlaceHolderType>
            {

                // Handle container placeholders            
                { DocumentPlaceHolders.PlaceholderContainerA, PlaceHolderType.Container },

                // Handle non recursive placeholders
                { DocumentPlaceHolders.PlaceholderMemberState, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceHolderLeadDG, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceHolderTitle, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceHolderInfringementReference, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceHolderReasonForSensitivity, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceHolderDecisionType, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderDecissionMS, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceHolderPolicyContext, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceHolderLineToTake, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceHolderDGCaseHandler, PlaceHolderType.NonRecursive }

                //{ DocumentPlaceHolders.PlaceHolderLeadDG, PlaceHolderType.NonRecursive }
            };

            return placeHolderTagToTypeCollection;
        }


        public byte[] GenerateAndMergeTemplates(string headerTemplateFile, string bodyTemplateFile, Content[] dataContext)
        {
            byte[] output = null;

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(this.generationInfo.TemplateData, 0, this.generationInfo.TemplateData.Length);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(ms, true))
                {
                    wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);
                    MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart;
                    Document mainDocument = mainDocumentPart.Document;

                    if (this.generationInfo.Metadata != null)
                    {
                        SetDocumentProperties(mainDocumentPart, this.generationInfo.Metadata);
                    }

                    if (this.generationInfo.IsDataBoundControls)
                    {
                        SaveDataToDataBoundControlsDataStore(mainDocumentPart);

                    }

                    if (this.generationInfo == null)
                    {
                        throw new ArgumentNullException("generationInfo");
                    }

                    if (this.generationInfo.TemplateData == null)
                    {
                        throw new ArgumentNullException("templateData");
                    }

                    this.generationInfo.PlaceHolderTagToTypeCollection = this.GetPlaceHolderTagToTypeCollection();

                    if (this.generationInfo.PlaceHolderTagToTypeCollection == null)
                    {
                        throw new ArgumentNullException("PlaceHolderTagToTypeCollection");
                    }

                    foreach (HeaderPart part in mainDocumentPart.HeaderParts)
                    {
                        this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = part.Header, DataContext = this.generationInfo.DataContext });
                        part.Header.Save();
                    }

                    foreach (FooterPart part in mainDocumentPart.FooterParts)
                    {
                        this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = part.Footer, DataContext = this.generationInfo.DataContext });
                        part.Footer.Save();
                    }

                    int counter = 0;
                    //for (int i = 0; i < content.Length; i++)
                    //{
                    //    Console.WriteLine(i + ") " + content[i].hit.reasonForSensitivity);
                    //}


                    foreach (var repo in dataContext)
                    {
                        Console.WriteLine("------------------------> " + repo.hit.article259 + " " + repo.hit.caseTitle);

                        using (FileStream fileStream = File.Open(bodyTemplateFile, FileMode.Open))
                        {
                            using (var memoryStream = new MemoryStream())
                            {
                                fileStream.CopyTo(memoryStream);

                                using (WordprocessingDocument chunkDocument = WordprocessingDocument.Open(memoryStream, true))
                                {
                                    MainDocumentPart mainDocPart = chunkDocument.MainDocumentPart;
                                    Document document = mainDocPart.Document;
                                    this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = document, DataContext = repo }); //here "DataContext = repo" should be replaced with "DataContext = this.generationInfo.DataContext"
                                    document.Save();
                                }

                                memoryStream.Seek(0, SeekOrigin.Begin);

                                string altChunkId = "AltChunkId" + Guid.NewGuid();
                                AlternativeFormatImportPart chunk = wordDocument.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

                                chunk.FeedData(memoryStream);

                                AltChunk altChunk = new AltChunk();
                                altChunk.Id = altChunkId;
                                counter++;
                                if (counter > 0)
                                {
                                    mainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                                }

                                mainDocumentPart.Document.Body.AppendChild(altChunk);
                                mainDocumentPart.Document.Save();

                            }
                        }
                    }


                    //this.openXmlHelper.EnsureUniqueContentControlIdsForMainDocumentPart(mainDocumentPart);
                }

                ms.Position = 0;
                output = new byte[ms.Length];
                ms.Read(output, 0, output.Length);

            }

            return output;
            //using (FileStream fs = File.Open(headerTemplateFile, FileMode.Open))
            //{
            //    using (WordprocessingDocument myDoc = WordprocessingDocument.Open(fs, true))
            //    {
            //        MainDocumentPart mainPart = myDoc.MainDocumentPart;

            //        int counter = 0;
            //        //foreach (var repo in generationInfo.DataContext)
            //        //{

            //        using (FileStream fileStream = File.Open(bodyTemplateFile, FileMode.Open))
            //        {
            //            using (var memoryStream = new MemoryStream())
            //            {
            //                fileStream.CopyTo(memoryStream);

            //                using (WordprocessingDocument chunkDocument = WordprocessingDocument.Open(memoryStream, true))
            //                {
            //                    MainDocumentPart mainDocPart = chunkDocument.MainDocumentPart;
            //                    Document document = mainDocPart.Document;
            //                    this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = document, DataContext = this.generationInfo.DataContext });

            //                }

            //                memoryStream.Seek(0, SeekOrigin.Begin);
            //                // Create an AlternativeFormatImportPart from the MemoryStream.
            //                string altChunkId = "AltChunkId" + Guid.NewGuid();
            //                AlternativeFormatImportPart chunk = myDoc.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

            //                chunk.FeedData(memoryStream);

            //                AltChunk altChunk = new AltChunk();
            //                altChunk.Id = altChunkId;
            //                counter++;
            //                if (counter > 0)
            //                {
            //                    mainPart.Document.Body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
            //                }

            //                mainPart.Document.Body.AppendChild(altChunk);
            //                mainPart.Document.Save();

            //            }

            //            //}
            //        }
            //    }
            //}
        }


        protected override void ContainerPlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null || openXmlElementDataContext.Element == null || openXmlElementDataContext.DataContext == null)
            {
                return;
            }

            string tagPlaceHolderValue = string.Empty;
            string tagGuidPart = string.Empty;
            GetTagValue(openXmlElementDataContext.Element as SdtElement, out tagPlaceHolderValue, out tagGuidPart);

            string tagValue = string.Empty;
            string content = string.Empty;

            switch (tagPlaceHolderValue)
            {
                case DocumentPlaceHolders.PlaceholderContainerA:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();

                    if (!string.IsNullOrEmpty(tagValue))
                    {
                        SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
                    }

                    foreach (var v in openXmlElementDataContext.Element.Elements())
                    {
                        SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = v, DataContext = openXmlElementDataContext.DataContext });
                    }

                    break;
            }
        }

        protected override void RecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
        {
            throw new NotImplementedException();

        }

        protected override void NonRecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null || openXmlElementDataContext.Element == null || openXmlElementDataContext.DataContext == null)
            {
                return;
            }

            string tagPlaceHolderValue = string.Empty;
            string tagGuidPart = string.Empty;
            GetTagValue(openXmlElementDataContext.Element as SdtElement, out tagPlaceHolderValue, out tagGuidPart);

            string tagValue = string.Empty;
            string content = string.Empty;

            switch (tagPlaceHolderValue)
            {
                case DocumentPlaceHolders.PlaceholderMemberState:
                    tagValue = openXmlElementDataContext.DataContext.ToString();
                    content = openXmlElementDataContext.DataContext.hit.memberState.ToString();
                    break;
                case DocumentPlaceHolders.PlaceHolderLeadDG:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.leadDg.text;
                    break;
                case DocumentPlaceHolders.PlaceHolderTitle:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.caseTitle;
                    break;
                case DocumentPlaceHolders.PlaceHolderInfringementReference:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.caseTitle;
                    break;
                    //case DocumentPlaceHolders.PlaceholderNonRecursiveE:
                    //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    //    content = openXmlElementDataContext.DataContext.hit.C14;
                    //    break;
                    //case DocumentPlaceHolders.PlaceholderNonRecursiveF:
                    //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    //    content = openXmlElementDataContext.DataContext.hit.P1;
                    //    break;
                    //case DocumentPlaceHolders.PlaceholderNonRecursiveH:
                    //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    //    content = openXmlElementDataContext.DataContext.hit.P22;
                    //    break;
                    //case DocumentPlaceHolders.PlaceholderNonRecursiveI:
                    //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    //    content = openXmlElementDataContext.DataContext.hit.P1;
                    //    break;
                    //case DocumentPlaceHolders.PlaceholderNonRecursiveJ:
                    //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    //    content = openXmlElementDataContext.DataContext.hit.P1;
                    //    break;
                    //case DocumentPlaceHolders.PlaceholderNonRecursiveK:
                    //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    //    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P22;
                    //    break;
                    //case DocumentPlaceHolders.PlaceholderNonRecursiveL:
                    //    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    //    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).ED1;
                    //    break;
                    //case DocumentPlaceHolders.PlaceholderNonRecursiveM:
                    //    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    //    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C9;
                    //    break;
                    //case DocumentPlaceHolders.PlaceholderNonRecursiveN:
                    //    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    //    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C28;
                    //    break;
            }

            if (!string.IsNullOrEmpty(tagValue))
            {
                SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
            }

            this.SetContentOfContentControl(openXmlElementDataContext.Element as SdtElement, content);
        }

        //public static string GetFullTagValue(string templateTagPart, string tagGuidPart)
        //{
        //    return templateTagPart + ":" + tagGuidPart;
        //}

        //public static void SetTagValue(SdtElement element, string tagValue)
        //{
        //    Tag tag = GetTag(element);
        //    tag.Val.Value = tagValue;
        //}

        //public static string GetTagValue(SdtElement element, out string templateTagPart, out string tagGuidPart)
        //{
        //    OpenXmlHelper openXmlHelper = new OpenXmlHelper(DocumentGenerationInfo.NamespaceUri);
        //    templateTagPart = string.Empty;
        //    tagGuidPart = string.Empty;
        //    Tag tag = GetTag(element);

        //    string fullTag = (tag == null || (tag.Val.HasValue == false)) ? string.Empty : tag.Val.Value;

        //    if (!string.IsNullOrEmpty(fullTag))
        //    {
        //        string[] tagParts = fullTag.Split(':');

        //        if (tagParts.Length == 2)
        //        {
        //            templateTagPart = tagParts[0];
        //            tagGuidPart = tagParts[1];
        //        }
        //        else if (tagParts.Length == 1)
        //        {
        //            templateTagPart = tagParts[0];
        //        }
        //    }

        //    return fullTag;
        //}

        //public static Tag GetTag(SdtElement element)
        //{
        //    if (element == null)
        //        throw new ArgumentNullException("element");

        //    return element.SdtProperties.Elements<Tag>().FirstOrDefault();
        //}
    }
}
