﻿using ConsoleApp1.Domain;
using ECSGDocumentGenerator.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ECSGDocumentGenerator.Domain;

namespace ECSGDocumentGenerator
{
    public class MyDocGenerator2 : DocumentGenerator
    {
        public static Dictionary<string, PlaceHolderType> PlaceHolderTagToTypeCollection { get; set; }
        #region PlaceHolders
        // Content Control Tags
        //public const string PlaceholderIgnoreA = "PlaceholderIgnoreA";
        //public const string PlaceholderIgnoreB = "PlaceholderIgnoreB";

        //public const string PlaceholderContainerA = "PlaceholderContainerA";

        //public const string PlaceholderRecursiveA = "PlaceholderRecursiveA";
        //public const string PlaceholderRecursiveB = "PlaceholderRecursiveB";

        //public const string PlaceholderNonRecursiveA = "PlaceholderNonRecursiveA";
        //public const string PlaceholderNonRecursiveB = "PlaceholderNonRecursiveB";
        //public const string PlaceholderNonRecursiveC = "PlaceholderNonRecursiveC";
        //public const string PlaceholderNonRecursiveD = "PlaceholderNonRecursiveD";
        //public const string PlaceholderNonRecursiveE = "PlaceholderNonRecursiveE";

        //public const string PlaceholderNonRecursiveF = "PlaceholderNonRecursiveF";
        //public const string PlaceholderNonRecursiveG = "PlaceholderNonRecursiveG";
        //public const string PlaceholderNonRecursiveH = "PlaceholderNonRecursiveH";
        //public const string PlaceholderNonRecursiveI = "PlaceholderNonRecursiveI";
        //public const string PlaceholderNonRecursiveJ = "PlaceholderNonRecursiveJ";

        //public const string PlaceholderNonRecursiveK = "PlaceholderNonRecursiveK";
        //public const string PlaceholderNonRecursiveL = "PlaceholderNonRecursiveL";
        //public const string PlaceholderNonRecursiveM = "PlaceholderNonRecursiveM";
        //public const string PlaceholderNonRecursiveN = "PlaceholderNonRecursiveN";
        #endregion

        //public DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SomeDocDocumentGenerator", "1.0", GetDataContext(), "body.docx", false);
        public MyDocGenerator2(DocumentGenerationInfo generationInfo) : base(generationInfo) { }
        protected override Dictionary<string, PlaceHolderType> GetPlaceHolderTagToTypeCollection()
        {
            Dictionary<string, PlaceHolderType> placeHolderTagToTypeCollection = new Dictionary<string, PlaceHolderType>
            {

                // Handle container placeholders            
                { DocumentPlaceHolders.PlaceholderContainerA, PlaceHolderType.Container },

                // Handle non recursive placeholders
                { DocumentPlaceHolders.PlaceholderNonRecursiveA, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveB, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveC, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveD, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveE, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveF, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveG, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveH, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveI, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveJ, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveK, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveL, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveM, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveN, PlaceHolderType.NonRecursive }
            };

            return placeHolderTagToTypeCollection;
        }


<<<<<<< HEAD
        public void GenerateAndMergeTemplates(string headerTemplateFile, string bodyTemplateFile, Content dataContext)
=======
        public void GenerateAndMergeTemplates(string headerTemplateFile, string bodyTemplateFile, SensitiveReport dataContext)
>>>>>>> 3890695f2ae98b5ec3af60a4f929077de2d09acb
        {

            using (FileStream fs = File.Open(headerTemplateFile, FileMode.Open))
            {
                using (WordprocessingDocument myDoc = WordprocessingDocument.Open(fs, true))
                {
                    MainDocumentPart mainPart = myDoc.MainDocumentPart;
                    DocumentGenerationInfo generationInfo = new DocumentGenerationInfo();
                    generationInfo.DataContext = dataContext;
                    int counter = 0;
<<<<<<< HEAD
                    //foreach (var repo in generationInfo.DataContext)
                    //{
=======
                    foreach (var repo in generationInfo.DataContext.content)
                    {
>>>>>>> 3890695f2ae98b5ec3af60a4f929077de2d09acb

                        using (FileStream fileStream = File.Open(bodyTemplateFile, FileMode.Open))
                        {
                            using (var memoryStream = new MemoryStream())
                            {
                                fileStream.CopyTo(memoryStream);

                                using (WordprocessingDocument chunkDocument = WordprocessingDocument.Open(memoryStream, true))
                                {
                                    MainDocumentPart mainDocPart = chunkDocument.MainDocumentPart;
                                    Document document = mainDocPart.Document;
                                    this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = document, DataContext = generationInfo.DataContext });

                                }

                                memoryStream.Seek(0, SeekOrigin.Begin);
                                // Create an AlternativeFormatImportPart from the MemoryStream.
                                string altChunkId = "AltChunkId" + Guid.NewGuid();
                                AlternativeFormatImportPart chunk = myDoc.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

                                chunk.FeedData(memoryStream);

                                AltChunk altChunk = new AltChunk();
                                altChunk.Id = altChunkId;
                                counter++;
                                if (counter > 0)
                                {
                                    mainPart.Document.Body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                                }

                                mainPart.Document.Body.AppendChild(altChunk);
                                mainPart.Document.Save();

                            }

                        }
                    //}
                }
            }
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
<<<<<<< HEAD
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
=======
                    tagValue = (openXmlElementDataContext.DataContext as NonSensitiveReport).Id.ToString();
>>>>>>> 3890695f2ae98b5ec3af60a4f929077de2d09acb

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
                case DocumentPlaceHolders.PlaceholderNonRecursiveA:
<<<<<<< HEAD
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.memberState;
=======
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C1;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveB:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C24;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveC:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C2;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveD:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C18;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveE:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C14;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveF:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P1;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveH:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P22;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveI:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P1;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveJ:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P1;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveK:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P22;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveL:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).ED1;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveM:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C9;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveN:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C28;
>>>>>>> 3890695f2ae98b5ec3af60a4f929077de2d09acb
                    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveB:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C24;
                //    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveC:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C2;
                //    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveD:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C18;
                //    break;
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
                //    content = openXmlElementDataContext.DataContext.hit.P22;
                //    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveL:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.ED1;
                //    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveM:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C9;
                //    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveN:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C28;
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
