using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp1.Model
{
    public abstract class DocumentGenerator
    {
        protected const string DocumentRootNode = "DocumentRootNode";

        protected const string DocumentNode = "Document";

        protected const string DocumentContainerPlaceHoldersNode = "DocumentContainerPlaceHolders";

        protected const string DataBoundControlsDataStoreNode = "DataBoundControlsDataStore";

        protected const string DataNode = "Data";

        protected const string DocumentTypeNodeName = "DocumentType";

        protected const string DocumentVersionNodeName = "Version";

        private DocumentGenerationInfo generationInfo;

        private readonly CustomXmlPartHelper customXmlPartHelper = new CustomXmlPartHelper(DocumentGenerationInfo.NamespaceUri);

        private readonly OpenXmlHelper openXmlHelper = new OpenXmlHelper(DocumentGenerationInfo.NamespaceUri);

        public DocumentGenerator(DocumentGenerationInfo generationInfo)
        {
            this.generationInfo = generationInfo;
        }

        protected abstract Dictionary<string, PlaceHolderType> GetPlaceHolderTagToTypeCollection();

        protected abstract void IgnorePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        protected abstract void NonRecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        protected abstract void RecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        protected abstract void ContainerPlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        protected virtual string SerializeDataContextToXml()
        {
            StringBuilder sb = new StringBuilder();

            if (generationInfo != null && generationInfo.DataContext != null)
            {
                XmlSerializer serializer = new XmlSerializer(generationInfo.DataContext.GetType());
                XmlWriterSettings writerSettings = new XmlWriterSettings();
                writerSettings.OmitXmlDeclaration = true;

                using (XmlWriter writer = XmlWriter.Create(sb, writerSettings))
                {
                    serializer.Serialize(writer, generationInfo.DataContext);
                }
            }

            return sb.ToString();
        }

        protected bool GetParentContainer(ref SdtElement parentContainer, string placeHolder)
        {
            bool isRefresh = false;
            MainDocumentPart mainDocumentPart = parentContainer.Ancestors<Document>().First().MainDocumentPart;
            KeyValuePair<string, string> nameToValue = this.customXmlPartHelper.GetNameToValueCollectionFromElementForType(mainDocumentPart, DocumentContainerPlaceHoldersNode, NodeType.Element).Where(f => f.Key.Equals(placeHolder)).FirstOrDefault();

            isRefresh = !string.IsNullOrEmpty(nameToValue.Value);

            if (isRefresh)
            {
                SdtElement parentElementFromCustomXmlPart = new SdtBlock(nameToValue.Value);
                parentContainer.Parent.ReplaceChild(parentElementFromCustomXmlPart, parentContainer);
                parentContainer = parentElementFromCustomXmlPart;
            }
            else
            {
                Dictionary<string, string> nameToValueCollection = new Dictionary<string, string>();
                nameToValueCollection.Add(placeHolder, parentContainer.OuterXml);
                this.customXmlPartHelper.SetElementFromNameToValueCollectionForType(mainDocumentPart, DocumentRootNode, DocumentContainerPlaceHoldersNode, nameToValueCollection, NodeType.Element);
            }

            return isRefresh;
        }

        protected string GetTagValue(SdtElement element, out string templateTagPart, out string tagGuidPart)
        {
            templateTagPart = string.Empty;
            tagGuidPart = string.Empty;
            Tag tag = openXmlHelper.GetTag(element);

            string fullTag = (tag == null || (tag.Val.HasValue == false)) ? string.Empty : tag.Val.Value;

            if (!string.IsNullOrEmpty(fullTag))
            {
                string[] tagParts = fullTag.Split(':');

                if (tagParts.Length == 2)
                {
                    templateTagPart = tagParts[0];
                    tagGuidPart = tagParts[1];
                }
                else if (tagParts.Length == 1)
                {
                    templateTagPart = tagParts[0];
                }
            }

            return fullTag;
        }

        protected string GetFullTagValue(string templateTagPart, string tagGuidPart)
        {
            return templateTagPart + ":" + tagGuidPart;
        }

        protected void SaveDataToDataBoundControlsDataStore(MainDocumentPart mainDocumentPart)
        {
            string dataContextAsXml = this.SerializeDataContextToXml();
            Dictionary<string, string> nameToValueCollection = new Dictionary<string, string>();
            nameToValueCollection.Add(DataNode, dataContextAsXml);
            this.customXmlPartHelper.SetElementFromNameToValueCollectionForType(mainDocumentPart, DocumentRootNode, DataBoundControlsDataStoreNode, nameToValueCollection, NodeType.Element);
        }

        protected void SetDataBinding(string xPath, SdtElement element)
        {
            element.SdtProperties.RemoveAllChildren<DataBinding>();
            DataBinding dataBinding = new DataBinding() { XPath = xPath, StoreItemId = new StringValue(this.customXmlPartHelper.customXmlPartCore.GetStoreItemId(element.Ancestors<Document>().First().MainDocumentPart)) };
            element.SdtProperties.Append(dataBinding);
        }
        protected object GetDataContext()
        {
            return generationInfo != null ? this.generationInfo.DataContext : null;
        }

        protected void SetTagValue(SdtElement element, string fullTagValue)
        {
            // Set the tag for the content control
            if (!string.IsNullOrEmpty(fullTagValue))
            {
                this.openXmlHelper.SetTagValue(element, fullTagValue);
            }
        }

        protected void SetContentOfContentControl(SdtElement element, string content)
        {
            // Set text without data binding
            this.openXmlHelper.SetContentOfContentControl(element, content);
        }

        protected void SetContentInPlaceholders(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (IsContentControl(openXmlElementDataContext))
            {
                string templateTagPart = string.Empty;
                string tagGuidPart = string.Empty;
                SdtElement element = openXmlElementDataContext.Element as SdtElement;
                GetTagValue(element, out templateTagPart, out tagGuidPart);

                if (this.generationInfo.PlaceHolderTagToTypeCollection.ContainsKey(templateTagPart))
                {
                    this.OnPlaceHolderFound(openXmlElementDataContext);
                }
                else
                {
                    this.PopulateOtherOpenXmlElements(openXmlElementDataContext);
                }
            }
            else
            {
                this.PopulateOtherOpenXmlElements(openXmlElementDataContext);
            }
        }

        protected SdtElement CloneElementAndSetContentInPlaceholders(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null)
            {
                throw new ArgumentNullException("openXmlElementDataContext");
            }

            if (openXmlElementDataContext.Element == null)
            {
                throw new ArgumentNullException("openXmlElementDataContext.element");
            }

            SdtElement clonedSdtElement = null;

            if (openXmlElementDataContext.Element.Parent != null && openXmlElementDataContext.Element.Parent is Paragraph)
            {
                Paragraph clonedPara = openXmlElementDataContext.Element.Parent.InsertBeforeSelf(openXmlElementDataContext.Element.Parent.CloneNode(true) as Paragraph);
                clonedSdtElement = clonedPara.Descendants<SdtElement>().First();
            }
            else
            {
                clonedSdtElement = openXmlElementDataContext.Element.InsertBeforeSelf(openXmlElementDataContext.Element.CloneNode(true) as SdtElement);
            }

            foreach (var v in clonedSdtElement.Elements())
            {
                this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = v, DataContext = openXmlElementDataContext.DataContext });
            }

            return clonedSdtElement;
        }

        protected void SetDocumentProperties(MainDocumentPart mainDocumentPart, DocumentMetadata docProperties)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            if (docProperties == null)
            {
                throw new ArgumentNullException("docProperties");
            }

            Dictionary<string, string> idtoValues = new Dictionary<string, string>();
            idtoValues.Add(DocumentTypeNodeName, string.IsNullOrEmpty(docProperties.DocumentType) ? string.Empty : docProperties.DocumentType);
            idtoValues.Add(DocumentVersionNodeName, string.IsNullOrEmpty(docProperties.DocumentVersion) ? string.Empty : docProperties.DocumentVersion);
            this.customXmlPartHelper.SetElementFromNameToValueCollectionForType(mainDocumentPart, DocumentRootNode, DocumentNode, idtoValues, NodeType.Attribute);
        }

        protected bool IsTemplateTagEqual(SdtElement element, string placeholderName)
        {
            if (element == null)
            {
                throw new ArgumentNullException("element");
            }

            if (placeholderName == null)
            {
                throw new ArgumentNullException("placeholderName");
            }

            string templateTagPart = string.Empty;
            string tagGuidPart = string.Empty;
            GetTagValue(element, out templateTagPart, out tagGuidPart);
            return placeholderName.Equals(templateTagPart);
        }

        public byte[] GenerateDocument()
        {
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

            return SetContentInPlaceholders();
        }

        private byte[] SetContentInPlaceholders()
        {
            byte[] output = null;

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(this.generationInfo.TemplateData, 0, this.generationInfo.TemplateData.Length);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(ms, true))
                {
                    wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);
                    MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart;
                    Document document = mainDocumentPart.Document;

                    if (this.generationInfo.Metadata != null)
                    {
                        SetDocumentProperties(mainDocumentPart, this.generationInfo.Metadata);
                    }

                    if (this.generationInfo.IsDataBoundControls)
                    {
                        SaveDataToDataBoundControlsDataStore(mainDocumentPart);
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

                    this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = document, DataContext = this.generationInfo.DataContext });

                    this.openXmlHelper.EnsureUniqueContentControlIdsForMainDocumentPart(mainDocumentPart);

                    document.Save();
                }

                ms.Position = 0;
                output = new byte[ms.Length];
                ms.Read(output, 0, output.Length);
            }

            return output;
        }

        private void PopulateOtherOpenXmlElements(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext.Element is OpenXmlCompositeElement && openXmlElementDataContext.Element.HasChildren)
            {
                List<OpenXmlElement> elements = openXmlElementDataContext.Element.Elements().ToList();

                foreach (var element in elements)
                {
                    if (element is OpenXmlCompositeElement)
                    {
                        this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = element, DataContext = openXmlElementDataContext.DataContext });
                    }
                }
            }
        }

        private bool IsContentControl(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null || openXmlElementDataContext.Element == null)
            {
                return false;
            }

            return openXmlElementDataContext.Element is SdtBlock || openXmlElementDataContext.Element is SdtRun || openXmlElementDataContext.Element is SdtRow || openXmlElementDataContext.Element is SdtCell;
        }

        private void OnPlaceHolderFound(OpenXmlElementDataContext openXmlElementDataContext)
        {
            string templateTagPart = string.Empty;
            string tagGuidPart = string.Empty;
            SdtElement element = openXmlElementDataContext.Element as SdtElement;
            GetTagValue(element, out templateTagPart, out tagGuidPart);

            if (this.generationInfo.PlaceHolderTagToTypeCollection.ContainsKey(templateTagPart))
            {
                switch (this.generationInfo.PlaceHolderTagToTypeCollection[templateTagPart])
                {
                    case PlaceHolderType.None:
                        break;
                    case PlaceHolderType.NonRecursive:
                        this.NonRecursivePlaceholderFound(templateTagPart, openXmlElementDataContext);
                        break;
                    case PlaceHolderType.Recursive:
                        this.RecursivePlaceholderFound(templateTagPart, openXmlElementDataContext);
                        break;
                    case PlaceHolderType.Ignore:
                        this.IgnorePlaceholderFound(templateTagPart, openXmlElementDataContext);
                        break;
                    case PlaceHolderType.Container:
                        this.ContainerPlaceholderFound(templateTagPart, openXmlElementDataContext);
                        break;
                }
            }
        }

    }
}
