using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;

namespace ECSGDocumentGenerator.Model
{
    public class CustomXmlPartCore
    {
        public readonly string namespaceUri = string.Empty;

        public CustomXmlPartCore(string namespaceUri)
        {
            this.namespaceUri = namespaceUri;
        }

        public CustomXmlPart AddCustomXmlPart(MainDocumentPart mainDocumentPart, string rootElementName)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            if (string.IsNullOrEmpty(rootElementName))
            {
                throw new ArgumentNullException("rootElementName");
            }

            XName rootElementXName = XName.Get(rootElementName, this.namespaceUri);
            XElement rootElement = new XElement(rootElementXName);
            CustomXmlPart customXmlPart = mainDocumentPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            CustomXmlPropertiesPart customXmlPropertiesPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
            GenerateCustomXmlPropertiesPartContent(customXmlPropertiesPart);
            WriteElementToCustomXmlPart(customXmlPart, rootElement);

            return customXmlPart;
        }

        //public void RemoveCustomXmlPart(MainDocumentPart mainDocumentPart, CustomXmlPart customXmlPart)
        //{
        //    if (mainDocumentPart == null)
        //    {
        //        throw new ArgumentNullException("mainDocumentPart");
        //    }

        //    if (customXmlPart != null)
        //    {
        //        RemoveCustomXmlParts(mainDocumentPart, new List<CustomXmlPart>(new CustomXmlPart[] { customXmlPart }));
        //    }
        //}

        //public void RemoveCustomXmlParts(OpenXmlPartContainer mainDocumentPart, IList<CustomXmlPart> customXmlParts)
        //{
        //    if (mainDocumentPart == null)
        //    {
        //        throw new ArgumentNullException("mainDocumentPart");
        //    }

        //    if (customXmlParts != null)
        //    {
        //        mainDocumentPart.DeleteParts<CustomXmlPart>(customXmlParts);
        //    }
        //}

        public CustomXmlPart GetCustomXmlPart(MainDocumentPart mainDocumentPart)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            CustomXmlPart result = null;

            foreach (CustomXmlPart part in mainDocumentPart.CustomXmlParts)
            {
                using (XmlTextReader reader = new XmlTextReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                {
                    
                    reader.MoveToContent();
                    Console.WriteLine("custmexmlpart: " + reader.LocalName + " --> " + reader.NamespaceURI);
                    bool exists = reader.NamespaceURI.Equals(this.namespaceUri);

                    if (exists)
                    {
                        result = part;
                        break;
                    }
                }
            }

            return result;
        }

        public string GetStoreItemId(MainDocumentPart mainDocumentPart)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            CustomXmlPart customXmlPart = GetCustomXmlPart(mainDocumentPart);
            CustomXmlPropertiesPart customXmlPropertiesPart = customXmlPart.CustomXmlPropertiesPart;
            return customXmlPropertiesPart.DataStoreItem.ItemId.ToString();
        }

        public XElement GetFirstElementFromCustomXmlPart(CustomXmlPart customXmlPart, string elementName)
        {
            if (customXmlPart == null)
            {
                throw new ArgumentNullException("customXmlPart");
            }

            if (string.IsNullOrEmpty(elementName))
            {
                throw new ArgumentNullException("elementName");
            }

            XDocument customPartDoc = null;

            using (XmlReader reader = XmlReader.Create(customXmlPart.GetStream(FileMode.Open, FileAccess.Read)))
            {
                customPartDoc = XDocument.Load(reader);
            }

            XElement element = null;

            if (customPartDoc != null)
            {
                XName elementXName = XName.Get(elementName, this.namespaceUri);
                element = (from e in customPartDoc.Descendants(elementXName)
                           select e).FirstOrDefault();
            }

            return element;
        }

        public void WriteElementToCustomXmlPart(CustomXmlPart customXmlPart, XElement rootElement)
        {
            if (customXmlPart == null)
            {
                throw new ArgumentNullException("customXmlPart");
            }

            if (rootElement == null)
            {
                throw new ArgumentNullException("rootElement");
            }

            using (XmlWriter writer = XmlWriter.Create(customXmlPart.GetStream(FileMode.Create, FileAccess.Write)))
            {
                rootElement.WriteTo(writer);
                writer.Flush();
            }
        }

        private void GenerateCustomXmlPropertiesPartContent(CustomXmlPropertiesPart customXmlPropertiesPart)
        {
            DataStoreItem dataStoreItem = new DataStoreItem() { ItemId = "{" + Guid.NewGuid().ToString() + "}" };
            dataStoreItem.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");
            SchemaReferences schemaReferences = new SchemaReferences();
            SchemaReference schemaReference = new SchemaReference() { Uri = namespaceUri };
            schemaReferences.Append(schemaReference);
            dataStoreItem.Append(schemaReferences);
            customXmlPropertiesPart.DataStoreItem = dataStoreItem;
        }

    }
}
