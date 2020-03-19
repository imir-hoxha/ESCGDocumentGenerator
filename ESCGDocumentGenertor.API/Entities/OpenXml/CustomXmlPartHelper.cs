using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace ECSGDocumentGenerator.Model
{
    public class CustomXmlPartHelper
    {
        public readonly CustomXmlPartCore customXmlPartCore = null;

        public CustomXmlPartHelper(string namespaceUri)
        {
            this.customXmlPartCore = new CustomXmlPartCore(namespaceUri);
        }

        public void SetElementFromNameToValueCollectionForType(MainDocumentPart mainDocumentPart, string rootElementName, string childElementName, Dictionary<string, string> nameToValueCollection, NodeType forNodeType)
        {
            if (mainDocumentPart == null)
            {
                throw new ArgumentNullException("mainDocumentPart");
            }

            if (string.IsNullOrEmpty(rootElementName))
            {
                throw new ArgumentNullException("rootElementName");
            }

            if (string.IsNullOrEmpty(childElementName))
            {
                throw new ArgumentNullException("childElementName");
            }

            if (nameToValueCollection == null)
            {
                throw new ArgumentNullException("nameToValueCollection");
            }

            XName rootElementXName = XName.Get(rootElementName, this.customXmlPartCore.namespaceUri);
            XName childElementXName = XName.Get(childElementName, this.customXmlPartCore.namespaceUri);
            XElement rootElement = new XElement(rootElementXName);
            XElement childElement = null;
            CustomXmlPart customXmlPart = this.customXmlPartCore.GetCustomXmlPart(mainDocumentPart);

            if (customXmlPart != null)
            {
                // Root element shall never be null if Custom Xml part is present
                rootElement = this.customXmlPartCore.GetFirstElementFromCustomXmlPart(customXmlPart, rootElementName);

                childElement = (from e in rootElement.Descendants(childElementXName)
                                select e).FirstOrDefault();

                if (childElement != null)
                {
                    foreach (KeyValuePair<string, string> idToValue in nameToValueCollection)
                    {
                        if (forNodeType == NodeType.Attribute)
                        {
                            AddOrUpdateAttribute(childElement, idToValue.Key, idToValue.Value);
                        }
                        else if (forNodeType == NodeType.Element)
                        {
                            AddOrUpdateChildElement(childElement, idToValue.Key, idToValue.Value);
                        }
                    }

                    this.customXmlPartCore.WriteElementToCustomXmlPart(customXmlPart, rootElement);
                }
                else
                {
                    childElement = GetElementFromNameToValueCollectionForType(nameToValueCollection, childElementXName, forNodeType);
                    rootElement.Add(childElement);
                }
            }
            else
            {
                customXmlPart = this.customXmlPartCore.AddCustomXmlPart(mainDocumentPart, rootElementName);
                childElement = GetElementFromNameToValueCollectionForType(nameToValueCollection, childElementXName, forNodeType);
                rootElement.Add(childElement);
            }

            this.customXmlPartCore.WriteElementToCustomXmlPart(customXmlPart, rootElement);
        }

        public Dictionary<string, string> GetNameToValueCollectionFromElementForType(MainDocumentPart mainDocumentPart, string elementName, NodeType forNodeType)
        {
            Dictionary<string, string> nameToValueCollection = new Dictionary<string, string>();
            CustomXmlPart customXmlPart = this.customXmlPartCore.GetCustomXmlPart(mainDocumentPart);

            if (customXmlPart != null)
            {
                XElement element = this.customXmlPartCore.GetFirstElementFromCustomXmlPart(customXmlPart, elementName);

                if (element != null)
                {
                    if (forNodeType == NodeType.Element)
                    {
                        foreach (XElement elem in element.Elements())
                        {
                            nameToValueCollection.Add(elem.Name.LocalName, elem.Nodes().Where(node => node.NodeType == XmlNodeType.Element).FirstOrDefault().ToString());
                        }
                    }
                    else if (forNodeType == NodeType.Attribute)
                    {
                        foreach (XAttribute attr in element.Attributes())
                        {
                            nameToValueCollection.Add(attr.Name.LocalName, attr.Value);
                        }
                    }
                }
            }

            return nameToValueCollection;
        }

        private XElement GetElementFromNameToValueCollectionForType(Dictionary<string, string> nameToValueCollection, XName elementXName, NodeType nodeType)
        {
            XElement element = new XElement(elementXName);

            foreach (KeyValuePair<string, string> idToValue in nameToValueCollection)
            {
                if (nodeType == NodeType.Element)
                {
                    AddOrUpdateChildElement(element, idToValue.Key, idToValue.Value);
                }
                else if (nodeType == NodeType.Attribute)
                {
                    AddOrUpdateAttribute(element, idToValue.Key, idToValue.Value);
                }
            }

            return element;
        }

        private void AddOrUpdateAttribute(XElement element, string attributeName, string attributeValue)
        {
            XAttribute attrToUpdate = element.Attributes().Where(attr => attr.Name.LocalName.Equals(attributeName)).FirstOrDefault();

            if (attrToUpdate != null)
            {
                attrToUpdate.Value = attributeValue;
            }
            else
            {
                XAttribute attr = new XAttribute(attributeName, attributeValue);
                element.Add(attr);
            }
        }

        private void AddOrUpdateChildElement(XElement element, string childElementName, string childElementValue)
        {
            XElement childElement = element.Elements().Where(elem => elem.Name.LocalName.Equals(childElementName)).FirstOrDefault();
            XElement newChildElement = new XElement(XName.Get(childElementName, this.customXmlPartCore.namespaceUri));
            newChildElement.Add(XElement.Parse(childElementValue));

            if (childElement != null)
            {
                childElement.ReplaceWith(newChildElement);
            }
            else
            {
                element.Add(newChildElement);
            }
        }

    }
}
