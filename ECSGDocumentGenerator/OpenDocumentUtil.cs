using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.VariantTypes;
using System.Collections.Generic;
using System.Web;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Text;

namespace ConsoleApp1
{
    internal enum PropertyTypes : int
    {
        YesNo,
        Text,
        DateTime,
        NumberInteger,
        NumberDouble
    }

    internal class OpenDocumentUtil
    {
        internal static string SetCustomProperty(WordprocessingDocument document, string propertyName, object propertyValue, PropertyTypes propertyType)
        {
            string returnValue = null;

            var newProp = new CustomDocumentProperty();
            bool propSet = false;

            // Calculate the correct type.
            switch (propertyType)
            {
                case PropertyTypes.DateTime:

                    // Be sure you were passed a real date, 
                    // and if so, format in the correct way. 
                    // The date/time value passed in should 
                    // represent a UTC date/time.
                    if ((propertyValue) is DateTime)
                    {
                        newProp.VTFileTime =
                            new VTFileTime(string.Format("{0:s}Z",
                                Convert.ToDateTime(propertyValue)));
                        propSet = true;
                    }

                    break;

                case PropertyTypes.NumberInteger:
                    if ((propertyValue) is int)
                    {
                        newProp.VTInt32 = new VTInt32(propertyValue.ToString());
                        propSet = true;
                    }

                    break;

                case PropertyTypes.NumberDouble:
                    if (propertyValue is double)
                    {
                        newProp.VTFloat = new VTFloat(propertyValue.ToString());
                        propSet = true;
                    }

                    break;

                case PropertyTypes.Text:
                    newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                    propSet = true;

                    break;

                case PropertyTypes.YesNo:
                    if (propertyValue is bool)
                    {
                        // Must be lowercase.
                        newProp.VTBool = new VTBool(
                          Convert.ToBoolean(propertyValue).ToString().ToLower());
                        propSet = true;
                    }
                    break;
            }

            if (!propSet)
            {
                // If the code was not able to convert the 
                // property to a valid value, throw an exception.
                throw new InvalidDataException("propertyValue");
            }

            // Now that you have handled the parameters, start
            // working on the document.
            newProp.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProp.Name = propertyName;

            var customProps = document.CustomFilePropertiesPart;
            if (customProps == null)
            {
                // No custom properties? Add the part, and the
                // collection of properties now.
                customProps = document.AddCustomFilePropertiesPart();
                customProps.Properties =
                    new Properties();
            }

            var props = customProps.Properties;
            if (props != null)
            {
                // This will trigger an exception if the property's Name 
                // property is null, but if that happens, the property is damaged, 
                // and probably should raise an exception.
                var prop = props.Where(p => ((CustomDocumentProperty)p).Name.Value == propertyName).FirstOrDefault();

                // Does the property exist? If so, get the return value, 
                // and then delete the property.
                if (prop != null)
                {
                    returnValue = prop.InnerText;
                    prop.Remove();
                }

                // Append the new property, and 
                // fix up all the property ID values. 
                // The PropertyId value must start at 2.
                props.AppendChild(newProp);
                int pid = 2;
                foreach (CustomDocumentProperty item in props)
                {
                    item.PropertyId = pid++;
                }
                props.Save();
            }
            return returnValue;
        }

        internal static string GetCustomPropertyValue(WordprocessingDocument document, string propertyName)
        {
            var customProps = document.CustomFilePropertiesPart;
            if (customProps != null)
            {
                var props = customProps.Properties;
                if (props != null)
                {
                    var prop =
                    props.Where(
                    p => ((CustomDocumentProperty)p).Name.Value
                        == propertyName).FirstOrDefault();

                    if (prop != null)
                    {
                        return prop.InnerText;
                    }
                }
            }

            return null;
        }

        internal static void SetTextValue(OpenXmlElement rootElement, string placeholder, string value, string color = null)
        {
            var textElements = rootElement.Descendants<Text>();
            foreach (var text in textElements)
            {
                if (text.InnerText.Contains(placeholder))
                {
                    if (color != null)
                    {
                        var parentRun = FindParentElement<Run>(text);

                        if (parentRun != null)
                        {
                            parentRun.RunProperties.Color = new Color() { Val = color };
                        }
                    }

                    string newValue = text.InnerText.Replace(placeholder, value);
                    text.Parent.ReplaceChild<Text>(new Text(newValue), text);
                }
            }
        }

        internal static void ImportHtmlChunk(WordprocessingDocument doc, OpenXmlElement rootElement, string placeholder, string html)
        {
            var textElement = rootElement.Descendants<Text>().Where(t => t.InnerText.Contains(placeholder)).FirstOrDefault();

            if (textElement != null)
            {
                string chunkId = "myId-" + Guid.NewGuid().ToString();
                var ms = new MemoryStream(new UTF8Encoding(true).GetPreamble().Concat(UTF8Encoding.UTF8.GetBytes(html)).ToArray());
                var formatImportPart = doc.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, chunkId);
                formatImportPart.FeedData(ms);
                var altChunk = new AltChunk();
                altChunk.Id = chunkId;

                var parentTableCell = FindParentElement<TableCell>(textElement);
                var parentParagraph = FindParentElement<Paragraph>(textElement);
                parentTableCell.InsertAfter(altChunk, parentParagraph);

                textElement.Parent.ReplaceChild<Text>(new Text(string.Empty), textElement);
            }
        }

        internal static T FindParentElement<T>(OpenXmlElement element) where T : class
        {
            if (element is T)
            {
                return element as T;
            }
            else if (element.Parent != null)
            {
                return FindParentElement<T>(element.Parent);
            }
            else
            {
                return null;
            }
        }

        internal static T FindTemplateElement<T>(OpenXmlElement rootElement, string placeholder) where T : class
        {
            var textElement = rootElement.Descendants<Text>().Where(el => el.InnerText.Contains(placeholder)).FirstOrDefault();
            if (textElement != null)
            {
                return FindParentElement<T>(textElement);
            }
            else
            {
                return null;
            }
        }
    }
}