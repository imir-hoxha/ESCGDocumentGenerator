using ECSGDocumentGenerator.Model;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp1
{
    internal class RefreshableDocumentGenerator : NonSensitiveDocumentGeneratorOLD
    {
        //private DocumentGenerationInfo generationInfo;

        public RefreshableDocumentGenerator(DocumentGenerationInfo generationInfo) : base(generationInfo)
        {

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
                case PlaceholderContainerA:
                    SdtElement parentContainer = openXmlElementDataContext.Element as SdtElement;
                    // Sets the parentContainer from CustomXmlPart if refresh else saves the parentContainer markup to CustomXmlPart 
                    this.GetParentContainer(ref parentContainer, tagPlaceHolderValue);
                    base.ContainerPlaceholderFound(placeholderTag, new OpenXmlElementDataContext() { Element = parentContainer, DataContext = openXmlElementDataContext.DataContext });
                    break;
            }
        }
    }
}