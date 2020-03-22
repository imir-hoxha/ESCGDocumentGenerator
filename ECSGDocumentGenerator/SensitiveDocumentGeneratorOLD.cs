using ConsoleApp1.Domain;
using ECSGDocumentGenerator.Model;
using DocumentFormat.OpenXml.Wordprocessing;
using ECSGDocumentGenerator;
using ECSGDocumentGenerator.Domain;
using System.Collections.Generic;

namespace ConsoleApp1
{
    public class SensitiveDocumentGeneratorOLD : DocumentGenerator
    {


        public SensitiveDocumentGeneratorOLD(DocumentGenerationInfo generationInfo) : base(generationInfo) { }

        protected override Dictionary<string, PlaceHolderType> GetPlaceHolderTagToTypeCollection()
        {
            Dictionary<string, PlaceHolderType> placeHolderTagToTypeCollection = new Dictionary<string, PlaceHolderType>
            {

    
                // Handle container placeholders            
                { DocumentPlaceHolders.PlaceholderContainerA, PlaceHolderType.Container },

                // Handle recursive placeholders       
                //{ DocumentPlaceHolders.PlaceholderRecursiveA, PlaceHolderType.Recursive },
                //{ DocumentPlaceHolders.PlaceholderRecursiveB, PlaceHolderType.Recursive },

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
                { DocumentPlaceHolders.PlaceholderNonRecursiveN, PlaceHolderType.NonRecursive },
                { DocumentPlaceHolders.PlaceholderNonRecursiveO, PlaceHolderType.NonRecursive }
            };

            return placeHolderTagToTypeCollection;
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
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveB:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.leadDg.text;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveC:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.caseTitle;
                    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveD:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.["Infringement Reference"];
                //    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveE:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.reasonForSensitivity;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveF:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.lastAdoptedProposalDecision;
                    break;
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
                    //break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveK:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.P22;
                //    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveL:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.lineToTake;
                    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveM:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C9;
                //    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveN:
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.leadDg.text;
                    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveO:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C28;
=======
                    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Hit).memberState;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveB:
                    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Hit).leadDg.text;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveC:
                    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Hit).caseTitle;
                    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveD:
                //    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                //    content = ((openXmlElementDataContext.DataContext) as Hit).["Infringement Reference"];
                //    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveE:
                    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Hit).reasonForSensitivity;
                    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveF:
                    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Hit).lastAdoptedProposalDecision;
                    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveH:
                //    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                //    content = ((openXmlElementDataContext.DataContext) as Hit).P22;
                //    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveI:
                //    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                //    content = ((openXmlElementDataContext.DataContext) as Hit).P1;
                //    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveJ:
                //    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                //    content = ((openXmlElementDataContext.DataContext) as Hit).P1;
                    //break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveK:
                //    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                //    content = ((openXmlElementDataContext.DataContext) as Hit).P22;
                //    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveL:
                    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Hit).lineToTake;
                    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveM:
                //    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                //    content = ((openXmlElementDataContext.DataContext) as Hit).C9;
                //    break;
                case DocumentPlaceHolders.PlaceholderNonRecursiveN:
                    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as Hit).leadDg.text;
                    break;
                //case DocumentPlaceHolders.PlaceholderNonRecursiveO:
                //    tagValue = ((openXmlElementDataContext.DataContext) as Hit).Id.ToString();
                //    content = ((openXmlElementDataContext.DataContext) as Hit).C28;
>>>>>>> 3890695f2ae98b5ec3af60a4f929077de2d09acb
                //    break;
            }

            if (!string.IsNullOrEmpty(tagValue))
            {
                this.SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
            }

            this.SetContentOfContentControl(openXmlElementDataContext.Element as SdtElement, content);
        }

        /*
            A C1 Member State
            B C24 Lead DG
            C C2 Title
            D C18 Infringement Reference
            E C32 Reason for sensitivity
            F P1 Decision Type
            G P1 Decision Type
            H P22 Decision Sent to the MS
            I P1 Decision Type
            J P22 Decision Sent to the MS
            K C36 Policy Context
            L C76 Line to take
            M C9 DG Case Handler
            N C24 Lead DG
            O C28 Date of Last Update for State of the Fiche
*/

        protected override void RecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
        {
            throw new System.NotImplementedException();
            //if (openXmlElementDataContext == null || openXmlElementDataContext.Element == null || openXmlElementDataContext.DataContext == null)
            //{
            //    return;
            //}

            //string tagGuidPart;
            //string tagPlaceHolderValue;
            //GetTagValue(openXmlElementDataContext.Element as SdtElement, out tagPlaceHolderValue, out tagGuidPart);

            //switch (tagPlaceHolderValue)
            //{
            //    case PlaceholderRecursiveA:

            //        //foreach (Vendor testB in ((openXmlElementDataContext.DataContext) as Order).vendors)
            //        //{
            //        //    SdtElement clonedElement = this.CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext() { Element = openXmlElementDataContext.Element, DataContext = testB });
            //        //}

            //        openXmlElementDataContext.Element.Remove();

            //        break;
            //    case PlaceholderRecursiveB:

            //        //foreach (Item testC in ((openXmlElementDataContext.DataContext) as Order).items)
            //        //{
            //        //    SdtElement clonedElement = this.CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext() { Element = openXmlElementDataContext.Element, DataContext = testC });
            //        //}

            //        openXmlElementDataContext.Element.Remove();
            //        break;
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
<<<<<<< HEAD
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
=======
                    tagValue = (openXmlElementDataContext.DataContext as NonSensitiveReport).Id.ToString();
>>>>>>> 3890695f2ae98b5ec3af60a4f929077de2d09acb

                    if (!string.IsNullOrEmpty(tagValue))
                    {
                        this.SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
                    }

                    foreach (var v in openXmlElementDataContext.Element.Elements())
                    {
                        this.SetContentInPlaceholders(new OpenXmlElementDataContext() { Element = v, DataContext = openXmlElementDataContext.DataContext });
                    }

                    break;
            }
        }

        //protected override void IgnorePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
        //{
        //    throw new System.NotImplementedException();
        //}

    }
}