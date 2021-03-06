﻿using ConsoleApp1.Domain;
using ECSGDocumentGenerator.Model;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace ConsoleApp1
{
    public class NonSensitiveDocumentGeneratorOLD : DocumentGenerator
    {
        #region PlaceHolders
        // Content Control Tags
        protected const string PlaceholderIgnoreA = "PlaceholderIgnoreA";
        protected const string PlaceholderIgnoreB = "PlaceholderIgnoreB";

        protected const string PlaceholderContainerA = "PlaceholderContainerA";

        protected const string PlaceholderRecursiveA = "PlaceholderRecursiveA";
        protected const string PlaceholderRecursiveB = "PlaceholderRecursiveB";

        protected const string PlaceholderNonRecursiveA = "PlaceholderNonRecursiveA";
        protected const string PlaceholderNonRecursiveB = "PlaceholderNonRecursiveB";
        protected const string PlaceholderNonRecursiveC = "PlaceholderNonRecursiveC";
        protected const string PlaceholderNonRecursiveD = "PlaceholderNonRecursiveD";
        protected const string PlaceholderNonRecursiveE = "PlaceholderNonRecursiveE";

        protected const string PlaceholderNonRecursiveF = "PlaceholderNonRecursiveF";
        protected const string PlaceholderNonRecursiveG = "PlaceholderNonRecursiveG";
        protected const string PlaceholderNonRecursiveH = "PlaceholderNonRecursiveH";
        protected const string PlaceholderNonRecursiveI = "PlaceholderNonRecursiveI";
        protected const string PlaceholderNonRecursiveJ = "PlaceholderNonRecursiveJ";

        protected const string PlaceholderNonRecursiveK = "PlaceholderNonRecursiveK";
        protected const string PlaceholderNonRecursiveL = "PlaceholderNonRecursiveL";
        protected const string PlaceholderNonRecursiveM = "PlaceholderNonRecursiveM";
        protected const string PlaceholderNonRecursiveN = "PlaceholderNonRecursiveN";
        #endregion

        public NonSensitiveDocumentGeneratorOLD(DocumentGenerationInfo generationInfo) : base(generationInfo) { }

        //it is overwritten here but it does not get used directly. it is called in the deriving class
        protected override Dictionary<string, PlaceHolderType> GetPlaceHolderTagToTypeCollection()
        {
            Dictionary<string, PlaceHolderType> placeHolderTagToTypeCollection = new Dictionary<string, PlaceHolderType>
            {

                // Handle ignore placeholders
                { PlaceholderIgnoreA, PlaceHolderType.Ignore },
                { PlaceholderIgnoreB, PlaceHolderType.Ignore },

                // Handle container placeholders            
                { PlaceholderContainerA, PlaceHolderType.Container },

                // Handle recursive placeholders       
                { PlaceholderRecursiveA, PlaceHolderType.Recursive },
                { PlaceholderRecursiveB, PlaceHolderType.Recursive },

                // Handle non recursive placeholders
                { PlaceholderNonRecursiveA, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveB, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveC, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveD, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveE, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveF, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveG, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveH, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveI, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveJ, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveK, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveL, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveM, PlaceHolderType.NonRecursive },
                { PlaceholderNonRecursiveN, PlaceHolderType.NonRecursive }
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
                case PlaceholderNonRecursiveA:
<<<<<<< HEAD
                    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                    content = openXmlElementDataContext.DataContext.hit.memberState;
                    break;
                //case PlaceholderNonRecursiveB:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C24;
                //    break;
                //case PlaceholderNonRecursiveC:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C2;
                //    break;
                //case PlaceholderNonRecursiveD:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C18;
                //    break;
                //case PlaceholderNonRecursiveE:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.C14;
                //    break;
                //case PlaceholderNonRecursiveF:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.P1;
                //    break;
                //case PlaceholderNonRecursiveH:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.P22;
                //    break;
                //case PlaceholderNonRecursiveI:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.P1;
                //    break;
                //case PlaceholderNonRecursiveJ:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.P1;
                //    break;
                //case PlaceholderNonRecursiveK:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.P22;
                //    break;
                //case PlaceholderNonRecursiveL:
                //    tagValue = openXmlElementDataContext.DataContext.hit.Id.ToString();
                //    content = openXmlElementDataContext.DataContext.hit.ED1;
                //    break;
                //case PlaceholderNonRecursiveM:
                //    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                //    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C9;
                //    break;
                //case PlaceholderNonRecursiveN:
                //    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                //    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C28;
                //    break;
=======
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C1;
                    break;
                case PlaceholderNonRecursiveB:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C24;
                    break;
                case PlaceholderNonRecursiveC:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C2;
                    break;
                case PlaceholderNonRecursiveD:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C18;
                    break;
                case PlaceholderNonRecursiveE:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C14;
                    break;
                case PlaceholderNonRecursiveF:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P1;
                    break;
                case PlaceholderNonRecursiveH:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P22;
                    break;
                case PlaceholderNonRecursiveI:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P1;
                    break;
                case PlaceholderNonRecursiveJ:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P1;
                    break;
                case PlaceholderNonRecursiveK:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).P22;
                    break;
                case PlaceholderNonRecursiveL:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).ED1;
                    break;
                case PlaceholderNonRecursiveM:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C9;
                    break;
                case PlaceholderNonRecursiveN:
                    tagValue = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).Id.ToString();
                    content = ((openXmlElementDataContext.DataContext) as NonSensitiveReport).C28;
                    break;
>>>>>>> 3890695f2ae98b5ec3af60a4f929077de2d09acb
            }

            if (!string.IsNullOrEmpty(tagValue))
            {
                this.SetTagValue(openXmlElementDataContext.Element as SdtElement, GetFullTagValue(tagPlaceHolderValue, tagValue));
            }

            this.SetContentOfContentControl(openXmlElementDataContext.Element as SdtElement, content);
        }

        protected override void RecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null || openXmlElementDataContext.Element == null || openXmlElementDataContext.DataContext == null)
            {
                return;
            }

            string tagGuidPart;
            string tagPlaceHolderValue;
            GetTagValue(openXmlElementDataContext.Element as SdtElement, out tagPlaceHolderValue, out tagGuidPart);

            switch (tagPlaceHolderValue)
            {
                case PlaceholderRecursiveA:

                    //foreach (Vendor testB in ((openXmlElementDataContext.DataContext) as Order).vendors)
                    //{
                    //    SdtElement clonedElement = this.CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext() { Element = openXmlElementDataContext.Element, DataContext = testB });
                    //}

                    openXmlElementDataContext.Element.Remove();

                    break;
                case PlaceholderRecursiveB:

                    //foreach (Item testC in ((openXmlElementDataContext.DataContext) as Order).items)
                    //{
                    //    SdtElement clonedElement = this.CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext() { Element = openXmlElementDataContext.Element, DataContext = testC });
                    //}

                    openXmlElementDataContext.Element.Remove();
                    break;
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
                case PlaceholderContainerA:
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
        //    //throw new System.NotImplementedException();
        //}

    }
}