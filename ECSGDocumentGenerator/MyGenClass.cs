using ECSGDocumentGenerator.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ECSGDocumentGenerator
{
    public class MyGenClass
    {
        //from DocumentGeneratorNS
        private static readonly OpenXmlHelper openXmlHelper = new OpenXmlHelper(DocumentGenerationInfo.NamespaceUri);
        //public static Dictionary<string, PlaceHolderType> PlaceHolderTagToTypeCollection { get; set; }
        //#region PlaceHolders
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
    }
}
