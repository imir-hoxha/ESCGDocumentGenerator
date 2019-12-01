using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ECSGDocumentGenerator.Model
{
    public class DocumentGenerationInfo
    {
        public const string NamespaceUri = "http://schemas.WordDocumentGenerator.com/DocumentGeneration";

        public Dictionary<string, PlaceHolderType> PlaceHolderTagToTypeCollection { get; set; }

        public DocumentMetadata Metadata { get; set; }

        public byte[] TemplateData { get; set; }

        public Content DataContext { get; set; }

        public Content[] Contents { get; set; }

        public bool IsDataBoundControls { get; set; }
    }
}
