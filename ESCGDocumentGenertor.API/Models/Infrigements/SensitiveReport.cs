using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ECSGDocumentGenerator
{

    public class SensitiveReport
    {
        public string Id { get; set; }
        public Content[] Content { get; set; }
        public string Pageable { get; set; }
        public int TotalElements { get; set; }
        public bool Last { get; set; }
        public int TotalPages { get; set; }
        public int Size { get; set; }
        public int Number { get; set; }
        public Sort Sort { get; set; }
        public bool First { get; set; }
        public int NumberOfElements { get; set; }
    }

}
