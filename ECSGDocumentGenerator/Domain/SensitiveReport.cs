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
        public int size { get; set; }
        public int number { get; set; }
        public Sort sort { get; set; }
        public bool first { get; set; }
        public int numberOfElements { get; set; }
    }

}
