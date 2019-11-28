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
        public Content[] content { get; set; }
        public string pageable { get; set; }
        public int totalElements { get; set; }
        public bool last { get; set; }
        public int totalPages { get; set; }
        public int size { get; set; }
        public int number { get; set; }
        public Sort sort { get; set; }
        public bool first { get; set; }
        public int numberOfElements { get; set; }
    }

}
