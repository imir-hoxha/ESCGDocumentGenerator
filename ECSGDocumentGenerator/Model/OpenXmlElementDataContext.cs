using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ECSGDocumentGenerator.Model
{
    public class OpenXmlElementDataContext
    {
        public OpenXmlElement Element { get; set; }

        public object DataContext { get; set; }
    }
}
