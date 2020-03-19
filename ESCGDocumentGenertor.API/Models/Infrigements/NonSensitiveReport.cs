using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1.Domain
{
    public class NonSensitiveReport
    {
        public Guid Id = Guid.Empty;
        /// <summary>
        /// C1 Member State
        /// </summary>
        public string C1 { get; set; }

             /// <summary>
        /// C2 Title
        /// </summary>
        public string C2 { get; set; }

        /// <summary>
        /// C14 Incriminated Fact
        /// </summary>
        public string C14 { get; set; }

        /// <summary>
        /// C18 Infrigement Reference
        /// </summary>
        public string C18 { get; set; }

        /// <summary>
        /// C24 Lead DG
        /// </summary>
        public string C24 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string P1 { get; set; }

        public string P22 { get; set; }

        public string ED1 { get; set; }

        public string C9 { get; set; }

        public string C28 { get; set; }

    }
}
