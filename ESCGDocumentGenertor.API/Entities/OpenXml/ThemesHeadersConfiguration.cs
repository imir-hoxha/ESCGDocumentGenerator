using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ECSGDocumentGenerator.Model
{
    public static class ThemesHeadersConfiguration
    {
        public static string ThemesServiceUrl { get => ConfigurationManager.AppSettings["themesInfrigementServiceUrl"]; }
        public static string ThemesAuthenticationToken { get => ConfigurationManager.AppSettings["themesAuthenticationToken"];}
        public static string ThemesApplicationHeader { get => ConfigurationManager.AppSettings["themesApplicationHeader"]; }
        public static string ThemesHost { get => ConfigurationManager.AppSettings["themesHost"]; }
    }
}
