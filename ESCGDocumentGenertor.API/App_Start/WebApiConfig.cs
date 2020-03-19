using ECSGDocumentGenerator;
using ESCGDocumentGenertor.API.Models;
using Microsoft.AspNet.OData.Builder;
using Microsoft.AspNet.OData.Extensions;
using Microsoft.OData;
using Microsoft.OData.Edm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;

namespace ESCGDocumentGenertor.API
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services

            // Web API routes
            //config.MapHttpAttributeRoutes();

            //config.Routes.MapHttpRoute(
            //    name: "DefaultApi",
            //    routeTemplate: "api/{controller}/{id}",
            //    defaults: new { id = RouteParameter.Optional }
            //);
            EnableCorsAttribute cors = new EnableCorsAttribute("*", "*", "GET,POST,PUT,DELETE");
            config.EnableCors(cors);
            config.MapODataServiceRoute("ODataRoute", "Odata", GetEDMModel());
            config.EnsureInitialized();
        }

        private static IEdmModel GetEDMModel()
        {
          //  var builder = new ODataConventionModelBuilder();
          //  builder.Namespace = "Infringement";
          //  builder.ContainerName = "InfringementContainer";

          //  builder.EntitySet<Document>("Documents");
          //  builder.EntityType<Document>().Collection
          //.Function("MergeDocuments")
          //.ReturnsCollectionFromEntitySet<Document>("Documents")
          //.CollectionParameter<string>("ArrayHere");

            var builder = new ODataConventionModelBuilder();

            builder.Namespace = "BriefingService";
            builder.ContainerName = "BriefingServiceContainer";
            builder.EntitySet<Document>("Documents");
            var mergeDocumentsAction = builder.EntityType<Document>().Collection
            .Function("MergeDocuments")
            .Returns<HttpResponseMessage>();
            mergeDocumentsAction.Parameter<string>("country");
            mergeDocumentsAction.Parameter<string>("listId");
            mergeDocumentsAction.Parameter<string>("uniqueId");
            mergeDocumentsAction.CollectionParameter<string>("docs");
            //builder.Namespace = "NS";

            //builder.EntitySet<Document>("Documents");

            //var searchTopics = builder.EntityType<Document>().Collection
            //   .Function("SearchTopics")
            //   .Returns<IHttpActionResult>();
            //searchTopics.Parameter<string>("dg");

            return builder.GetEdmModel();
        }
    }
}
