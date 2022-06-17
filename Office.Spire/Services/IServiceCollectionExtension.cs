using Microsoft.Extensions.DependencyInjection;
using Office.SpireOffice.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace Office.SpireOffice.Services
{
    public static class IServiceCollectionExtension
    {
        public static IServiceCollection AddLibraryServices(this IServiceCollection services)
        {
            services.AddScoped<IDocumentGenerator, DocumentGenerator>();
            return services;
        }
    }
}
