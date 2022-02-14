using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Text;


[assembly:FunctionsStartup(typeof(exceltocsv.StartUp))]
namespace exceltocsv
{
    public class StartUp:FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
           
        }
    }
}
