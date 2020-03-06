using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Google.Maps;
using Google.Maps.Geocoding;

namespace myWebApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //CreateHostBuilder(args).Build().Run();
            Console.WriteLine("COVER LETTER GENERATOR");
            CoverLetterGenerator clg = new CoverLetterGenerator();
            clg.GetTime();
            clg.GetCompanyName();
            clg.GetPositionTitle();
            clg.GetNewDocPath();
            Console.WriteLine(clg.TemplatePath);
            Console.WriteLine(clg.NewDocPath);
            clg.GetCompanyAddress();
            clg.SearchAndReplace();
        }

        //public static IHostBuilder CreateHostBuilder(string[] args) =>
        //    Host.CreateDefaultBuilder(args)
        //        .ConfigureWebHostDefaults(webBuilder =>
        //        {
        //            webBuilder.UseStartup<Startup>();
        //        });
    }
}
