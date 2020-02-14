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

        //    static void Main(string[] args)
        //    {
        //        Console.WriteLine("COVER LETTER GENERATOR");

        //        string docPath = "/Users/richardestes/Desktop/template.docx";
        //        string companyNameInput;
        //        string positionTitleInput;
        //        AddressObject companyAddressObj = new AddressObject();

        //        // setup date variables
        //        DateTime today = DateTime.Today; // As DateTime
        //        string todayAsString = today.ToString("MM/dd/yyyy"); // As String

        //        // get company name
        //        Console.Write("Enter the company name - ");
        //        companyNameInput = Console.ReadLine();
        //        string formattedCompanyName = companyNameInput.Replace(" ", "");

        //        // get position title
        //        Console.Write("Enter the position title - ");
        //        positionTitleInput = Console.ReadLine();
        //        string formattedPositionTitle = positionTitleInput.Replace(" ", "");

        //        // new file name
        //        // ex - Starbucks-SoftwareEngineer-CoverLetter.docx
        //        string newDocPath = "/Users/richardestes/Desktop/" + formattedCompanyName + "-" + formattedPositionTitle + "-CoverLetter.docx";
        //        GetCompanyAddress(companyAddressObj, companyNameInput);
        //        SearchAndReplace(docPath, newDocPath, companyNameInput, positionTitleInput, todayAsString, companyAddressObj);
        //        Console.WriteLine("Creating " + companyNameInput + " cover letter...");
        //    }
        //    

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureWebHostDefaults(webBuilder =>
                {
                    webBuilder.UseStartup<Startup>();
                });
    }
}
