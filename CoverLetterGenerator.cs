using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Google.Maps;
using Google.Maps.Geocoding;

namespace myWebApp
{
    public class CoverLetterGenerator
    {
        // private fields
        private string _companyName;
        private string _positionTitle;
        private string _companyNameInput;
        private string _positionTitleInput;
        private DateTime _currentDateTime;
        private string _currentDate;
        private string _templatePath;
        private string _newDocPath;
        private string _formattedCompanyName;
        private string _formattedPositionTitle;
        private string _closestCity;
        private string _closestCityInput;
        private AddressObject _companyAddressObj;
        private string _documentText;
        private Regex _regexCompanyName;
        private Regex _regexCurrentDate;
        private Regex _regexPositionTitle;
        private Regex _regexCompanyStreet;
        private Regex _regexCompanyCityState;
        private Regex _regexCompanyZipCode;

        // class properties
        public string CompanyName
        {
            get => _companyName;
            set => _companyName = value;
        }
        public string PositionTitle
        {
            get => _positionTitle;
            set => _positionTitle = value;
        }
        public string CompanyNameInput
        {
            get => _companyNameInput;
            set => _companyNameInput = value;
        }
        public string PositionTitleInput
        {
            get => _positionTitleInput;
            set => _positionTitleInput = value;
        }
        public DateTime CurrentDateTime
        {
            get => _currentDateTime;
            set => _currentDateTime = value;
        }
        public string CurrentDate
        {
            get => _currentDate;
            set => _currentDate = value;
        }
        public string TemplatePath
        {
            get => _templatePath;
            set => _templatePath = value;
        }
        public string NewDocPath
        {
            get => _newDocPath;
            set => _newDocPath = value;
        }
        public string FormattedCompanyName
        {
            get => _formattedCompanyName;
            set => _formattedCompanyName = value;
        }
        public string FormattedPositionTitle
        {
            get => _formattedPositionTitle;
            set => _formattedPositionTitle = value;
        }
        public string ClosestCity
        {
            get => _closestCity;
            set => _closestCity = value;
        }
        public string ClosestCityInput
        {
            get => _closestCityInput;
            set => _closestCityInput = value;
        }
        public AddressObject CompanyAddressObj
        {
            get => _companyAddressObj;
            set => _companyAddressObj = value;
        }
        public string DocumentText
        {
            get => _documentText;
            set => _documentText = value;
        }
        public Regex RegexCompanyName
        {
            get => _regexCompanyName;
            set => _regexCompanyName = value;
        }
        public Regex RegexCurrentDate
        {
            get => _regexCurrentDate;
            set => _regexCurrentDate = value;
        }
        public Regex RegexPositionTitle
        {
            get => _regexPositionTitle;
            set => _regexPositionTitle = value;
        }
        public Regex RegexCompanyStreet
        {
            get => _regexCompanyStreet;
            set => _regexCompanyStreet = value;
        }
        public Regex RegexCompanyCityState
        {
            get => _regexCompanyCityState;
            set => _regexCompanyCityState = value;
        }
        public Regex RegexCompanyZipCode
        {
            get => _regexCompanyZipCode;
            set => _regexCompanyZipCode = value;
        }

        // constructor
        public CoverLetterGenerator()
        {
            CompanyAddressObj = new AddressObject();
            TemplatePath = "coverLetterTemplates/template.docx";
        }


        public class AddressObject
        {

            // private fields
            private string _streetAddress;
            private string _city;
            private string _state;
            private string _cityState;
            private string _zipCode;

            // class properties
            public string StreetAddress
            {
                get => _streetAddress;
                set => _streetAddress = value;
            }
            public string City
            {
                get => _city;
                set => _city = value;
            }
            public string State
            {
                get => _state;
                set => _state = value;
            }
            public string CityState
            {
                get => _cityState;
                set => _cityState = value;
            }
            public string ZipCode
            {
                get => _zipCode;
                set => _zipCode = value;
            }

        }

        public void GetCompanyName()
        {
            Console.Write("Enter the company name - ");
            CompanyNameInput = Console.ReadLine();
            CompanyName = CompanyNameInput;
            FormattedCompanyName = FormatString(CompanyNameInput);
            return;
        }
        public void GetPositionTitle()
        {
            Console.Write("Enter the position title - ");
            PositionTitleInput = Console.ReadLine();
            PositionTitle = PositionTitleInput;
            FormattedPositionTitle = FormatString(PositionTitleInput);
        }
        public void GetTime()
        {
            CurrentDateTime = DateTime.Today; // As DateTime
            CurrentDate = CurrentDateTime.ToString("MM/dd/yyyy"); // As String
        }
        public string FormatString(string input)
        {
            return input.Replace(" ", "");
        }
        public void GetNewDocPath()
        {
            NewDocPath = "output/" + FormattedCompanyName + "-" + FormattedPositionTitle + "-CoverLetter.docx";
        }

        public void RegexReplace()
        {
            RegexCurrentDate = new Regex("CURRENTDATE");
            RegexCompanyName = new Regex("COMPANYNAME");
            //regex is stupid so i have to do this
            RegexCompanyStreet = new Regex("COMPANYLOCATION");
            RegexCompanyCityState = new Regex("COMPANYCITYSTATE");
            RegexCompanyZipCode = new Regex("COMPANYZIPCODE");
            RegexPositionTitle = new Regex("COMPANYPOSITION");

            // regex replace
            DocumentText = RegexCurrentDate.Replace(DocumentText, CurrentDate);
            DocumentText = RegexCompanyName.Replace(DocumentText, CompanyName);
            DocumentText = RegexCompanyStreet.Replace(DocumentText, CompanyAddressObj.StreetAddress);
            DocumentText = RegexCompanyCityState.Replace(DocumentText, CompanyAddressObj.CityState);
            DocumentText = RegexCompanyZipCode.Replace(DocumentText, CompanyAddressObj.ZipCode);
            DocumentText = RegexPositionTitle.Replace(DocumentText, PositionTitle);
            DocumentText = RegexCompanyName.Replace(DocumentText, CompanyName);
            DocumentText = RegexCompanyCityState.Replace(DocumentText, CompanyAddressObj.CityState);
            DocumentText = RegexCompanyName.Replace(DocumentText, CompanyName);

        }
        public string GetClosestCity()
        {
            Console.Write("(Optional): Enter the nearest city to " + CompanyName + " - ");
            ClosestCityInput = Console.ReadLine();
            ClosestCity = ClosestCityInput;
            string googleMapsSearchQuery = CompanyName + ClosestCity;
            return googleMapsSearchQuery;
        }

        public void GetCompanyAddress()
        {

            Console.WriteLine("Using Google Maps Services...");
            GoogleSigned.AssignAllServices(new GoogleSigned("AIzaSyCtOR1S6lPWMyK1jE9IiTGMX10-s8z3e9Q"));
            var request = new GeocodingRequest();
            request.Address = GetClosestCity();
            var response = new GeocodingService().GetResponse(request);
            //The GeocodingService class submits the request to the API web service, and returns the
            //response strongly typed as a GeocodeResponse object which may contain zero, one or more results.

            //Assuming we received at least one result, let's get some of its properties:
            if (response.Status == ServiceResponseStatus.Ok && response.Results.Count() > 0)
            {
                var result = response.Results.First();
                string[] addressString = result.FormattedAddress.Split(","); // ex 12345 Street St, City, State Zip
                string companyStreet = addressString[0]; // street address
                string companyCity = addressString[1]; // city
                string companyStateZip = addressString[2]; // state and zipcode

                // remove leading spaces
                companyStreet = companyStreet.Trim();
                companyCity = companyCity.Trim();
                companyStateZip = companyStateZip.Replace(" ", "");

                /// separate companyStateZip into companyState and companyZip WITHOUT REGEX
                int index = companyStateZip.IndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
                string companyState = companyStateZip.Substring(0, index);
                int companyZipCode = Int32.Parse(companyStateZip.Substring(index));

                // remove leading space
                companyState = companyState.Trim();

                // add to AddressObject for access
                CompanyAddressObj.StreetAddress = companyStreet;
                CompanyAddressObj.City = companyCity;
                CompanyAddressObj.State = companyState;
                CompanyAddressObj.CityState = companyCity + ", " + companyState;
                CompanyAddressObj.ZipCode = companyZipCode.ToString();

                LogLocation();
            }
            else
            {
                Console.WriteLine("Unable to geocode.  Status={0} and ErrorMessage={1}", response.Status, response.ErrorMessage);
            }
        }
        public void LogLocation()
        {
            Console.WriteLine("Street Address: " + CompanyAddressObj.StreetAddress);
            Console.WriteLine("City: " + CompanyAddressObj.City);
            Console.WriteLine("State: " + CompanyAddressObj.State);
            Console.WriteLine("Zip Code: " + CompanyAddressObj.ZipCode);
            Console.WriteLine();
        }
        public void SearchAndReplace()
        {

            Console.WriteLine("Creating " + CompanyName + " cover letter...");

            // create copy of template document so we don't lose it
            File.Copy(TemplatePath, NewDocPath);

            // open the file and read contents to var
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(NewDocPath, true))
            {
                DocumentText = null;

                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    DocumentText = sr.ReadToEnd();
                }
                RegexReplace();
                // write to new file
                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(DocumentText);
                }

                Console.WriteLine("Done!");
                Console.WriteLine();
            }
        }

    }
}
