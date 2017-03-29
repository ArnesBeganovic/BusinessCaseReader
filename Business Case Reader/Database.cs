using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Business_Case_Reader
{
    //Fileds for Translation Table. Logic that is used to analyse provided file and understands if provided sheet is formated correctly
    public class TranlationTable
    {
        //Properties explanation

        /*
         * Id is used to numerically represents what is in description. Financial Business Case sheet contains header data and details data together
         * Each data part is marked with unique ID so that Salesforce team can create logic for data import. For example, if "Building and Land" information
         * is marked with id 16 that means Salesforce team can build import logic based on their data storage logic. "Building and Land" data will always have id 16
         * in Adient exported file. 
         * 
         * Description is used as user friendly id explanation. It would be best to store real index data in some external SQL table and generate objects for this
         * class from there, but at this stage, Adient is hard coding this logic within code itself.
         */
        public int id { get; set; }
        public string description { get; set; }


        /*
         * fieldForValue contains Excel Cell reference which has value that needs to be exported in connection to id.
         * exportValue is placeholder for value that needs to be exported into Salesforce and used for reporting
         */
        public string exportValue { get; set; }
        public string fieldForValue { get; set; }

        /*
         * fieldForCheck is contains Excel Cell reference which has value that needs to be checked.
         * textToCheck is hard coded string that has to be in Excel cell in order to accept value for export.
         * If cell text does not match text from textToCheck in field refered by fieldForCheck than it means
         * templates does not match. User needs to have exactly the same template.
         */
        public string textToCheck { get; set; } //what to search in order to be sure you are taking correct info
        public string fieldForCheck { get; set; } //where is string from previous field placed

        /*
         * year is important for detailed data and it is empty for header data. Buildings and Land line can contain value for multiple years and it will use
         * same id. Year will distinct two objects with same export Id. This can be recognizes by Salesforce team while importing.
         */
        public string year { get; set; } //Not all fields need this

        /*
         * Booleans values that checks if code execution was correct.
         * isOK - is by default false. Whenever property value textToCheck matches with sheet data this will become true. This means that if this property is
         * still false, algoritm has not found string provided by textToCheck in fieldForCheck which means that template does not match. 
         * 
         * isForTBL - Data export is used for two Salesforce applications simultaneously. One of them is using all provided data but the other one is using some
         * of the lines. If this filed is true, then this line is used by TBL only.
         * 
         * isHeader - Header and details data are normaly stored in two separate tables. This field is making distinction between them. Header data are 
         * marked with 1, detail data is marked with 0
         * 
         * ExportPart - export file needs to be splited into three parts:
         *                      1) user credentials - for storing and security purposes. We need to know who sends file to the Salesforce. This part contains 
         *                      hash token which will be matched with header data. User credentials have 1 as ExportPart value.
         *                      2) file header data - some basic informations about file that is exported. Date, time, file name, user id. They generate hash
         *                      token for user credentials as well. Apart of that, this part of exported file contains data as Business Unit, Commodity, , Region
         *                      etc. FIle header data have 2 as ExportPart value.
         *                      3) objects from this class or data from excel file itself. THey have 3 as ExportPart value.
         *                      
         */
        public bool isOK { get; set; }
        public bool isForTBL { get; set; }
        public bool isHeader { get; set; }
        public int ExportPart { get; set; }
    }


    //Fileds for Export Table. Not all fileds from Translation table needs to be exported to the Salesforce
    public class ExportTable
    {
        /*
         * Objects from this class will be exported to csv file. Properties meaning is same as in Translation table.
         */
        public int id { get; set; }
        public string description { get; set; }
        public string exportValue { get; set; }
        public string year { get; set; }
        public bool isOK { get; set; }
        public bool isForTBL { get; set; }
        public bool isHeader { get; set; }
        public int ExportPart { get; set; }
    }
}
