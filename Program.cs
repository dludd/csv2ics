using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Globalization;

namespace csv2ics
{
    static class Constants
    {
        // programmnames and version
        public const String programname_short = "CSV2ICS";
        public const String programname_long = "CSV to ICS-File";
        public const String programversion = "0.3";
        public const String programauthor = "Danilo Ludwig";
        // returncodes;
        public const int rc_OK = 0;
        public const int rc_noArguments = 1;
        public const int rc_inputfilenameNotExists = 2;
        public const int rc_fileAccessError = 50;
        
        public static void show_help(int rc)
        {
            // show programname and Version and help for arguments
            System.Console.WriteLine(Constants.programname_short + " v" + Constants.programversion + " - " + Constants.programname_long + " - (c) 2016,2019 by " + Constants.programauthor);
            switch (rc)
            {
                case rc_noArguments:
                    System.Console.WriteLine("please add arguments");
                    break;
                case rc_inputfilenameNotExists:
                    System.Console.WriteLine("input file dosn't exist");
                    break;
                default:
                    System.Console.WriteLine("something wrong with the arguments");
                    break;
            }
            System.Console.WriteLine("arguments are:");
            System.Console.WriteLine("full csv-filename");
         }//show_help()

    }//class Constants

    class Program
    {

        public static String RemoveDiacritics(String s)
        {
            String normalizedString = s.Normalize(NormalizationForm.FormD);
            StringBuilder stringBuilder = new StringBuilder();

            for (int i = 0; i < normalizedString.Length; i++)
            {
                Char c = normalizedString[i];
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    stringBuilder.Append(c);
            }
            return stringBuilder.ToString();
        }

        public static String ReplaceUmlaute(String s)
        {
            s = s.Replace("ä", "ae").Replace("Ä", "Ae");
            s = s.Replace("ö", "oe").Replace("Ö", "Oe");
            s = s.Replace("ü", "ue").Replace("Ü", "Ue");
            s = s.Replace("ß", "ss");
            return s;
        }

        static int Main(string[] args)
        {
            // check, if there are arguments
            if (args.Length == 0)
            {
                Constants.show_help(Constants.rc_noArguments); // show programname and Version and help for arguments
                return Constants.rc_noArguments; // end the program with returncode for missing arguments
            }
            else
            {
                // here is the real start
                // get the arguments and store it as inputfilename
                string inputfilename = args[0];

                // check if inputfilename exist
                if (!File.Exists(inputfilename))
                {
                    Constants.show_help(Constants.rc_inputfilenameNotExists);
                    return Constants.rc_inputfilenameNotExists;
                }
                
                // now try to open the inputfilename
                try
                {
                    String icsSummery;
                    String icsUid;
                    String icsLocation;
                    String icsDescription;
                    String icsStart;
                    String icsEnd;
                    String csvGastID;
                    String csvGastname;
                    String csvRoomsID;
                    String csvRaumnr;
                    String csvRaumname;
                    String csvVon;
                    String csvBis;
                    String csvAnzahlTage;
                    int i;

                    //var inputfilereader = new StreamReader(File.OpenRead(@inputfilename), System.Text.Encoding.Default);
                    var inputfilereader = new StreamReader(@inputfilename, Encoding.GetEncoding(850));
                    
                    //print ics file header to stdout
                    System.Console.WriteLine("BEGIN:VCALENDAR");
                    System.Console.WriteLine("VERSION:2.0");
                    System.Console.WriteLine("PRODID: -//csv2ics//csv2ics//DE");
                    //System.Console.WriteLine("PRODID: JesusHausHHT");
                    //System.Console.WriteLine("PRODID: -//Mozilla.org/NONSGML Mozilla Calendar V1.1//EN");
                    //System.Console.WriteLine("METHOD:PUBLISH");

                    //read the csv lines and create the ics vevents
                    while (!inputfilereader.EndOfStream)
                    {
                        var CSVline = inputfilereader.ReadLine(); //read the csv file line
                        var CSVvalues = CSVline.Split(';');       //split the values by ;

                        //get the csvData
                        csvGastID = CSVvalues[0];
                        //csvGastname = RemoveDiacritics(CSVvalues[1].Trim('"'));
                        csvGastname = RemoveDiacritics(ReplaceUmlaute(CSVvalues[1].Trim('"')));
                        csvRoomsID = CSVvalues[2].Trim('"');
                        csvRaumnr = RemoveDiacritics(ReplaceUmlaute(CSVvalues[3]));
                        csvRaumname = RemoveDiacritics(ReplaceUmlaute(CSVvalues[4].Trim('"')));
                        csvVon = CSVvalues[5].Remove(10);
                        csvBis = CSVvalues[6].Remove(10);
                        csvAnzahlTage = CSVvalues[7];

                        //generate icsData out of csvData
                        icsUid = csvGastID + '-' + csvRoomsID + '-'+ csvVon.Replace('.','-') + '-' + csvBis.Replace('.','-') + '-' + csvAnzahlTage;
                        icsSummery = csvRaumnr + ' ' + csvGastname;
                        icsLocation = "Zimmer " + csvRaumnr;
                        icsDescription = icsLocation + " (" + csvRaumname + ") gebucht von " + csvGastname +" " + "vom " + csvVon + " bis " +csvBis;
                        icsStart = csvVon.Substring(6,4) + csvVon.Substring(3,2) + csvVon.Substring(0,2);
                        Int32.TryParse(csvBis.Substring(0, 2), out i);
                        i++;
                        //icsEnd = csvBis.Substring(6,4) + csvBis.Substring(3,2) + csvBis.Substring(0,2);
                        icsEnd = csvBis.Substring(6, 4) + csvBis.Substring(3, 2) + i.ToString("00");

                        //print ics vevent header
                        System.Console.WriteLine("BEGIN:VEVENT");
                        System.Console.WriteLine("UID:"+icsUid);

                        //print ics vevent data
                        //System.Console.WriteLine("DTSTAMP; VALUE = DATE:20160808
                        System.Console.WriteLine("DTSTAMP:" + DateTime.Now.ToString("yyyyMMdd") + "T" + DateTime.Now.ToString("HHmmss") + "Z");
                        //System.Console.WriteLine("DTSTAMP:" + DateTime.Now.ToString("yyyyMMddHHmmss"));
                        System.Console.WriteLine("SUMMARY:" + icsSummery);
                        System.Console.WriteLine("CATEGORIES:Zimmerbuchung");
                        System.Console.WriteLine("DTSTART;VALUE=DATE:" + icsStart);
                        //System.Console.WriteLine("DTSTART:" + icsStart + "T120000");
                        System.Console.WriteLine("DTEND;VALUE=DATE:" + icsEnd);
                        //System.Console.WriteLine("DTEND:" + icsEnd + "T120000");
                        System.Console.WriteLine("LOCATION:" + icsLocation);
                        System.Console.WriteLine("DESCRIPTION:" + icsDescription );
                        //print ics vevent footer
                        System.Console.WriteLine("END:VEVENT");
                    }
                    //print ics file footer to stdout
                    System.Console.WriteLine("END:VCALENDAR");
                    
                } // try read csvfile
                catch (Exception e)
                {   // if somethings went wrong show the error message
                    System.Console.WriteLine(e.Message);
                    return Constants.rc_fileAccessError;
                }
            }//else
            return Constants.rc_OK;             // end programm with returncode for OK
        }//main
    }//class Program
}
