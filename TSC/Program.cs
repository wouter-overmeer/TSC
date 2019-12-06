using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TSC
{
    /// <summary>
    /// TODO: 
    /// - Fix row 1 index on consolidation sheet
    /// - No project = no actuals (Davy rule)
    /// - Checkin GIT
    /// </summary>
    class Program
    {
        private static ExcelHandler handler = null;
        public static ExcelHandler Handler { get => handler; set => handler = value; }

        static void Main(string[] args)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            Handler = new ExcelHandler();

            var timeSheetFolder = @"C:\Files\";
            string[] files = Directory.GetFiles(timeSheetFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            Console.WriteLine(string.Format("{0} files found in timesheet folder.", files.Count()));

            var userInput = RetrieveUserInput();

            foreach (string file in files)
            {
                
                Stopwatch stopWatchSheet = new Stopwatch();
                stopWatchSheet.Start();

                if (file.StartsWith("~")) {
                    break;
                }

                var fileName = file.Replace(timeSheetFolder, "");
                Console.WriteLine(string.Format("Processing file: {0}.", fileName));

                Handler.OpenTeamMemberWorkbook(file);
                var timeSheetName = "Input";//Handler.GetVisibleSheetNameByIndex(Handler.XlTeamMemberWorkbook, 1);
                //Console.WriteLine("Workbook: " + timeSheetName);

                switch (userInput)
                {
                    case 1:
                        Console.WriteLine("Let's consolidate some timesheets!");
                        ConsolidateWorkbook("Input");
                        break;
                    case 2:
                        Console.WriteLine("Updating values.");
                        UpdateValues("Values");
                        break;
                    case 3:
                        Console.WriteLine("Converting timesheets.");
                        TransformWorkbook(timeSheetName);
                        break;
                    case 4:
                        Console.WriteLine("Update Sheet links.");
                        UpdateLinks(timeSheetName);
                        break;
                    case 5:
                        Console.WriteLine("Update Values and Consolidate workbooks.");
                        UpdateValuesAndConsolidateWorkbook(timeSheetName);
                        break;
                    default:
                        break;

                }

                if (Handler.CloseWorkbook(Handler.XlTeamMemberWorkbook))
                {
                    Handler.XlTeamMemberWorkbook = null;
                }

                stopWatchSheet.Stop();
                // Get the elapsed time as a TimeSpan value.
                TimeSpan tsSheet = stopWatchSheet.Elapsed;

                Console.WriteLine(string.Format("Execution time: {0}", string.Format("{0}m:{1:D2}s", tsSheet.Minutes, tsSheet.Seconds)));
            }

            Handler.Dispose();

            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            Console.WriteLine(string.Format("Execution time: {0}", string.Format("{0}m:{1:D2}s", ts.Minutes, ts.Seconds)));

            Console.ReadKey();


        }

        public static int RetrieveUserInput()
        {
            int userInput = int.MinValue;

            Console.WriteLine("Welcome, How can I help you today?");

            Console.WriteLine("1 - I want to consolidate timesheets");
            Console.WriteLine("2 - I want to update value sheets");
            Console.WriteLine("3 - I want to convert timesheets");
            Console.WriteLine("4 - I want to update the links in the Excel files");
            Console.WriteLine("5 - I want to update value sheets and consolidate timesheets");

            Console.WriteLine("Please enter your selection: ");

            string input = Console.ReadLine();

            if(int.TryParse(input, out userInput))
            {
                return userInput;
            }

            return userInput;
        }

        private static void UpdateValuesAndConsolidateWorkbook(string timesheetName)
        {
            Handler.OpenReferenceWorkbook(@"C:\Files\Reference\Digital Studio Account Overview.xlsx");

            Handler.CopyValuesByTableName("Values", "lookup_projects");
            Handler.CopyValuesByTableName("Values", "lookup_team");
            //Handler.UpdateLinks(Handler.XlTeamMemberWorkbook);

            //Handler.ConsolidateTimesheet(timesheetName, "Actuals");
            Handler.CopyPasteTimeTable(timesheetName, "Actuals");

            if (Handler.CloseWorkbook(Handler.XlReferenceWorkbook))
            {
                Handler.XlReferenceWorkbook = null;
            }

        }

        private static void ConsolidateWorkbook(string timesheetName)
        {
            Handler.OpenReferenceWorkbook(@"C:\Files\Reference\Digital Studio Account Overview.xlsx");
            //Handler.ConsolidateTimesheet(timesheetName, "Actuals");
            //Handler.UpdateLinks(Handler.XlTeamMemberWorkbook);

            Handler.CopyPasteTimeTable(timesheetName, "Actuals"); 

            if (Handler.CloseWorkbook(Handler.XlReferenceWorkbook))
            {
                Handler.XlReferenceWorkbook = null;
            }

        }

        private static void UpdateLinks(string timesheetName)
        {
           // Handler.OpenReferenceWorkbook(@"C:\Files\Reference\Digital Studio Account Overview.xlsx");
           // Handler.ConsolidateTimesheet(timesheetName, "Actuals");
            Handler.UpdateLinks(Handler.XlTeamMemberWorkbook);
        }

        private static void UpdateValues(string timesheetName)
        {
            Handler.OpenReferenceWorkbook(@"C:\Files\Reference\Digital Studio Account Overview.xlsx");
            Handler.CopyValuesByTableName("Values", "lookup_projects");
            Handler.CopyValuesByTableName("Values", "lookup_team");
            Handler.UpdateLinks(Handler.XlTeamMemberWorkbook);

            if (Handler.CloseWorkbook(Handler.XlReferenceWorkbook))
            {
                Handler.XlReferenceWorkbook = null;
            }
        }

        /// <summary>
        /// One off converstion of old Excel sheet format to new one
        /// </summary>
        /// <param name="originalTimeSheetName"></param>
        private static void TransformWorkbook(string originalTimeSheetName)
        {
            //List<int> indexList = null;
            //if (Handler.XlReferenceWorkbook == null)
            //{
            //    Handler.OpenReferenceWorkbook(@"C:\Files\Reference\Timesheet - reference.xlsx");
            //}

            //indexList = new List<int>();
            //indexList.Add(3);
            //indexList.Add(2);
            //indexList.Add(1);

            //Handler.CopyReferenceSheets(indexList);
            //Handler.ConvertTimeSheet(originalTimeSheetName, "Input");

            //Handler.HideSheetByName(Handler.XlTeamMemberWorkbook, "Values");
            //Handler.HideSheetByName(Handler.XlTeamMemberWorkbook, originalTimeSheetName);
            //Handler.UpdateLinks(Handler.XlTeamMemberWorkbook);
        }
    }
}
