using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Collections.Specialized;
using System.Reflection;
using System.IO;
using System.Xml.Linq;


namespace LaborLens {
    class Program {

      public static string payrollFilename = "paydata2.txt";
      public static string timecardFilename = "timedata.txt";
      public static string project = "Keisha";

      static void Main(string[] args)
      {
         string currentDir = AppDomain.CurrentDomain.BaseDirectory;

         if (!currentDir.ToUpper().Contains(project.ToUpper())) {
            throw new Exception("DB in use is not correct");
         }

         #region timecard importer
         //var connString = @"Data Source=CODICI;User ID=codici;Password=agppci22;Initial Catalog=staging;Encrypt=False";
         //var importer = new TimecardImporter(connString);

         //string projectDir = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "..", ".."));
         ////get all the paths of files with the name time in a directory
         //var paths = Directory.EnumerateFiles(projectDir, "*", SearchOption.AllDirectories)
         //                     .Where(p => Path.GetFileName(p)
         //                         .IndexOf("time", StringComparison.OrdinalIgnoreCase) >= 0)
         //                     .ToList();

         //foreach (var path in paths) {
         //   importer.ImportExcel(path, project);
         //}
         #endregion

         #region one time ingestion of Timedata from PDFs if needed
         //Ingest TNA1
         //var clocks = new Parser().ParseSingleLinePDF(Program.workComputer + Program.timecardFilename);
         //SQL.SQLRepository repo = new SQL.SQLRepository();
         //foreach (var card in clocks) {
         //   repo.AddTimecards(dbName, card.identifier, card.shifDate, card.clockIn, card.ClockOut, card.lp == true ? 1 : 0, card.hrsWorked, card.paycode == null ? String.Empty : card.paycode, card.breakTotal);
         //}
         #endregion

         #region SQL Data Parser - Timecards
         var timecards = new SQL.SQLRepository().GetStagingTimecards(project);
         var empCards = new SQL.SQLRepository().ConvertDataToDict(timecards);

         //  var empCards = new Dictionary<string, List<Timecard>>();
         //  var timecards2 = new SQL.SQLRepository().GetTimecards2(dbName);
         //   empCards = new SQL.SQLRepository().ConvertDataToDict(empCards, timecards2);

         //new ExcelWriter().WriteTimecardsFlat(empCards);
         #endregion
         #region write Data
         //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\CODICI\Desktop\output.txt", true)) {
         //   foreach (KeyValuePair<string, List<Timecard>> entry in empCards)
         //      foreach (Timecard s in entry.Value)
         //         if (s.timepunches.Count == 4 && s.totalHrsActual.TotalHours > 8) {

         //            double diff = s.totalHrsActual.TotalHours - s.regHrsListed;
         //            file.WriteLine(String.Format("{0}^{1}", diff > 0 ? diff : 0, diff < 0 ? Math.Abs(diff) : 0));

         //         }
         //}
         #endregion
         #region SQL Pay Parser
         //var stubdata = new SQL.SQLRepository().GetPaydata(project);
        // var stubs = new SQL.SQLRepository().ConvertPayDataToDict(stubdata);

          Dictionary<string, List<PayStub>> stubs = new Dictionary<string, List<PayStub>>();
         // var stubdata2 = new SQL.SQLRepository().GetPaydata2(dbName);
         // stubs = new SQL.SQLRepository().ConvertPayDataToDict2(stubs, stubdata2);

         #region PayStubs From Text
         //Dictionary<string, List<PayStub>> stubs = new Dictionary<string, List<PayStub>>();
         //var stubs = new PayStubParser().GetStubsFromText(@ConfigurationManager.AppSettings.Get("userDirectory") + "paydata1.txt");
         // stubs = new PayStubParser().GetStubsFromText(stubs, @ConfigurationManager.AppSettings.Get("userDirectory") + Program.payrollFilename);
         #endregion
         #endregion
         #region Rest Breaks Analysis
       //  var analyzer = new RestPeriodAnalyzer();
       //  var results = analyzer.AnalyzeEmployeeRestPeriods(empCards);
       //  var summary = analyzer.GenerateRestPeriodSummary(results);
         #endregion
         #region ADP PDF to CSV Parser
         //ADPParser adpParser = new ADPParser();
         //var adpCards = adpParser.ProcessTimecards(@"C:\Users\CYAN1\OneDrive\Desktop\Law Cases\JGill\Gonzalez v. Eminent\pay\paydata.csv");
         //var adpStubs = adpParser.ProcessPayStubs(@"C:\Users\CYAN1\OneDrive\Desktop\Law Cases\JGill\Gonzalez v. Eminent\pay\paydata.csv");
         #endregion
         #region Roster Analysis
         var dateProcessor = new DataProcessor();
         var rosterResults = dateProcessor.ProcessDateRanges(empCards, stubs);
         #endregion
         #region Create Shift analysis calculations
         var shifts = dateProcessor.GetAnalyzedShifts(empCards);

         Analysis analysis = new Analysis(shifts, stubs);
         analysis.totalEmployeesTimedata = empCards.Count;   
         
         int over35 = Shift.over35;
         int totalShifts = 0;
         double totalLength = 0;

         int c = 0;
         int meal60 = 0;

         foreach (KeyValuePair<string, List<Timecard>> entry in empCards) {
            c += entry.Value.Where(x => x.totalHrsActual.TotalHours == 8).Count();
            meal60 += entry.Value.Where(x => x.mealIs60).Count();

            foreach (Timecard t in entry.Value) {
               totalShifts++;
               totalLength += t.totalHrsActual.TotalHours;
            }
         }
         analysis.avgShiftlength = totalLength / totalShifts;
         #endregion

         #region count split shifts
         int split1hr = 0; //shifts separated by 1 hour or less
         int split2hr = 0; //shifts separated by 2 hour or less

         foreach (KeyValuePair<string, List<Timecard>> entry in empCards) {

            for (int pos = 0; pos < entry.Value.Count - 1; pos++) {
               double timeBtwnShfts = entry.Value[pos + 1].timepunches[0].datetime.Subtract(entry.Value[pos].timepunches[entry.Value[pos].timepunches.Count - 1].datetime).TotalHours;
               if (timeBtwnShfts <= 2 && timeBtwnShfts > 0)
                  split2hr++;
               if (timeBtwnShfts <= 1 && timeBtwnShfts > 0)
                  split1hr++;
            }
         }
         #endregion

         #region PAGA ANALYSIS (timesheets without stubs)
         ///////////Create timesheets from payperiod dates, only if there are timecards available (not stubs needed --> good in case they were not provided)
         Paga p = new Paga(Globals.pagaInitDt, Globals.pagaEndDt); //Create 1st to 15th, then remaining part of the month
        // Paga p = new Paga(Globals.pagaInitDt, Globals.pagaEndDt, 14); //Use this one to calculate at 14 or 7 days apart

         //////These timesheets do not have the paystubs, since some could be missing//////////////////////////
         Dictionary<string, List<Timesheet>> timesheetsWithoutStubs = p.PopulatePeriods(empCards);

      //    new ExcelWriter().WritePagaViolations(timesheetsWithoutStubs, p.periods);

         ////////Create PAGA penalty for the months//////////////////////
         var pagaData = new PayPeriods();
         pagaData.CalculatePagaPeriods(timesheetsWithoutStubs, Globals.pagaInitDt); //Use the PAGA start date, populate using Sheets w/o stubs

         analysis.totalWorkweeks = analysis.PeriodAnalysis(timesheetsWithoutStubs);
         #endregion

         #region Populate Timesheets
         var timeSheets = new PayPeriods().PopulateADPTimesheets(stubs, empCards); //MUST DO FOR PAY PERIODS    
         analysis.PeriodAnalysis(timeSheets); //Get Period analysis
          new PayPeriods().CalculatePagaPeriods(timeSheets, Globals.pagaInitDt); //Use the PAGA start date, populate using Sheets w/o stubs

         #endregion

         analysis.UnpaidOvertime(timeSheets);

         #region Rounding


         // new ExcelWriter().WriteRoundingActualVsListed(timesheetsWithoutStubs);
         #endregion

         #region AutoDeductions
         AutoDeductions deductions = new AutoDeductions();
         var autoDeductHhrs = deductions.AutoDeductHours(timeSheets);
         #endregion

         #region Create Word Analysis Output
         analysis.CompleteCalcuations();
         // new DocWriter().WriteDocument(analysis);
         #endregion

         #region Create Analysis Spreadsheet
         analysis.meal30 = Shift.mealIs30;
         analysis.mealsTaken = Shift.totalMeals;
         analysis.shift8 = Shift.shiftIs8;

         new ExcelWriter().PoulateGraphData(shifts, empCards.Count, pagaData, over35, analysis);
         #endregion

         #region Salary analysis
         // new ExcelWriter().PoulateRoster(rosterResults);
        //  new ExcelWriter().PoulateSummaryPayData(stubs, analysis); //Write pay data by year
        //  new ExcelWriter().WritePayDetails(stubs); //Write pay data by employee

         //double totalHrs = 0;
         //double cnt = 0;

         //using (System.IO.StreamWriter file = new System.IO.StreamWriter(output, true)) {
         //   foreach (KeyValuePair<string, List<PayStub>> entry in stubs) {
         //      // c += entry.Value.Where(x => x.totalHrsActual.TotalHours < 2).Count();
         //      DateTime lastPaycheck = DateTime.MinValue;

         //      foreach (PayStub t in entry.Value) {
         //         if (!t.periodBegin.HasValue)
         //            continue;
         //         //  cnt++;
         //         if (t.periodBegin.Value > lastPaycheck)
         //            lastPaycheck = t.periodEnd.Value;
         //      }

         //      file.WriteLine(String.Format("{0} | {1}", entry.Key, lastPaycheck.ToShortDateString()));
         //   }
         //}
         #endregion


         new ExcelWriter().WriteTimesheetViolations(timeSheets);
      }



      public static int GetTotalWorkweeks(Dictionary<string, List<Timesheet>> empSheets)
      {
         int total = 0;
         foreach (KeyValuePair<string, List<Timesheet>> entry in empSheets) {
            foreach (Timesheet s in entry.Value) {
               total += s.GetWorkWeeks();
            }
         }
         return total;
      }

      public static Dictionary<DateTime, int> GetPayPeriodTotals(Dictionary<string, List<Timesheet>> empSheets)
      {
         Dictionary<DateTime, int> dict = new Dictionary<DateTime, int>();

         foreach (KeyValuePair<string, List<Timesheet>> entry in empSheets) {
            foreach (Timesheet s in entry.Value) {
               if (s.periodBegin != null) {

                  if (!dict.ContainsKey(s.periodBegin.Value))
                     dict[s.periodBegin.Value] = 0;

                  dict[s.periodBegin.Value]++;
               }
            }
         }

         return dict;
      }
   }
}

