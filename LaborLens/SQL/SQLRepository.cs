
using System;
using System.Collections.Generic;
using System.Data;


namespace LaborLens.SQL {
   class SQLRepository {

      int duplicates = 0;
      /// <summary>
      /// PeyKar
      /// </summary>
      /// <returns></returns>
      public DataTable GetTimecards(string dbName)
      {
         var storedProcedure = new StoredProcedure();
         storedProcedure.StoredProcedureName = dbName + ".dbo" + ".[Timecards.GetShifts]";
         DataTable result = storedProcedure.ExecuteDataSet();
         return result;
      }

      public DataTable GetTimecards2(string dbName)
      {
         var storedProcedure = new StoredProcedure();
         storedProcedure.StoredProcedureName = dbName + ".dbo" + ".[Timecards.GetShifts2]";
         DataTable result = storedProcedure.ExecuteDataSet();
         return result;
      }

      public DataTable GetTimecardsB(string dbName)
      {
         var storedProcedure = new StoredProcedure();
         storedProcedure.StoredProcedureName = dbName + ".dbo" + ".[Timecards.GetShiftsB]";
         DataTable result = storedProcedure.ExecuteDataSet();
         return result;
      }

      public DataTable GetPaydata(string dbName)
      {
         var storedProcedure = new StoredProcedure();
         storedProcedure.StoredProcedureName = dbName + ".dbo" + ".[Timecards.GetPay]";
         DataTable result = storedProcedure.ExecuteDataSet();
         return result;
      }

      public DataTable GetPaydata2(string dbName)
      {
         var storedProcedure = new StoredProcedure();
         storedProcedure.StoredProcedureName = dbName + ".dbo" + ".[Timecards.GetPay2]";
         DataTable result = storedProcedure.ExecuteDataSet();
         return result;
      }

      public DataTable GetSaraviaTimecards()
      {
         var storedProcedure = new StoredProcedure();
         storedProcedure.StoredProcedureName = "saravia.dbo" + ".[Timecards.GetShifts]";
         DataTable result = storedProcedure.ExecuteDataSet();
         return result;
      }

      public double ParseHours(string hrsString)
      {
         if (string.IsNullOrWhiteSpace(hrsString))
            return 0;

         var parts = hrsString.Split(':');
         if (parts.Length != 2)
            throw new FormatException($"Invalid hrs format: {hrsString}");

         int hours = int.Parse(parts[0]);
         int minutes = int.Parse(parts[1]);

         return hours + (minutes / 60.0);
      }

      public Dictionary<string, List<PayStub>> ConvertPayDataToDict(DataTable dt)
      {
         Dictionary<string, List<PayStub>> stubs = new Dictionary<string, List<PayStub>>();
         string identifier = string.Empty;

         foreach (DataRow row in dt.Rows) {

            try {

               identifier = row["EE_ID"].ToString().ToUpper();
               DateTime checkDate = DateTime.Parse(row["Pay_Date"].ToString());
               DateTime end =  checkDate.AddDays(-5); //  ////
               DateTime start = checkDate.AddDays(-6); //  ////

               /////////////setup bi-monthly times////////////////////
               //if (checkDate.Day <= 16) {
               //   end = new DateTime(checkDate.Year != 1 ? checkDate.Year : checkDate.Year - 1, checkDate.Month == 1 ? 12 : checkDate.Month - 1, DateTime.DaysInMonth(checkDate.Year, checkDate.Month == 1 ? 12 : checkDate.Month - 1));
               //   start = new DateTime(checkDate.Year != 1 ? checkDate.Year : checkDate.Year - 1, checkDate.Month == 1 ? 12 : checkDate.Month - 1, 16);
               //}
               //else {
               //   start = new DateTime(checkDate.Year, checkDate.Month, 1);
               //   end = new DateTime(checkDate.Year, checkDate.Month, 15);
               //}

               #region pay data on a single line
               double regRate = 0;// !String.IsNullOrEmpty(row["Rate"].ToString().Replace("$","")) ? Double.Parse(row["Rate"].ToString().Replace("$", "")) : 0;

               double regPay = row["Regular_Wages"].ToString() != string.Empty ? Double.Parse(row["Regular_Wages"].ToString().Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "").Replace("$", "")) : 0;
               double regHrs = row["Reg_hrs"].ToString() != string.Empty ? Double.Parse(row["Reg_hrs"].ToString()) : 0;

               double otHrs = !String.IsNullOrEmpty(row["Overtime_Hours_Total"].ToString()) ? Double.Parse(row["Overtime_Hours_Total"].ToString()) : 0;
               double otPay = !String.IsNullOrEmpty(row["OT_Wages"].ToString().Replace("$", "")) ? Double.Parse(row["OT_Wages"].ToString().Replace("$", "")) : 0;

               double dblOTHrs = !String.IsNullOrEmpty(row["DoubleTime_Hours"].ToString()) ? Double.Parse(row["DoubleTime_Hours"].ToString()) : 0;
               double dblOTPay = !String.IsNullOrEmpty(row["DoubleTime_Earnings"].ToString().Replace("$", "")) ? Double.Parse(row["DoubleTime_Earnings"].ToString().Replace("$", "")) : 0;

             //  double pnltyHrs = !String.IsNullOrEmpty(row["Premium_Hours"].ToString()) ? Double.Parse(row["Premium_Hours"].ToString()) : 0;
             //  double pnltyTPay = !String.IsNullOrEmpty(row["Premium_Earnings"].ToString()) ? Double.Parse(row["Premium_Earnings"].ToString()) : 0;

               double bonus = !String.IsNullOrEmpty(row["B_Bonus_earnings"].ToString()) ? Double.Parse(row["B_Bonus_earnings"].ToString()) : 0;
               //if (bonus == 0)
               //   bonus = !String.IsNullOrEmpty(row["B_Bonus_earnings2"].ToString()) ? Double.Parse(row["B_Bonus_earnings2"].ToString()) : 0;
               //double commissions = !String.IsNullOrEmpty(row["Commission_Total_Amount"].ToString()) ? Double.Parse(row["Commission_Total_Amount"].ToString()) : 0;
               #endregion

               #region pay data on multiple lines
               //double regRate = 0;
               //double regHrs = 0;
               //double regPay = 0;
               //double otHrs = 0;
               //double otPay = 0;
               //double dblOTHrs = 0;
               //double dblOTPay = 0;
               //double pnltyHrs = 0;
               //double pnltyTPay = 0;
               //double bonus = 0;

               //if (row["code"].ToString().ToUpper() == "HOURLY" || row["code"].ToString().ToUpper().Contains("REGULAR")) {
               //   regPay = Double.Parse(row["pay"].ToString().Trim(' ').Trim('$').Trim('(').Trim(')').Trim('$').Replace(",", ""));
               //   if (regPay < 0)
               //      continue;
               //   if (row["hrs"].ToString() == String.Empty)
               //      continue;

               //   if (row["hrs"].ToString().Contains(':'))
               //      regHrs = ParseHours(row["hrs"].ToString());
               //   else
               //      regHrs = !String.IsNullOrEmpty(row["hrs"].ToString()) ? Double.Parse(row["hrs"].ToString()) : 0;
               //}

               //if (row["code"].ToString().ToUpper() == ("OVERTIME")) {

               //   if(row["hrs"].ToString().Contains(':'))
               //     otHrs = ParseHours(row["hrs"].ToString()); 
               //   else
               //      otHrs = !String.IsNullOrEmpty(row["hrs"].ToString()) ? Double.Parse(row["hrs"].ToString()) : 0;

               //      otPay = !String.IsNullOrEmpty(row["pay"].ToString()) ? Double.Parse(row["pay"].ToString().Trim('$').Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "")) : 0;
               //}

               //if (row["code"].ToString().ToUpper() == "DT" | row["code"].ToString().ToUpper().Contains("DOUBLE")) {
               //   dblOTHrs = !String.IsNullOrEmpty(row["hrs"].ToString()) ? Double.Parse(row["hrs"].ToString()) : 0;
               //   dblOTPay = !String.IsNullOrEmpty(row["pay"].ToString()) ? Double.Parse(row["pay"].ToString().Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "")) : 0;
               //}

               //if (row["code"].ToString().ToUpper().Contains("MEAL")) {
               //   pnltyHrs = !String.IsNullOrEmpty(row["hrs"].ToString()) ? Double.Parse(row["hrs"].ToString()) : 0;
               //   pnltyTPay = !String.IsNullOrEmpty(row["pay"].ToString()) ? Double.Parse(row["pay"].ToString().Trim('$').Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "")) : 0;
               //}

               //if (row["code"].ToString().ToUpper().Contains("BONUS")) {
               //   //  pnltyHrs = !String.IsNullOrEmpty(row["hrs"].ToString()) ? Double.Parse(row["hrs"].ToString()) : 0;
               //   bonus = !String.IsNullOrEmpty(row["pay"].ToString()) ? Double.Parse(row["pay"].ToString().Trim('$').Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "")) : 0;
               //}
               #endregion

               if (!stubs.ContainsKey(identifier)) {
                  stubs[identifier] = new List<PayStub>();
               }
               if (regHrs > 0) {
                  if (regPay > 0)
                     regRate = regPay / regHrs; //Double.Parse(row["Rate"].ToString().ToUpper());
                  else if (otPay > 0 && otHrs > 0)
                     regRate = (otPay / otHrs) / 1.5;
               }

               #region unused for calculating mid and end
               //  DateTime endOfMonth = new DateTime(end.Year, end.Month, DateTime.DaysInMonth(end.Year, end.Month));
               //  DateTime midMonth = new DateTime(end.Year, end.Month, 15);
               //
               //determine if one week back from paycheck is closer to the end of the month or mid-month
               //if (Math.Abs((end - endOfMonth).TotalDays) < Math.Abs((end - midMonth).TotalDays) || end.Day == 1) {
               //   if (end.Day == 1) {
               //      endOfMonth = endOfMonth.AddDays(-1);
               //   }

               //   end = endOfMonth;
               //   begin = new DateTime(end.Year, end.Month, 16);
               //}
               //else {

               //   end = midMonth;
               //   begin = new DateTime(end.Year, end.Month, 1);
               //}
               #endregion

               PayStub stub = new PayStub() {
                  identifier = identifier,

                  periodBegin = start,
                  periodEnd = end,
                  //checkDate = checkDate,
                  regHrs = regHrs,
                  regPay = regPay,
                  regRate = regRate, //regHrs > 0 ? regPay / regHrs : 0,//regRate,//
                  bonus = bonus,
                  //commissions = commissions,
                  otRate = otHrs != 0 ? otPay / otHrs : 0,
                  otHrs = otHrs,
                  otPay = otPay,
                  doubltOtRate = dblOTHrs != 0 ? dblOTPay / dblOTHrs : 0,
                  doubleOtHrs = dblOTHrs,
                  doubleOtPay = dblOTPay,
                //  penaltyHrs = pnltyHrs,
               //   penaltyPay = pnltyTPay
               };

               if (stubs[identifier].Contains(stub)) {
                  int pos = stubs[identifier].IndexOf(stub);
                  stubs[identifier][pos].Merge(stub);
               } else
                  stubs[identifier].Add(stub);

            } catch (Exception e) {
               throw new Exception("Bad Wage Data Present");
            }
         }

         foreach (KeyValuePair<string, List<PayStub>> entry in stubs) {
            for (int pos = 0; pos < entry.Value.Count; pos++) {
               var stub = entry.Value[pos];
               if (stub.regHrs == 0 && stub.bonus == 0 && stub.commissions == 0) {

                  if (stub.otHrs > 0 && stub.regPay > 0) {
                     stub.regRate = stub.otRate / 1.5;
                     stub.regHrs = stub.regPay / stub.regRate;
                  } else {
                     entry.Value.RemoveAt(pos);
                     pos--;
                  }
               }
            }
         }

         return stubs;
      }

      public Dictionary<string, List<PayStub>> ConvertPayDataToDict2(Dictionary<string, List<PayStub>> stubs, DataTable dt)
      {
         // Dictionary<string, List<PayStub>> stubs = new Dictionary<string, List<PayStub>>();
         string identifier = string.Empty;

         foreach (DataRow row in dt.Rows) {

            try {

               identifier = row["EE_ID"].ToString();
               // DateTime checkDate = DateTime.Parse(row["Check_Date"].ToString());
               DateTime end = DateTime.Parse(row["Period_End"].ToString());// checkDate.AddDays(-5); //  ////
               DateTime start = DateTime.Parse(row["Period_Begin"].ToString());//end.AddDays(-13); // // //

               /////////////setup bi-monthly times////////////////////
               //if (checkDate.Day <= 16) {
               //   end = new DateTime(checkDate.Year != 1 ? checkDate.Year : checkDate.Year - 1, checkDate.Month == 1 ? 12 : checkDate.Month - 1, DateTime.DaysInMonth(checkDate.Year, checkDate.Month == 1 ? 12 : checkDate.Month - 1));
               //   start = new DateTime(checkDate.Year != 1 ? checkDate.Year : checkDate.Year - 1, checkDate.Month == 1 ? 12 : checkDate.Month - 1, 16);
               //}
               //else {
               //   start = new DateTime(checkDate.Year, checkDate.Month, 1);
               //   end = new DateTime(checkDate.Year, checkDate.Month, 15);
               //}

               #region pay data on a single line
               //double regRate = 0;// !String.IsNullOrEmpty(row["Rate"].ToString().Replace("$","")) ? Double.Parse(row["Rate"].ToString().Replace("$", "")) : 0;

               //double regPay = row["Regular_Wages"].ToString() != string.Empty ? Double.Parse(row["Regular_Wages"].ToString().Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "").Replace("$", "")) : 0;
               //double regHrs = row["Reg_hrs"].ToString() != string.Empty ? Double.Parse(row["Reg_hrs"].ToString()) : 0;

               //double otHrs = !String.IsNullOrEmpty(row["Overtime_Hours_Total"].ToString()) ? Double.Parse(row["Overtime_Hours_Total"].ToString()) : 0;
               //double otPay = !String.IsNullOrEmpty(row["OT_Wages"].ToString().Replace("$", "")) ? Double.Parse(row["OT_Wages"].ToString().Replace("$", "")) : 0;

               //double dblOTHrs = !String.IsNullOrEmpty(row["DoubleTime_Hours"].ToString()) ? Double.Parse(row["DoubleTime_Hours"].ToString()) : 0;
               //double dblOTPay = !String.IsNullOrEmpty(row["DoubleTime_Earnings"].ToString().Replace("$", "")) ? Double.Parse(row["DoubleTime_Earnings"].ToString().Replace("$", "")) : 0;

               //double pnltyHrs = !String.IsNullOrEmpty(row["Premium_Hours"].ToString()) ? Double.Parse(row["Premium_Hours"].ToString()) : 0;
               //double pnltyTPay = !String.IsNullOrEmpty(row["Premium_Earnings"].ToString()) ? Double.Parse(row["Premium_Earnings"].ToString()) : 0;

               //double bonus = !String.IsNullOrEmpty(row["B_Bonus_earnings"].ToString()) ? Double.Parse(row["B_Bonus_earnings"].ToString()) : 0;
               //       if (bonus == 0)
               //           bonus = !String.IsNullOrEmpty(row["B_Bonus_earnings2"].ToString()) ? Double.Parse(row["B_Bonus_earnings2"].ToString()) : 0;
               //  double commissions = !String.IsNullOrEmpty(row["Commission_Total_Amount"].ToString()) ? Double.Parse(row["Commission_Total_Amount"].ToString()) : 0;
               #endregion

               #region pay data on multiple lines
               double regRate = 0;
               double regHrs = 0;
               double regPay = 0;
               double otHrs = 0;
               double otPay = 0;
               double dblOTHrs = 0;
               double dblOTPay = 0;
               double pnltyHrs = 0;
               double pnltyTPay = 0;
               double bonus = 0;

               if (row["code"].ToString().ToUpper() == "HOURLY" || row["code"].ToString().ToUpper().Contains("REG")) {
                  regPay = Double.Parse(row["pay"].ToString().Trim(' ').Trim('$').Trim('(').Trim(')').Trim('$').Replace(",", ""));
                  if (regPay < 0)
                     continue;
                  if (row["hrs"].ToString() == String.Empty)
                     continue;
                  regHrs = Double.Parse(row["hrs"].ToString());
               }

               if (row["code"].ToString().ToUpper() == ("OVERTIME") || row["code"].ToString().ToUpper() == ("OT")) {
                  otHrs = !String.IsNullOrEmpty(row["hrs"].ToString()) ? Double.Parse(row["hrs"].ToString()) : 0;
                  otPay = !String.IsNullOrEmpty(row["pay"].ToString()) ? Double.Parse(row["pay"].ToString().Trim('$').Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "")) : 0;
               }

               if (row["code"].ToString().ToUpper() == "DT" | row["code"].ToString().ToUpper().Contains("DOUBLE")) {
                  dblOTHrs = !String.IsNullOrEmpty(row["hrs"].ToString()) ? Double.Parse(row["hrs"].ToString()) : 0;
                  dblOTPay = !String.IsNullOrEmpty(row["pay"].ToString()) ? Double.Parse(row["pay"].ToString().Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "")) : 0;
               }

               if (row["code"].ToString().ToUpper().Contains("MEAL")) {
                  pnltyHrs = !String.IsNullOrEmpty(row["hrs"].ToString()) ? Double.Parse(row["hrs"].ToString()) : 0;
                  pnltyTPay = !String.IsNullOrEmpty(row["pay"].ToString()) ? Double.Parse(row["pay"].ToString().Trim('$').Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "")) : 0;
               }

               if (row["code"].ToString().ToUpper().Contains("BONUS") || row["code"].ToString().ToUpper() == ("B")) {
                  //  pnltyHrs = !String.IsNullOrEmpty(row["hrs"].ToString()) ? Double.Parse(row["hrs"].ToString()) : 0;
                  bonus = !String.IsNullOrEmpty(row["pay"].ToString()) ? Double.Parse(row["pay"].ToString().Trim('$').Trim(' ').Trim('(').Trim(')').Trim('$').Replace(",", "")) : 0;
               }
               #endregion

               if (!stubs.ContainsKey(identifier)) {
                  stubs[identifier] = new List<PayStub>();
               }
               if (regHrs > 0) {
                  if (regPay > 0)
                     regRate = regPay / regHrs;
                  else if (otPay > 0 && otHrs > 0)
                     regRate = (otPay / otHrs) / 1.5;
               }

               #region unused for calculating mid and end
               //  DateTime endOfMonth = new DateTime(end.Year, end.Month, DateTime.DaysInMonth(end.Year, end.Month));
               //  DateTime midMonth = new DateTime(end.Year, end.Month, 15);
               //
               //determine if one week back from paycheck is closer to the end of the month or mid-month
               //if (Math.Abs((end - endOfMonth).TotalDays) < Math.Abs((end - midMonth).TotalDays) || end.Day == 1) {
               //   if (end.Day == 1) {
               //      endOfMonth = endOfMonth.AddDays(-1);
               //   }

               //   end = endOfMonth;
               //   begin = new DateTime(end.Year, end.Month, 16);
               //}
               //else {

               //   end = midMonth;
               //   begin = new DateTime(end.Year, end.Month, 1);
               //}
               #endregion

               PayStub stub = new PayStub() {
                  identifier = identifier,

                  periodBegin = start,
                  periodEnd = end,
                  //checkDate = checkDate,
                  regHrs = regHrs,
                  regPay = regPay,
                  regRate = regRate, //regHrs > 0 ? regPay / regHrs : 0,//regRate,//
                  bonus = bonus,
                  //commissions = commissions,
                  otRate = otHrs != 0 ? otPay / otHrs : 0,
                  otHrs = otHrs,
                  otPay = otPay,
                  doubltOtRate = dblOTHrs != 0 ? dblOTPay / dblOTHrs : 0,
                  doubleOtHrs = dblOTHrs,
                  doubleOtPay = dblOTPay,
                  penaltyHrs = pnltyHrs,
                  penaltyPay = pnltyTPay
               };

               if (stubs[identifier].Contains(stub)) {
                  int pos = stubs[identifier].IndexOf(stub);
                  stubs[identifier][pos].Merge(stub);
               } else
                  stubs[identifier].Add(stub);

            } catch (Exception e) {
               throw new Exception("Bad Wage Data Present");
            }
         }

         foreach (KeyValuePair<string, List<PayStub>> entry in stubs) {
            for (int pos = 0; pos < entry.Value.Count; pos++) {
               var stub = entry.Value[pos];
               if (stub.regHrs == 0 && stub.bonus == 0 && stub.commissions == 0) {

                  if (stub.otHrs > 0 && stub.regPay > 0) {
                     stub.regRate = stub.otRate / 1.5;
                     stub.regHrs = stub.regPay / stub.regRate;
                  } else {
                     entry.Value.RemoveAt(pos);
                     pos--;
                  }
               }
            }
         }

         return stubs;
      }

      /// <summary>
      /// Single Line Timedata parser, from 2 up to 8 punches on a line, that hacve clocks for lunch
      /// </summary>
      /// <param name="dt"></param>
      /// <returns></returns>
      public Dictionary<string, List<Timecard>> ConvertDataToDictWithLunches(DataTable dt)
      {
         Dictionary<string, List<Timecard>> cards = new Dictionary<string, List<Timecard>>();
         Timecard t = new Timecard();

         Dictionary<int, string> lookups = new Dictionary<int, string>();
         lookups[0] = "TimeIn";
         lookups[1] = "LunchOut";
         lookups[2] = "LunchIn";
         lookups[3] = "TimeOut";
         lookups[4] = "TimeIn2";
         lookups[5] = "LunchOut2";
         lookups[6] = "LunchIn2";
         lookups[7] = "TimeOut2";

         foreach (DataRow row in dt.Rows) {

            t = new Timecard();

            DateTime currentDate = DateTime.Parse(row["WorkDate"].ToString());
            string currentID = row["EmployeeID"].ToString();

            t.identifier = currentID;
            t.shiftDate = currentDate;


            int pos1 = 0;
            int pos2 = 1;

            if (t.shiftDate.Value.Day == 26 && t.shiftDate.Value.Month == 1) {
               int pause = 0;
            }

            try {

               t.regHrsListed = row["RegularHours"].ToString() == String.Empty ? 0 : Double.Parse(row["RegularHours"].ToString());
               t.otListed = row["OvertimeHours"].ToString() == String.Empty ? 0 : Double.Parse(row["OvertimeHours"].ToString());
               t.dtListed = row["DbltimeHours"].ToString() == String.Empty ? 0 : Double.Parse(row["DbltimeHours"].ToString());

               while (pos2 < 8) {
                  while (pos1 < 8 && row[lookups[pos1]].ToString() == String.Empty)
                     pos1++;

                  while (pos2 <= pos1 || row[lookups[pos2]].ToString() == String.Empty) {
                     pos2 += 1;
                     if (pos2 == 8)
                        break;
                  }
                  if (pos2 == 8)
                     break;


                  t.timepunches.Add(new Timepunch {
                     datetime = DateTime.Parse(currentDate.ToShortDateString() + " " + row[lookups[pos1]].ToString())
                  });

                  t.timepunches.Add(new Timepunch {
                     datetime = DateTime.Parse(currentDate.ToShortDateString() + " " + row[lookups[pos2]].ToString())
                  });

                  pos1 = pos2 + 1;

               }
            } catch (Exception e) {
               throw new Exception("Unexpected values");
            }

            if (!cards.ContainsKey(currentID)) {
               cards[currentID] = new List<Timecard>();
            }

            if (t.timepunches.Count > 0)
               cards[currentID].Add(t);

            t.AnalyzeTimeCard();

         }

         return cards;
      }

      /// <summary>
      /// Converts the datatable of pucnhes to timecards, for entries that do not have a shift date
      /// </summary>
      /// <param name="dt"></param>
      /// <returns></returns>
      public Dictionary<string, List<Timecard>> ConvertDataToDictNoShiftDate(DataTable dt)
      {
         Dictionary<string, List<Timecard>> cards = new Dictionary<string, List<Timecard>>();
         Timecard t = new Timecard();

         string identifier = string.Empty;
         DateTime shiftdate = DateTime.Now;
         DateTime currentDate = DateTime.Now;

         currentDate = DateTime.Now;

         foreach (DataRow row in dt.Rows) {

            shiftdate = DateTime.Parse(row["ShifDate"].ToString()).Date;
            string currentID = row["EE_ID"].ToString().ToUpper();

            if (currentID == identifier && currentDate.Date == shiftdate.Date) {
               //if it is the same person and date, do not add it to dictionary yet
            } else {
               if (!cards.ContainsKey(identifier)) {
                  cards[identifier] = new List<Timecard>();
               }

               if (t.timepunches.Count > 0) {
                  cards[identifier].Add(t);
                  t.AnalyzeTimeCard();
               }

               shiftdate = currentDate;
               identifier = currentID;

               t = new Timecard();
               t.identifier = identifier;
               t.shiftDate = shiftdate;
            }

            try {

               if (row["Start"].ToString() == String.Empty || row["Stop"].ToString() == String.Empty) {
                  t.invalid = true;
               } else {
                  t.timepunches.Add(new TimeIn {
                     datetime = DateTime.Parse(row["Start"].ToString())
                     , hrsListed = Double.Parse(row["hours"].ToString())
                  });

                  t.timepunches.Add(new Timeout {
                     datetime = DateTime.Parse(row["Stop"].ToString())
                  });
               }
            } catch (Exception e) {

            }
         }
         return cards;
      }



      public Dictionary<string, List<Timecard>> ConvertDataToDict(Dictionary<string, List<Timecard>> cards, DataTable dt)
      {
         Timecard t = new Timecard();

         string identifier = "-1";
         DateTime shiftdate = DateTime.Now;
         DateTime currentDate = DateTime.Now;

         HashSet<string> clocks = new HashSet<string>();
         int badTimecards = 0;

         foreach (DataRow row in dt.Rows) {

            shiftdate = DateTime.TryParse(row["ShiftDate"].ToString(), out DateTime dtx) == true ? dtx : DateTime.Now;
            string currentID = string.Empty;

            if (row["EE_ID"].ToString().Trim() != string.Empty) {
               currentID = row["EE_ID"].ToString().Replace("HBS", "").Trim().ToString();
               //  int val = int.Parse(currentID);
               // currentID = val.ToString();
            } else
               currentID = identifier;

            #region for testing
            if (identifier.Trim() == "ZELADA. SAMANTHA" && currentDate.Year == 2020 && currentDate.Month == 4 && currentDate.Day == 13) {
               // && currentDate.Year == 2020 && currentDate.Month == 4 && currentDate.Day == 1) {
               int pause = 0;
            }
            if (currentID == "14395372" && shiftdate.Year == 2020 && shiftdate.Month == 8 && shiftdate.Day == 13) {
               int pause = 0;
            }
            #endregion

            #region move day back
            //if(row["Day"].ToString().Trim() == string.Empty) {
            //   if(currentDate.AddDays(1).Date == shiftdate.Date) {
            //      DateTime da = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString());
            //      DateTime db = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time"].ToString());
            //      if(db < da) {
            //         shiftdate = shiftdate.AddDays(-1);
            //      }
            //   }
            //}
            #endregion
            if (row["In_time"].ToString() == string.Empty)
               continue;

            //check to make sure shift on the same day is not a split shift
            TimeSpan timeBetweenPunches = TimeSpan.FromMinutes(30);
            if (t.timepunches.Count > 0)
               timeBetweenPunches = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString()).Subtract(t.timepunches[t.timepunches.Count - 1].datetime);

            //If shift is the same date AND the punches are less than 6 hours apart
            if (currentID == identifier && timeBetweenPunches.TotalHours == 0) {

            } else if (currentID == identifier && currentDate.Date == shiftdate.Date && timeBetweenPunches.TotalHours < 3 && (timeBetweenPunches.TotalHours >= 0 || (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24)) && cards.ContainsKey(t.identifier)) {
               //TODO: OR ID is a mATCH and day is next is plus 1 and time is less than 2 hours
            } else if (currentID == identifier && currentDate.AddDays(1).Date == shiftdate.Date &&
               ((timeBetweenPunches.TotalHours < 2.5 && timeBetweenPunches.TotalHours > 0) ||
               (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24) //||
            //   (timeBetweenPunches.TotalHours > 24 && timeBetweenPunches.TotalHours <= 25.5)
               )) { //split shift
               int pause = 0;
            } else {
               currentDate = shiftdate;
               identifier = currentID;

               // if we haven't come across this employee             
               if (currentID != null && !cards.ContainsKey(currentID)) {
                  cards[currentID] = new List<Timecard>();
               }

               if (t.timepunches.Count > 0) {
                  t.AnalyzeTimeCard();

                  if (t.totalHrsActual.TotalHours < 0 || t.totalHrsActual.TotalHours > 15) {
                     badTimecards++;
                  } else
                     cards[t.identifier].Add(t);
               }

               shiftdate = currentDate;
               identifier = currentID;

               t = new Timecard();
               t.identifier = identifier;
               t.shiftDate = shiftdate;
            }

            try {

               if (row["In_time"].ToString() == String.Empty || row["Out_time"].ToString() == String.Empty) {
                  t.invalid = true;
               } else {

                  string val = currentID + shiftdate.ToShortDateString() + row["In_time"].ToString() + row["Out_time"].ToString();
                  if (!clocks.Contains(val))
                     clocks.Add(val);
                  else {
                     // throw new Exception("Duplicates present");
                     duplicates++;
                     continue;
                  }
                  double hrs = 0;
                  //double mealVar = Double.TryParse(row["MP"].ToString(),  out double hrs) == true ? hrs : 0;
                  //if (mealVar == 1) {
                  //   t.paidMealPremium = true;
                  //}

                  t.timepunches.Add(new TimeIn() {
                     datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString()),
                     // clockType = c,
                     //isRestBreak = row["Cat"].ToString() == "BREAK" ? true : false,
                     hrsListed = Double.TryParse(row["Reg_hrs"].ToString(), out hrs) == true ? hrs : 0,
                     //  otHrsListed = Double.TryParse(row["OT"].ToString(), out hrs) == true ? hrs : 0,
                     //   dblOtListed = Double.TryParse(row["DT"].ToString(), out  hrs) == true ? hrs : 0,
                  });

                  #region subsequent punches
                  //if (row["In_time2"].ToString() != String.Empty && row["Out_time2"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn() {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time2"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn() {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time2"].ToString()),
                  //   });
                  //}

                  //if (row["In_time3"].ToString() != String.Empty && row["Out_time3"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn() {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time3"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn() {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time3"].ToString()),
                  //   });
                  //}

                  //if (row["In_time4"].ToString() != String.Empty && row["Out_time4"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn() {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time4"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn() {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time4"].ToString()),
                  //   });
                  //}

                  //if (row["In_time5"].ToString() != String.Empty && row["Out_time5"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time5"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time5"].ToString()),
                  //   });
                  //}
                  #endregion

                  t.timepunches.Add(new Timeout() {
                     datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time"].ToString())
                  });

                  if (t.timepunches[t.timepunches.Count - 1].datetime < t.timepunches[t.timepunches.Count - 2].datetime) { //If Timeout is before Timein

                     TimeSpan tDiff = t.timepunches[t.timepunches.Count - 1].datetime.Subtract(t.timepunches[t.timepunches.Count - 2].datetime);
                     if (Math.Abs(tDiff.TotalHours) > 12) { //Has the wrong Meridian AM/PM value
                        t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(24);
                     } else //showing the previous date when they clocked out over night
                        t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(12);
                  }

               }
            } catch (Exception e) {
               throw new Exception(e.Message);
            }
         }

         //if (badTimecards > 10)
         //   throw new Exception("Issues with total hours - shifts out of order");

         return cards;
      }

      public Dictionary<string, List<Timecard>> ConvertDataToDict(DataTable dt)
      {
         Dictionary<string, List<Timecard>> cards = new Dictionary<string, List<Timecard>>();
         Timecard t = new Timecard();

         string identifier = "-1";
         DateTime shiftdate = DateTime.Now;
         DateTime currentDate = DateTime.Now;

         HashSet<string> clocks = new HashSet<string>();
         int badTimecards = 0;

         foreach (DataRow row in dt.Rows) {

            shiftdate = DateTime.TryParse(row["ShiftDate"].ToString(), out DateTime dtx) == true ? dtx : DateTime.Now;
            string currentID = string.Empty;

            if (row["EE_ID"].ToString().Trim() != string.Empty) {
               currentID = row["EE_ID"].ToString().Replace("HBS", "").Trim().ToString().ToUpper();
               //  int val = int.Parse(currentID);
               // currentID = val.ToString();
            } else
               currentID = identifier;

            #region for testing
            if (identifier.Trim() == "ZELADA. SAMANTHA" && currentDate.Year == 2020 && currentDate.Month == 4 && currentDate.Day == 13) {
               // && currentDate.Year == 2020 && currentDate.Month == 4 && currentDate.Day == 1) {
               int pause = 0;
            }
            if (currentID == "14395372" && shiftdate.Year == 2020 && shiftdate.Month == 8 && shiftdate.Day == 13) {
               int pause = 0;
            }
            #endregion

            #region move day back
            //if(row["Day"].ToString().Trim() == string.Empty) {
            //   if(currentDate.AddDays(1).Date == shiftdate.Date) {
            //      DateTime da = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString());
            //      DateTime db = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time"].ToString());
            //      if(db < da) {
            //         shiftdate = shiftdate.AddDays(-1);
            //      }
            //   }
            //}
            #endregion
            if (row["In_time"].ToString() == string.Empty)
               continue;

            //check to make sure shift on the same day is not a split shift
            TimeSpan timeBetweenPunches = TimeSpan.FromMinutes(30);
            if (t.timepunches.Count > 0)
               timeBetweenPunches = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString()).Subtract(t.timepunches[t.timepunches.Count - 1].datetime);

            //If shift is the same date AND the punches are less than 6 hours apart
            if (currentID == identifier && timeBetweenPunches.TotalHours == 0) {

            } else if (currentID == identifier && currentDate.Date == shiftdate.Date && timeBetweenPunches.TotalHours < 3 && (timeBetweenPunches.TotalHours >= 0 || (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24)) && cards.ContainsKey(t.identifier)) {
               //TODO: OR ID is a mATCH and day is next is plus 1 and time is less than 2 hours
            } else if (currentID == identifier && currentDate.AddDays(1).Date == shiftdate.Date &&
               ((timeBetweenPunches.TotalHours < 2.5 && timeBetweenPunches.TotalHours > 0) ||
               (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24) //||
            //   (timeBetweenPunches.TotalHours > 24 && timeBetweenPunches.TotalHours <= 25.5)
               )) { //split shift
               int pause = 0;
            } else {
               currentDate = shiftdate;
               identifier = currentID;

               // if we haven't come across this employee             
               if (currentID != null && !cards.ContainsKey(currentID)) {
                  cards[currentID] = new List<Timecard>();
               }

               if (t.timepunches.Count > 0) {
                  t.AnalyzeTimeCard();

                  if (t.totalHrsActual.TotalHours < 0 || t.totalHrsActual.TotalHours > 24) {
                     badTimecards++;
                  } else
                     cards[t.identifier].Add(t);
               }

               shiftdate = currentDate;
               identifier = currentID;

               t = new Timecard();
               t.identifier = identifier;
               t.shiftDate = shiftdate;
            }

            try {

               if (row["In_time"].ToString() == String.Empty || row["Out_time"].ToString() == String.Empty) {
                  t.invalid = true;
               } else {

                  string val = currentID + shiftdate.ToShortDateString() + row["In_time"].ToString() + row["Out_time"].ToString();
                  if (!clocks.Contains(val))
                     clocks.Add(val);
                  else {
                     // throw new Exception("Duplicates present");
                     duplicates++;
                     continue;
                  }
                  double hrs = 0;
                  //double mealVar = Double.TryParse(row["MP"].ToString(),  out double hrs) == true ? hrs : 0;
                  //if (mealVar == 1) {
                  //   t.paidMealPremium = true;
                  //}

                  t.timepunches.Add(new TimeIn() {
                     datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString()),
                     // clockType = c,
                     //isRestBreak = row["Cat"].ToString() == "BREAK" ? true : false,
                     hrsListed = Double.TryParse(row["Reg_hrs"].ToString(), out hrs) == true ? hrs : 0,
                    //   otHrsListed = Double.TryParse(row["OT"].ToString(), out hrs) == true ? hrs : 0,
                    //    dblOtListed = Double.TryParse(row["DT"].ToString(), out  hrs) == true ? hrs : 0,
                  });

                  t.timepunches.Add(new Timeout() {
                     datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time"].ToString())
                  });

                  #region subsequent punches
                  //if (row["In_time2"].ToString() != String.Empty && row["Out_time2"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time2"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time2"].ToString()),
                  //   });
                  //}

                  //if (row["In_time3"].ToString() != String.Empty && row["Out_time3"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time3"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time3"].ToString()),
                  //   });
                  //}

                  //if (row["In_time4"].ToString() != String.Empty && row["Out_time4"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time4"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time4"].ToString()),
                  //   });
                  //}

                  //if (row["In_time5"].ToString() != String.Empty && row["Out_time5"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time5"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time5"].ToString()),
                  //   });
                  //}
                  #endregion

                  if (t.timepunches[t.timepunches.Count - 1].datetime < t.timepunches[t.timepunches.Count - 2].datetime) { //If Timeout is before Timein

                     TimeSpan tDiff = t.timepunches[t.timepunches.Count - 1].datetime.Subtract(t.timepunches[t.timepunches.Count - 2].datetime);
                     if (Math.Abs(tDiff.TotalHours) > 12) { //Has the wrong Meridian AM/PM value
                        t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(24);
                     } else //showing the previous date when they clocked out over night
                        t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(12);
                  }

               }
            } catch (Exception e) {
               throw new Exception(e.Message);
            }
         }

         if (badTimecards > 10)
            throw new Exception("Issues with total hours - shifts out of order");

         return cards;
      }

      public Dictionary<string, List<Timecard>> ConvertDataToDictKronos(DataTable dt)
      {
         Dictionary<string, List<Timecard>> cards = new Dictionary<string, List<Timecard>>();
         Timecard t = new Timecard();

         string identifier = "-1";
         DateTime shiftdate = DateTime.Now;
         DateTime currentDate = DateTime.Now;

         HashSet<string> clocks = new HashSet<string>();

         int badShifts = 0;


         //    foreach (DataRow row in dt.Rows) {
         for (int rowCnt = 0; rowCnt < dt.Rows.Count; rowCnt++) {
            DataRow dataRow = dt.Rows[rowCnt];
            DataRow row = dataRow;


            //   var f = dt.Rows[5][""].ToString();

            shiftdate = DateTime.TryParse(row["TimePunch"].ToString(), out DateTime dtx) == true ? dtx : DateTime.Now;
            string currentID = string.Empty;

            if (row["EE_ID"].ToString().Trim() != string.Empty) {
               currentID = row["EE_ID"].ToString().Replace("HBS", "").Trim().ToString();
               int val = int.Parse(currentID);
               currentID = val.ToString();
            } else
               currentID = identifier;

            #region for testing
            if (identifier.Trim() == "ZELADA. SAMANTHA" && currentDate.Year == 2020 && currentDate.Month == 4 && currentDate.Day == 13) {
               // && currentDate.Year == 2020 && currentDate.Month == 4 && currentDate.Day == 1) {
               int pause = 0;
            }
            if (currentID == "00075" && shiftdate.Year == 2020 && shiftdate.Month == 4 && shiftdate.Day == 13) {
               int pause = 0;
            }
            #endregion

            #region move day back
            //if(row["Day"].ToString().Trim() == string.Empty) {
            //   if(currentDate.AddDays(1).Date == shiftdate.Date) {
            //      DateTime da = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString());
            //      DateTime db = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time"].ToString());
            //      if(db < da) {
            //         shiftdate = shiftdate.AddDays(-1);
            //      }
            //   }
            //}
            #endregion
            if (row["TimePunch"].ToString() == string.Empty)
               continue;

            //check to make sure shift on the same day is not a split shift
            TimeSpan timeBetweenPunches = TimeSpan.FromMinutes(30);
            if (t.timepunches.Count > 0)
               timeBetweenPunches = DateTime.Parse(row["TimePunch"].ToString()).Subtract(t.timepunches[t.timepunches.Count - 1].datetime);

            //If shift is the same date AND the punches are less than 6 hours apart
            if (t.timepunches.Count % 2 != 0 && currentID == identifier && currentDate.Date == shiftdate.Date && timeBetweenPunches.TotalHours < 9 && (timeBetweenPunches.TotalHours >= 0 || (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24)) && cards.ContainsKey(t.identifier)) {
               //TODO: OR ID is a mATCH and day is next is plus 1 and time is less than 2 hours
            } else if (t.timepunches.Count % 2 == 0 && currentID == identifier && currentDate.Date == shiftdate.Date && timeBetweenPunches.TotalHours < 3.5 && (timeBetweenPunches.TotalHours >= 0 || (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24)) && cards.ContainsKey(t.identifier)) {
               //TODO: OR ID is a mATCH and day is next is plus 1 and time is less than 2 hours
            } else if (currentID == identifier && currentDate.AddDays(1).Date == shiftdate.Date && ((timeBetweenPunches.TotalHours < 3.5 && timeBetweenPunches.TotalHours > 0) || (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24))) { //split shift
               int pause = 0;
            } else {
               currentDate = shiftdate;
               identifier = currentID;

               // if we haven't come across this employee             
               if (currentID != null && !cards.ContainsKey(currentID)) {
                  cards[currentID] = new List<Timecard>();
               }

               if (t.timepunches.Count > 0 && t.timepunches.Count % 2 == 0) {
                  t.AnalyzeTimeCard();
                  if (t.totalHrsActual.TotalHours > 2)
                     cards[t.identifier].Add(t);
                  else {
                     badShifts++;
                     rowCnt--;
                  }

               }

               shiftdate = currentDate;
               identifier = currentID;

               t = new Timecard();
               t.identifier = identifier;
               t.shiftDate = shiftdate;
            }

            try {

               if (row["TimePunch"].ToString() == String.Empty) {
                  t.invalid = true;
               } else {

                  string val = currentID + shiftdate.ToShortDateString() + row["TimePunch"].ToString();
                  if (!clocks.Contains(val))
                     clocks.Add(val);
                  else {
                     //throw new Exception("Duplicates present");
                     duplicates++;
                     continue;
                  }

                  if (t.timepunches.Count % 2 == 0) {
                     t.timepunches.Add(new TimeIn() {
                        datetime = DateTime.Parse(row["TimePunch"].ToString()),

                     });
                  } else {
                     t.timepunches.Add(new Timeout() {
                        datetime = DateTime.Parse(row["TimePunch"].ToString())
                     });
                  }


                  //if (t.timepunches[t.timepunches.Count - 1].datetime < t.timepunches[t.timepunches.Count - 2].datetime) { //If Timeout is before Timein

                  //   TimeSpan tDiff = t.timepunches[t.timepunches.Count - 1].datetime.Subtract(t.timepunches[t.timepunches.Count - 2].datetime);
                  //   if (Math.Abs(tDiff.TotalHours) > 12) { //Has the wrong Meridian AM/PM value
                  //      t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(24);
                  //   }
                  //   else //showing the previous date when they clocked out over night
                  //      t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(12);
                  //}

               }
            } catch (Exception e) {
               throw new Exception(e.Message);
            }
         }

         return cards;
      }
      public Dictionary<string, List<Timecard>> ConvertDataToDict(DataTable dt, bool hasShiftDate)
      {
         Dictionary<string, List<Timecard>> cards = new Dictionary<string, List<Timecard>>();
         Timecard t = new Timecard();

         string identifier = "-1";
         DateTime shiftdate = DateTime.Now;
         DateTime currentDate = DateTime.Now;

         HashSet<string> clocks = new HashSet<string>();

         foreach (DataRow row in dt.Rows) {

            shiftdate = DateTime.TryParse(row["ShiftDate"].ToString(), out DateTime dtx) == true ? dtx : DateTime.Now;
            string currentID = string.Empty;

            if (row["EE_ID"].ToString().Trim() != string.Empty) {
               currentID = row["EE_ID"].ToString().Trim().ToString();
            } else
               currentID = identifier;

            #region for testing
            if (identifier.Trim() == "ZELADA. SAMANTHA" && currentDate.Year == 2020 && currentDate.Month == 4 && currentDate.Day == 13) {
               // && currentDate.Year == 2020 && currentDate.Month == 4 && currentDate.Day == 1) {
               int pause = 0;
            }
            if (currentID == "00075" && shiftdate.Year == 2020 && shiftdate.Month == 4 && shiftdate.Day == 13) {
               int pause = 0;
            }
            #endregion

            #region move day back
            //if(row["Day"].ToString().Trim() == string.Empty) {
            //   if(currentDate.AddDays(1).Date == shiftdate.Date) {
            //      DateTime da = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString());
            //      DateTime db = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time"].ToString());
            //      if(db < da) {
            //         shiftdate = shiftdate.AddDays(-1);
            //      }
            //   }
            //}
            #endregion
            if (row["In_time"].ToString() == string.Empty)
               continue;

            //check to make sure shift on the same day is not a split shift
            TimeSpan timeBetweenPunches = TimeSpan.FromMinutes(30);
            if (t.timepunches.Count > 0)
               timeBetweenPunches = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString()).Subtract(t.timepunches[t.timepunches.Count - 1].datetime);

            //If shift is the same date AND the punches are less than 6 hours apart
            if (currentID == identifier && currentDate.Date == shiftdate.Date && timeBetweenPunches.TotalHours < 3 && (timeBetweenPunches.TotalHours >= 0 || (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24)) && cards.ContainsKey(t.identifier)) {
               //TODO: OR ID is a mATCH and day is next is plus 1 and time is less than 2 hours
            } else if (currentID == identifier && currentDate.AddDays(1).Date == shiftdate.Date && ((timeBetweenPunches.TotalHours < 2.5 && timeBetweenPunches.TotalHours > 0) || (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24))) { //split shift
               int pause = 0;
            } else {
               currentDate = shiftdate;
               identifier = currentID;

               // if we haven't come across this employee             
               if (currentID != null && !cards.ContainsKey(currentID)) {
                  cards[currentID] = new List<Timecard>();
               }

               if (t.timepunches.Count > 0) {
                  cards[t.identifier].Add(t);
                  t.AnalyzeTimeCard();
               }

               shiftdate = currentDate;
               identifier = currentID;

               t = new Timecard();
               t.identifier = identifier;
               t.shiftDate = shiftdate;
            }

            try {

               if (row["In_time"].ToString() == String.Empty || row["Out_time"].ToString() == String.Empty) {
                  t.invalid = true;
               } else {

                  string val = currentID + shiftdate.ToShortDateString() + row["In_time"].ToString() + row["Out_time"].ToString();
                  if (!clocks.Contains(val))
                     clocks.Add(val);
                  else {
                     throw new Exception("Duplicates present");
                     // continue;
                  }
                  double hrs = 0;
                  //double mealVar = Double.TryParse(row["MP"].ToString(),  out double hrs) == true ? hrs : 0;
                  //if (mealVar == 1) {
                  //   t.paidMealPremium = true;
                  //}

                  t.timepunches.Add(new TimeIn() {
                     datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString()),
                     // clockType = c,
                     //isRestBreak = row["Cat"].ToString() == "BREAK" ? true : false,
                     //  hrsListed = Double.TryParse(row["Reg_hrs"].ToString(), out hrs) == true ? hrs : 0,
                     //  otHrsListed = Double.TryParse(row["OT"].ToString(), out hrs) == true ? hrs : 0,
                     //   dblOtListed = Double.TryParse(row["DT"].ToString(), out  hrs) == true ? hrs : 0,
                  });

                  t.timepunches.Add(new Timeout() {
                     datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time"].ToString())
                  });

                  #region subsequent punches
                  //if (row["In_time2"].ToString() != String.Empty && row["Out_time2"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time2"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time2"].ToString()),
                  //   });
                  //}

                  //if (row["In_time3"].ToString() != String.Empty && row["Out_time3"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time3"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time3"].ToString()),
                  //   });
                  //}

                  //if (row["In_time4"].ToString() != String.Empty && row["Out_time4"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time4"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time4"].ToString()),
                  //   });
                  //}

                  //if (row["In_time5"].ToString() != String.Empty && row["Out_time5"].ToString() != String.Empty) {
                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time5"].ToString()),
                  //   });

                  //   t.timepunches.Add(new TimeIn()
                  //   {
                  //      datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time5"].ToString()),
                  //   });
                  //}
                  #endregion

                  //if (t.timepunches[t.timepunches.Count - 1].datetime < t.timepunches[t.timepunches.Count - 2].datetime) { //If Timeout is before Timein

                  //   TimeSpan tDiff = t.timepunches[t.timepunches.Count - 1].datetime.Subtract(t.timepunches[t.timepunches.Count - 2].datetime);
                  //   if (Math.Abs(tDiff.TotalHours) > 12) { //Has the wrong Meridian AM/PM value
                  //      t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(24);
                  //   }
                  //   else //showing the previous date when they clocked out over night
                  //      t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(12);
                  //}

               }
            } catch (Exception e) {
               throw new Exception(e.Message);
            }
         }

         return cards;
      }

      public Dictionary<string, List<Timecard>> ConvertDataToDictFromPDF(DataTable dt)
      {
         Dictionary<string, List<Timecard>> cards = new Dictionary<string, List<Timecard>>();
         Timecard t = new Timecard();

         string identifier = int.Parse(dt.Rows[0]["EE_ID"].ToString()).ToString();
         DateTime shiftdate = DateTime.Now;
         DateTime currentDate = DateTime.Parse(dt.Rows[0]["ShiftDate"].ToString());
         t.shiftDate = shiftdate;
         t.identifier = identifier;

         foreach (DataRow row in dt.Rows) {

            string currentID = string.Empty;

            if (row["EE_ID"] != null && row["EE_ID"].ToString().Trim() != string.Empty) {
               currentID = int.Parse(row["EE_ID"].ToString().Trim()).ToString();
            } else
               currentID = identifier;

            if (row["In_time"] == null || row["In_time"].ToString().Trim() == string.Empty)
               continue;

            //if(row["ShiftDate"] != null && row["ShiftDate"].ToString() != string.Empty)
            shiftdate = DateTime.TryParse(row["ShiftDate"].ToString(), out DateTime dtx) == true ? dtx : DateTime.Now;

            //  shiftdate = DateTime.TryParse(row["ShiftDate"].ToString(), out DateTime var);

            #region for testing
            if (identifier.Trim() == "1524") {// && currentDate.Year == 2020 && currentDate.Month == 4 && currentDate.Day == 1) {
               int pause = 0;

               if (currentID == "337" && shiftdate.Year == 2021 && shiftdate.Month == 1 && shiftdate.Day == 11) {
                  pause = 0;
               }
            }
            #endregion

            //check to make sure shift on the same day is not a split shift
            TimeSpan timeBetweenPunches = TimeSpan.FromMinutes(30);
            if (t.timepunches.Count > 0)
               timeBetweenPunches = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString()).Subtract(t.timepunches[t.timepunches.Count - 1].datetime);

            //If shift is the same date AND the punches are less than 6 hours apart
            if (currentID == identifier && currentDate.Date == shiftdate.Date && timeBetweenPunches.TotalHours < 4 && (timeBetweenPunches.TotalHours >= 0 || (timeBetweenPunches.TotalHours <= -23 && timeBetweenPunches.TotalHours >= -24))) {
               //TODO: OR ID is a mATCH and day is next is plus 1 and time is less than 2 hours
            } else if (currentID == identifier && currentDate.AddDays(-1).Date == shiftdate.Date && timeBetweenPunches.TotalHours < 2.5 && timeBetweenPunches.TotalHours > 0) { //split shift
               int pause = 0;
            } else {
               currentDate = shiftdate;
               identifier = currentID;

               if (t.identifier != null && !cards.ContainsKey(t.identifier)) {
                  cards[t.identifier] = new List<Timecard>();
               }

               if (t.timepunches.Count > 0) {
                  cards[t.identifier].Add(t);
                  t.AnalyzeTimeCard();
               }

               shiftdate = currentDate;
               identifier = currentID;

               t = new Timecard();
               t.identifier = identifier;
               t.shiftDate = shiftdate;
            }

            try {

               if (row["In_time"].ToString() == String.Empty || row["Out_time"].ToString() == String.Empty) {
                  t.invalid = true;
               } else {

                  t.timepunches.Add(new TimeIn() {
                     datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["In_time"].ToString()),
                     // clockType = c,
                     //hrsListed = Double.TryParse(row["Reg_hrs"].ToString(), out double hrs) == true ? hrs : 0,
                     // otHrsListed = Double.TryParse(row["OT_Hours"].ToString(), out hrs) == true ? hrs : 0,
                     //dblOtListed = Double.TryParse(row["DT"].ToString(), out  hrs) == true ? hrs : 0,
                  });

                  t.timepunches.Add(new Timeout() {
                     datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Out_time"].ToString())
                  });

                  //if (t.timepunches[t.timepunches.Count - 1].datetime < t.timepunches[t.timepunches.Count - 2].datetime) { //If Timeout is before Timein

                  //   TimeSpan tDiff = t.timepunches[t.timepunches.Count - 1].datetime.Subtract(t.timepunches[t.timepunches.Count - 2].datetime);
                  //   if (Math.Abs(tDiff.TotalHours) > 12) { //Has the wrong Meridian AM/PM value
                  //      t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(24);
                  //   }
                  //   else //showing the previous date when they clocked out over night
                  //      t.timepunches[t.timepunches.Count - 1].datetime = t.timepunches[t.timepunches.Count - 1].datetime.AddHours(12);
                  //}
               }
            } catch (Exception e) {

            }
         }

         return cards;
      }
      public Dictionary<string, List<Timecard>> ConvertDataToDictPeykar(DataTable dt)
      {
         Dictionary<string, List<Timecard>> cards = new Dictionary<string, List<Timecard>>();
         Timecard t = new Timecard();
         //t.invalid = true;
         string identifier = string.Empty;
         DateTime shiftdate = DateTime.Now;

         foreach (DataRow row in dt.Rows) {
            DateTime currentDate = DateTime.Now;

            if (row["Date"].ToString() == String.Empty) {
               currentDate = shiftdate; //use the previous date if the datetime DNE
            } else
               currentDate = DateTime.Parse(row["Date"].ToString());

            string currentID = row["EMP_ID"].ToString();

            if (currentID == identifier && currentDate.Date == shiftdate.Date) {

            } else {
               if (!cards.ContainsKey(identifier)) {
                  cards[identifier] = new List<Timecard>();
               }

               if (t.timepunches.Count > 0)
                  cards[identifier].Add(t);
               t.AnalyzeTimeCard();

               shiftdate = currentDate;
               identifier = currentID;

               t = new Timecard();
               t.identifier = identifier;
               t.shiftDate = shiftdate;
            }

            try {
               t.timepunches.Add(new TimeIn() {
                  datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Start"].ToString())
               });

               t.timepunches.Add(new Timeout() {
                  datetime = DateTime.Parse(shiftdate.ToShortDateString() + " " + row["Stop"].ToString())
               });
            } catch (Exception e) {

            }

         }

         return cards;
      }

      public bool AddTimecards(string dbName, string identifier, DateTime shiftDate, DateTime clockIn, DateTime clockOut, int lp, double hrs, string payCode, double breakamt)
      {
         var storedProcedure = new StoredProcedure() { StoredProcedureName = dbName + ".dbo " + ".[Timecards.Add]" };

         storedProcedure.Parameters.Add(new StoredProcedureParameter("@ShiftDate", ParameterType.DBDateTime, shiftDate));
         storedProcedure.Parameters.Add(new StoredProcedureParameter("@ClockIn", ParameterType.DBDateTime, clockIn));
         storedProcedure.Parameters.Add(new StoredProcedureParameter("@ClockOut", ParameterType.DBDateTime, clockOut));
         // storedProcedure.Parameters.Add(new StoredProcedureParameter("@Hours", ParameterType.DBDouble, hrs));
         storedProcedure.Parameters.Add(new StoredProcedureParameter("@Identifier", ParameterType.DBString, identifier));
         // storedProcedure.Parameters.Add(new StoredProcedureParameter("@Break", ParameterType.DBDouble, breakamt));
         try {
            storedProcedure.ExecuteNonQuery();
            return true;
         } catch (Exception e) {
            return false;
         }
      }

      public bool AddFlightHistory(string icao, DateTime date, string tailNumber)
      {
         var storedProcedure = new StoredProcedure() { StoredProcedureName = "Flights.dbo" + ".[Flights.AddHistory]" };

         storedProcedure.Parameters.Add(new StoredProcedureParameter("@ICAO", ParameterType.DBString, icao));
         storedProcedure.Parameters.Add(new StoredProcedureParameter("@ArrivalDate", ParameterType.DBDateTime, date));
         storedProcedure.Parameters.Add(new StoredProcedureParameter("@TailNumber", ParameterType.DBString, tailNumber));
         try {
            storedProcedure.ExecuteNonQuery();
            return true;
         } catch (Exception e) {
            return false;
         }
      }
   }
}
