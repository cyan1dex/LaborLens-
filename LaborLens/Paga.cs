using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {
   class Paga {
      public List<Period> periods = new List<Period>();

      /// <summary>
      /// Create paychecks splitting the months in two evenly
      /// </summary>
      /// <param name="start"></param>
      /// <param name="end"></param>
      public Paga(DateTime start, DateTime end)
      {
         HashSet<DateTime> set = new HashSet<DateTime>();
         // int periodLength = 15;

         DateTime current = start;

         while (current < end) {

            if (current.Day <= 15) //first half of month
               periods.Add(new Period() { start = current, end = new DateTime(current.Year, current.Month, 15) });
            else
               periods.Add(new Period() { start = new DateTime(current.Year, current.Month, 16), end = new DateTime(current.Year, current.Month, DateTime.DaysInMonth(current.Year, current.Month)) });

            if (current.Day == 16) {
               if (current.Month < 12)
                  current = new DateTime(current.Year, current.Month + 1, 1);
               else
                  current = new DateTime(current.Year + 1, 1, 1);
            } else
               current = current.AddDays(15);
         }
      }

      /// <summary>
      /// Creates pay periods from the start date of first paycheck till the end of paga
      /// </summary>
      /// <param name="start">date of first paycheck</param>
      /// <param name="end">when the paga perio ends</param>
      /// <param name="length">length of pay period</param>
      public Paga(DateTime start, DateTime end, int length)
      {
         HashSet<DateTime> set = new HashSet<DateTime>();
         // int periodLength = 15;

         DateTime current = start;

         while (current < end) {
            periods.Add(new Period() { start = current, end = current.AddDays(length - 1) });

            if (current.AddDays(length - 1).Year == 2017 && current.AddDays(length - 1).Month == 7 && current.AddDays(length - 1).Day == 8) {
               length = 14;
            }
            current = current.AddDays(length);

         }
      }

      public Paga(Dictionary<string, List<PayStub>> empStubs)
      {
         HashSet<DateTime> set = new HashSet<DateTime>();
         int periodLength = 13;

         foreach (KeyValuePair<string, List<PayStub>> entry in empStubs) {
            foreach (PayStub stub in entry.Value)
               if (stub.periodBegin != null) {
                  if (!set.Contains(stub.periodBegin.Value)) {
                     set.Add(stub.periodBegin.Value);
                     if (stub.periodEnd.Value.Subtract(stub.periodBegin.Value).TotalDays == periodLength) //Validate the period length is correct
                        periods.Add(new Period() { start = stub.periodBegin.Value, end = stub.periodEnd.Value });
                  }
               }
         }

         periods = periods.OrderBy(x => x.start).ToList();
      }

      public Paga(Dictionary<string, List<PayStub>> empStubs, DateTime start, DateTime end)
      {
         HashSet<DateTime> set = new HashSet<DateTime>();
         // int periodLength = 15;

         foreach (KeyValuePair<string, List<PayStub>> entry in empStubs) {
            foreach (PayStub stub in entry.Value)
               if (stub.periodEnd != null) {
                  if (!set.Contains(stub.periodEnd.Value)) {
                     set.Add(stub.periodEnd.Value);
                     // if (stub.periodEnd.Value.Subtract(stub.periodBegin.Value).TotalDays == periodLength) //Validate the period length is correct
                     periods.Add(new Period() { start = stub.periodEnd.Value.AddDays(-13), end = stub.periodEnd.Value });
                  }
               }
         }

         periods = periods.OrderBy(x => x.start).ToList();

         //Logic abouve is thrown out here because there are gaps, not needed when there is enough data
         //Period [] temp = { periods[0] };
         //periods = temp.ToList();
         //   //Branch out from earliest to latest

         //   Period earliest = periods[0];
         //   Period latest = periods[periods.Count() - 1];

         //   while (earliest.start > start) {
         //       Period p = new Period() { start = earliest.start.AddDays(-13), end = earliest.end.AddDays(-13) };
         //       periods.Add(p);
         //       earliest = p;
         //   }

         //   while (latest.end < end) {
         //       Period p = new Period() { start = latest.start.AddDays(13), end = latest.end.AddDays(13) };
         //       periods.Add(p);
         //       latest = p;
         //   }
         //   periods = periods.OrderBy(x => x.start).ToList();
      }


      public Dictionary<string, List<Timesheet>> PopulatePeriods(Dictionary<string, List<Timecard>> timeCards)
      {
         var timesheets = new Dictionary<string, List<Timesheet>>();
         if (timeCards == null || timeCards.Count == 0) return timesheets;

         var periodsStarts = periods?.ToArray() ?? Array.Empty<Period>();
         if (periodsStarts.Length == 0) return timesheets;

         for (int empPos = 0; empPos < timeCards.Count; empPos++) {
            var employeeCards = timeCards.ElementAt(empPos).Value;
            if (employeeCards == null || employeeCards.Count == 0) continue;

            // sort by date
            employeeCards = employeeCards.OrderBy(x => x.shiftDate).ToList();

            // find first period that can contain the first card
            int position = 0;
            var firstCard = employeeCards[0];
            while (position < periodsStarts.Length - 1 &&
                   firstCard.shiftDate.HasValue &&
                   firstCard.shiftDate.Value.Date > periodsStarts[position].end.Date) {
               position++;
            }

            // working sheet for the current period
            var sheet = new Timesheet { identifier = timeCards.ElementAt(empPos).Key };
            sheet.periodBegin = periodsStarts[position].start;
            sheet.periodEnd = periodsStarts[position].end;

            bool addedAnyInThisPeriod = false; // <- TRUE only when we actually add a card

            for (int i = 0; i < employeeCards.Count; i++) {
               var card = employeeCards[i];
               if (card?.shiftDate == null) continue;

               // if current card is after current period, advance periods until it fits
               while (position < periodsStarts.Length &&
                      card.shiftDate.Value.Date > periodsStarts[position].end.Date) {
                  // If we had cards in the current period, emit the sheet for that period
                  if (addedAnyInThisPeriod) {
                     if (!timesheets.ContainsKey(sheet.identifier))
                        timesheets[sheet.identifier] = new List<Timesheet>();

                     timesheets[sheet.identifier].Add(sheet);
                     sheet.CalculateNonStub();

                     // start a new sheet for the next period
                     sheet = new Timesheet { identifier = sheet.identifier };
                     addedAnyInThisPeriod = false;
                  }

                  // advance to next period window
                  position++;
                  if (position >= periodsStarts.Length) break;

                  // update the working sheet’s period window
                  sheet.periodBegin = periodsStarts[position].start;
                  sheet.periodEnd = periodsStarts[position].end;
               }

               if (position >= periodsStarts.Length) break;

               // If card is within current period window, add it (only if it meets your hours rule)
               if (card.shiftDate.Value.Date >= periodsStarts[position].start.Date &&
                   card.shiftDate.Value.Date <= periodsStarts[position].end.Date) {
                  // your rule: only keep cards with > 6 actual hours
                  if (card.totalHrsActual.TotalHours > 6) {
                     sheet.timeCards.Add(card);
                     addedAnyInThisPeriod = true; // <- set TRUE only when we actually added
                  }
                  // else: do nothing; DO NOT flip addedAnyInThisPeriod
               }
               // else (card before current period) cannot happen due to sorting & initial positioning
            }

            // flush the last period only if it has cards
            if (addedAnyInThisPeriod) {
               if (!timesheets.ContainsKey(sheet.identifier))
                  timesheets[sheet.identifier] = new List<Timesheet>();

               timesheets[sheet.identifier].Add(sheet);
               sheet.CalculateNonStub();
            }
         }

         return timesheets;
      }



      //public Dictionary<string, List<Timesheet>> PopulatePeriods(Dictionary<string, List<Timecard>> timeCards)
      //{
      //   Dictionary<string, List<Timesheet>> timesheets = new Dictionary<string, List<Timesheet>>();

      //   for (int empPos = 0; empPos < timeCards.Count; empPos++) {

      //      if (timeCards.ElementAt(empPos).Value.Count == 0)
      //         continue;

      //      Timesheet sheet = new Timesheet() { identifier = timeCards.ElementAt(empPos).Key };
      //      Period[] periodsStarts = periods.ToArray();

      //      #region sort timecards
      //      var employee = timeCards.ElementAt(empPos).Value;
      //      employee = employee.OrderBy(x => x.shiftDate).ToList();
      //      #endregion

      //      int position = 0;
      //      // Timecard card = timeCards.ElementAt(empPos).Value[0];
      //      Timecard card = employee[0];

      //      string id = card.identifier;

      //      if (id == "J5R009207") {
      //         int p = 0;
      //      }
      //      bool added = false;
      //      bool changed = false;

      //      while (card.shiftDate.Value.Date > periodsStarts[position].end.Date && position < periodsStarts.Count() - 1) //while the date is greater than the current date, move to next position
      //         position++;

      //      sheet.periodBegin = periodsStarts[position].start;
      //      sheet.periodEnd = periodsStarts[position].end;


      //      for (int pos = 0; pos < employee.Count; pos++) {
      //         card = employee[pos];

      //         if (position >= periodsStarts.Count())
      //            break;

      //         if (card.shiftDate <= periodsStarts[position].end) {//if current card is lte than end, than its in range

      //            if (card.totalHrsActual.TotalHours > 6)
      //               sheet.timeCards.Add(card);
      //            added = true;
      //         } else {
      //            while (position < periodsStarts.Count() && card.shiftDate > periodsStarts[position].end) {
      //               position++;
      //               changed = true;
      //            }
      //         }

      //         if (added && changed) {
      //            if (!timesheets.ContainsKey(sheet.identifier))
      //               timesheets[sheet.identifier] = new List<Timesheet>();

      //            timesheets[sheet.identifier].Add(sheet);
      //            sheet.CalculateNonStub();
      //            sheet = new Timesheet() { identifier = sheet.identifier };
      //            sheet.periodBegin = periodsStarts[position].start;
      //            sheet.periodEnd = periodsStarts[position].end;
      //            sheet.timeCards.Add(card);
      //            added = false;
      //            changed = false;
      //         }

      //      }

      //      if (added && changed == false) {
      //         if (!timesheets.ContainsKey(sheet.identifier))
      //            timesheets[sheet.identifier] = new List<Timesheet>();

      //         timesheets[sheet.identifier].Add(sheet);
      //         sheet.CalculateNonStub();
      //      } else {
      //         int pause = 0;
      //      }

      //      //sheet.Analyze();
      //   }

      //   return timesheets;
      //}
   }
}
