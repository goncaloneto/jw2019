﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using Designacoes;
using Ganss.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using static Designacoes.BusTrip;
using Delegate = Designacoes.Delegate;

class Program
{
    Microsoft.Office.Interop.Word.Application WordApplication;
    Microsoft.Office.Interop.Excel.Application XLApplication;
    Document Document;

    bool IsFirst;
    bool IsLast;
    string NextTripInSlot = String.Empty;
    string PreviousTripInSlot = String.Empty;

    List<Hotel> Hotels;
    List<PUL> PULs;
    List<BusTrip> Trips;
    List<Assignment> Assignments;
    List<Assignment> DOCs;
    List<Volunteer> Volunteers;
    List<Delegate> Delegates;

    Workbook Workbook;
    _Worksheet Worksheet;

    static void Main()
    {
        Program p = new Program();
        p.StartWordApp(true);
        try
        {
            p.Run();
        }
        catch (Exception e)
        {
            throw;
        }
        finally
        {
            p.QuitWordApp();
        }


    }

    public void Run()
    {
        int g = typeof(Hotel).GetProperties()
                             .Select(field => field.Name)
                             .ToList().Count;
        Hotels = new ExcelMapper("bustrips.xlsx").Fetch<Hotel>().ToList();
        PULs = new ExcelMapper("PUL.xlsx").Fetch<PUL>().ToList();
        Trips = new ExcelMapper("bustrips.xlsx").Fetch<BusTrip>().ToList();
        Assignments = new ExcelMapper("Assignments.xlsx").Fetch<Assignment>().ToList();
        Volunteers = new ExcelMapper("Volunteers.xlsx").Fetch<Volunteer>().ToList();
        Delegates = new ExcelMapper("Delegates.xlsx").Fetch<Delegate>().ToList();

        DOCs = Assignments.Where(x => x.Usage.Equals("AT_Drop")).ToList();
        var slotsDone = new List<string>();
        var activitiesDone = new List<string>();
        var datesDone = new List<string>(); 

        // for each trip // LOOP
        foreach (BusTrip trip in Trips)
        {
            if (activitiesDone.Contains(trip.ActivityName))
            {
                continue;
            }

            // Filter by trip's activity
            var tripsOfActivity = Trips.Where(x => x.ActivityName.Equals(trip.ActivityName));

            // LOOP
            foreach (BusTrip tripOfActivity in tripsOfActivity)
            {
                if (datesDone.Contains(tripOfActivity.StartTimeDate))
                {
                    continue;
                }

                // filter by date
                var datesOfActivity = tripsOfActivity.Where(x => x.StartTimeDate.Equals(tripOfActivity.StartTimeDate));

                // for each date // LOOP
                foreach (BusTrip day in datesOfActivity)
                {
                    /////////////////////////////////// CHANGE?? //////////////////////////////////////

                    // filter by slot 
                    var slotsOfDay = datesOfActivity.Where(x => x.SlotName.Equals(day.SlotName));

                    // for each slot // LOOP
                    foreach(BusTrip slot in slotsOfDay)
                    {
                        if (slotsDone.Contains(day.SlotName))
                        {
                            continue;
                        }

                        // DO STUFF

                        // Update slots done
                        slotsDone.Add(day.SlotName);
                        
                        // END LOOP 
                    }
                }

                // update list of done dates
                datesDone.Add(tripOfActivity.StartTimeDate);

                slotsDone.Clear();

                // END LOOP 
            }
            // Update list of activities done
            activitiesDone.Add(trip.ActivityName);

            datesDone.Clear();

            // END LOOP
        }












        /////////////////////////////////////////////////////
        for (int i = 0; i < DOCs.Count; i++)
        {
            OpenDocument("input.docx");
            ActivateDocument();

            // Do it by slot 
            var currentSlot = DOCs[i].SlotName;
            // If slot is already done
            if (slotsDone.Contains(currentSlot))
            {
                continue;
            }

            // List of assignments for the current slot
            var assignmentsOfSlot = DOCs.Where(x => x.SlotName.Equals(currentSlot)).ToList();
            // Sort the list by start time
            assignmentsOfSlot.Sort(new StartComparer());

            var activityName = Trips.First(x => x.SlotName.Equals(currentSlot)).ActivityName;
            var tlName = $"{ToTitleCase(assignmentsOfSlot.First().VolunteerName)} {ToTitleCase(assignmentsOfSlot.First().VolunteerSurname)}";

            //Header
            HeaderFindAndReplace("{ACTIVITYNAME}", activityName);
            HeaderFindAndReplace("{DATE}", assignmentsOfSlot.First().StartDate);
            HeaderFindAndReplace("{BUSID}", "?????");
            HeaderFindAndReplace("{PUL}", GetNameByCode(assignmentsOfSlot.First().Location));
            HeaderFindAndReplace("{TIME}", assignmentsOfSlot.First().StartTime);
            HeaderFindAndReplace("{ADDRESS}", GetAddressByCode(assignmentsOfSlot.First().Location));

            //Footer
            FooterFindAndReplace("{ACTIVITYNAME}", activityName);
            FooterFindAndReplace("{DATE}", assignmentsOfSlot.First().StartDate);

            //First Table
            FindAndReplace("{TOURLEADER}", tlName);
            FindAndReplace("{ACTIVITYNAME}", activityName);
            FindAndReplace("{SLOTNAME}", currentSlot);
            FindAndReplace("{BUSID}", "?????");


            // Fill PULs
            var trips = Trips.Where(x => x.SlotName.Equals(currentSlot)).ToList();
            trips.Sort(new StartTimeComparer());

            foreach (BusTrip trip in trips)
            {
                // Duplicate row
                WordApplication.Selection.Collapse();
                WordApplication.Selection.Find.Execute("{LOCATION}");
                WordApplication.Selection.Rows[1].Range.Copy();
                WordApplication.Selection.Rows[1].Select();
                WordApplication.Selection.Range.Paste();
                WordApplication.Selection.Collapse();

                // ORDER MATTERS when replacing just one string
                FindAndReplace("{LOCATION}", $"[PUL] {trip.Location}", WdReplace.wdReplaceOne);
                FindAndReplace("{ARRIVAL}", "-", WdReplace.wdReplaceOne);
                FindAndReplace("{DEPARTURE}", trip.StartTimeTime, WdReplace.wdReplaceOne);
                FindAndReplace("{NDEL}", trip.Delegates, WdReplace.wdReplaceOne);

                var ppcs = Assignments.Where(x => x.Location.Equals(trip.Location) && x.StartDate.Equals(trip.StartTimeDate) && x.Usage.Equals("AT_Pick"));
                Console.WriteLine($"AT_Pick Counts: {ppcs.Count()}");
                var ppc = ppcs.FirstOrDefault();
                string name = ppc == null ? "N/A" : $"{ToTitleCase(ppc.VolunteerName)} {ToTitleCase(ppc.VolunteerSurname)}";
                string mobile = ppc == null ? "N/A" : Volunteers.First(x => x.Email.Equals(ppc.Email)).Mobile;

                FindAndReplace("{PPC}", name, WdReplace.wdReplaceOne);
                FindAndReplace("{PPCMOBILE}", mobile, WdReplace.wdReplaceOne);
            }

            FindAndReplace("{BUSCAPTAIN}", GetBCBySlot(currentSlot));

            // Delegates per slot
            var delegatesOnSlot = Delegates.Where(x => x.SlotName.Equals(currentSlot)).ToList();
            // Sort by hotel
            delegatesOnSlot.Sort(new HotelComparer());

            // Add list of delegates to word
            string replace = String.Empty;
            delegatesOnSlot.ForEach(x =>
            {
                replace = $"{GetNameByCode(x.Hotel)} - {x.Name} {x.Surname}\v{{DELEGATESLIST}}";
                FindAndReplace("{DELEGATESLIST}", replace);
            });

            // Cleanup 
            int count = CountOccurrences("PPCMOBILE");
            for (int j = 0; j < count; j++)
            {
                DeleteRowContainsExpression("{PPCMOBILE}");
            }
            FindAndReplace("{DELEGATESLIST}", "");

            SaveAs($"{activityName.Replace(" ", "")}_{assignmentsOfSlot.First().StartDate.Replace("/", "")}");
            //SaveAsPDF($"{currentLocation}");
            CloseDocument();
        }


        //    SendEmail("goncalomadeiraneto@gmail.com", "basquet7GMru", emails.ElementAt(i), "Test", GetEmailBody(), $"Test{i}.pdf");
        //    DeleteFile($"Test{i}.pdf");
        //    DeleteFile($"Test{i}.docx");

    }

    public int CountOccurrences(string s)
    {
        var wordCounter = 0;
        // Loop through all words in the document.
        for (var i = 1; i <= Document.Words.Count; i++)
            if (Document.Words[i].Text.TrimEnd() == s)
                wordCounter++;
        return wordCounter;
    }

    public string GetTLBySlot(string slot) => GetVolunteerBySlot(slot, "AT_TL");

    public string GetBCBySlot(string slot) => GetVolunteerBySlot(slot, "TR_BC");

    public string GetVolunteerBySlot(string slotName, string usage)
    {
        var assignment = Assignments.FirstOrDefault(x => x.SlotName.Equals(slotName) && x.Usage.Equals(usage, StringComparison.InvariantCultureIgnoreCase));

        if (usage.Equals("AT_TL") && (slotName.Equals("BB02XSEASEC-CC01") || slotName.Equals("BB27CC01.STL")))
        {
            return "Desloca-se ao local";
        }

        if (assignment == null && usage.Equals("TR_BC"))
        {
            return "N/A";
        }

        if (assignment == null)
        {
            return "Não há Tour Leader";
        }

        if (slotName.StartsWith("OC") && usage.Equals("AT_TL"))
        {
            return "Desloca-se ao local";
        }

        Volunteer v = Volunteers.FirstOrDefault(x => x.Email.Equals(assignment.Email));
        string name = $"{System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(assignment.VolunteerName)} {System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(assignment.VolunteerSurname)}";
        return $"{name} ({v.Mobile})";
    }

    public string ToTitleCase(string s) => System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s);

    public string GetNameByCode(string code)
    {
        PUL p = PULs.FirstOrDefault(x => x.Code.Equals(code));
        return p == null ? code : p.Name;
    }

    public string GetAddressByCode(string code)
    {
        PUL p = PULs.FirstOrDefault(x => x.Code.Equals(code));
        return p == null ? code : p.Address;
    }

    public int CountTotalBus(DateTime date)
    {
        int i = Trips.Count(x => IsSameDay(x.StartTime, date));
        return i;
    }

    public bool IsSameDay(DateTime a, DateTime b) => a.Year.Equals(b.Year) && a.Month.Equals(b.Month) && a.Day.Equals(b.Day);

    public void CheckIndex(BusTrip trip)
    {
        var trips = Trips.Where(x => x.SlotName.Equals(trip.SlotName)).ToList();
        trips.Sort(new StartTimeComparer());

        if (trips.First().Equals(trip) && trips.Last().Equals(trip))
        {
            IsFirst = true;
            IsLast = true;
        }
        else if (trips.First().Equals(trip))
        {
            IsFirst = true;
            IsLast = false;
            NextTripInSlot = trips.ElementAt(trips.IndexOf(trip) + 1).Location;
        }
        else if (trips.Last().Equals(trip))
        {
            IsFirst = false;
            PreviousTripInSlot = trips.ElementAt(trips.IndexOf(trip) - 1).Location;
            IsLast = true;
        }
        else
        {
            IsFirst = false;
            PreviousTripInSlot = trips.ElementAt(trips.IndexOf(trip) - 1).Location;
            IsLast = false;
            NextTripInSlot = trips.ElementAt(trips.IndexOf(trip) + 1).Location;
        }
    }
    Microsoft.Office.Interop.Word.Table templateTable;
    Microsoft.Office.Interop.Word.Range range;
    object oMissing = System.Reflection.Missing.Value;
    public void CopyTable()
    {
        templateTable = Document.Tables[1];
        range = templateTable.Range;
        range.SetRange(templateTable.Range.Start, templateTable.Range.End);
        range.Copy();
    }

    public void PasteTable()
    {
        //range.SetRange(templateTable.Range.End + 1, templateTable.Range.End + 1);
        //Microsoft.Office.Interop.Word.Table tableCopy = Document.Tables.Add(range, 1, 1, ref oMissing, ref oMissing);
        //tableCopy.Range.Paste();

        WordApplication.ActiveDocument.Characters.Last.Select();
        WordApplication.Selection.Collapse();
        WordApplication.Selection.Paste();
        WordApplication.Selection.TypeText("  ");
    }


    //public string GetEmailBody() => "<h1>Title</h1>" +
    //        "<p style=\"color:red; text-align:center;\">Red paragraph.</p>" +
    //        "<table>" +
    //        "<thead>" +
    //        "<tr>" +
    //        "<th>1</th><th>2</th>" +
    //        "</tr>" +
    //        "</thead>" +
    //        "<tbody>" +
    //        "<tr>" +
    //        "<td>Teste</td><td>Teste2</td>" +
    //        "</tr>" +
    //        "</tbody>" +
    //        "</table>";

    public string GetEmailBody() => "<body style=\"background-color: lightblue;\"><h1 style = \"color: white;text-align: center;\" > My First CSS Example</h1><p style = \"font-family: verdana;font-size: 20px;\" > This is a paragraph.</p></body>";


    public void SendEmail(string fromEmail, string fromPassword, string toEmail, string subject, string body, string attachment)
    {
        var smtp = new SmtpClient
        {
            Host = "smtp.gmail.com",
            Port = 587,
            EnableSsl = true,
            DeliveryMethod = SmtpDeliveryMethod.Network,
            UseDefaultCredentials = false,
            Credentials = new NetworkCredential(fromEmail, fromPassword)
        };
        using (var message = new System.Net.Mail.MailMessage(fromEmail, toEmail)
        {
            Subject = subject,
            Body = body
        })
        {
            message.IsBodyHtml = true;
            message.Attachments.Add(new Attachment(Path.Combine(Directory.GetCurrentDirectory(), attachment)));
            smtp.Send(message);
        }
    }

    public void StartWordApp(bool visible = false) => WordApplication = new Microsoft.Office.Interop.Word.Application() { Visible = visible };
    public void StartExcelApp() => XLApplication = new Microsoft.Office.Interop.Excel.Application();

    public void OpenDocument(string filename, bool visible = false) => Document = WordApplication.Documents.Open(Path.Combine(Directory.GetCurrentDirectory(), filename), ReadOnly: false);

    public void ActivateDocument() => Document.Activate();

    public void SaveDocument() => Document.Save();
    public void CloseDocument() => Document.Close();

    public void DeleteFile(string filename) => File.Delete(Path.Combine(Directory.GetCurrentDirectory(), filename));

    public void SaveAs(string filename) => Document.SaveAs2(Path.Combine(Directory.GetCurrentDirectory(), "saved", $"{filename}.docx"));

    public void SaveAsPDF(string filename) => Document.SaveAs2(Path.Combine(Directory.GetCurrentDirectory(), $"{filename}.pdf"), WdSaveFormat.wdFormatPDF);

    public void QuitWordApp() => WordApplication.Quit();

    public void QuitExcelApp() => XLApplication.Quit();

    public void FindAndReplace(object findText, object replaceWithText, WdReplace replace = WdReplace.wdReplaceAll)
    {
        //options
        object matchCase = false;
        object matchWholeWord = true;
        object matchWildCards = false;
        object matchSoundsLike = false;
        object matchAllWordForms = false;
        object forward = true;
        object format = false;
        object matchKashida = false;
        object matchDiacritics = false;
        object matchAlefHamza = false;
        object matchControl = false;
        object read_only = false;
        object visible = true;
        object wrap = 1;

        WordApplication.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
    }

    public void DeleteRowContainsExpression(string s)
    {
        WordApplication.Selection.Collapse();
        WordApplication.Selection.Find.Execute(s);
        WordApplication.Selection.Rows[1].Delete();
    }

    public void HeaderFindAndReplace(object findText, object replaceWithText)
    {
        foreach (Microsoft.Office.Interop.Word.Section section in Document.Sections)
        {
            Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            headerRange.Find.Text = findText.ToString();
            headerRange.Find.Replacement.Text = replaceWithText.ToString();

            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            headerRange.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }

    public void FooterFindAndReplace(object findText, object replaceWithText)
    {
        foreach (Microsoft.Office.Interop.Word.Section section in Document.Sections)
        {
            Microsoft.Office.Interop.Word.Range footer = section.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            footer.Find.Text = findText.ToString();
            footer.Find.Replacement.Text = replaceWithText.ToString();

            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            footer.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }

    public void InitWorksheet(string filename, int sheet = 1)
    {
        Workbook = XLApplication.Workbooks.Open(Path.Combine(Directory.GetCurrentDirectory(), filename));
        Worksheet = (_Worksheet)Workbook.Sheets[sheet];
    }

    public List<string> ReadColumn(string columnHeader)
    {
        List<string> list = new List<string>();
        Microsoft.Office.Interop.Excel.Range xlRange = Worksheet.UsedRange;

        int rowIndex = 2;
        int colIndex = 1;

        foreach (Microsoft.Office.Interop.Excel.Range col in xlRange.Columns)
        {
            Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)col.Cells[1, 1];
            if (cell.Value != null && cell.Value.Equals(columnHeader))
            {
                break;
            }
            colIndex++;
        }

        for (int i = rowIndex; i <= xlRange.Rows.Count; i++)
        {
            Microsoft.Office.Interop.Excel.Range cell = xlRange.Cells[i, colIndex];
            if (cell.Value != null)
            {
                list.Add(cell.Value);
            }
        }

        return list;
    }

    //private static void WriteLineToExcel(params string[] line)
    //{
    //    Application xlApp = new Application();
    //    Workbook xlWorkbook = xlApp.Workbooks.Open(Path.Combine(Directory.GetCurrentDirectory(), filename));
    //    _Worksheet xlWorksheet = (_Worksheet)xlWorkbook.Sheets[1];
    //    Range xlRange = xlWorksheet.UsedRange;

    //    int col = ColIndex;

    //    foreach (string s in line)
    //    {
    //        xlRange.Cells[RowIndex, col++] = s;
    //    }

    //    RowIndex++;
    //    xlWorkbook.Save();
    //    xlWorkbook.Close();
    //}
}
