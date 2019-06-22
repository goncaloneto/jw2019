using System;
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

class Program
{
    Microsoft.Office.Interop.Word.Application WordApplication;
    Microsoft.Office.Interop.Excel.Application XLApplication;
    Document Document;

    bool IsFirst;
    bool IsLast;
    string NextTripInSlot = String.Empty;
    string PreviousTripInSlot = String.Empty;

    Workbook Workbook;
    _Worksheet Worksheet;

    static void Main()
    {
        Program program = new Program();

        int g = typeof(Hotels).GetProperties()
                             .Select(field => field.Name)
                             .ToList().Count;
        var hotelsList = new ExcelMapper("bustrips.xlsx").Fetch<Hotels>();
        var puls = new ExcelMapper("PUL.xlsx").Fetch<PUL>();
        var trips = new ExcelMapper("bustrips.xlsx").Fetch<BusTrip>();
        var slots = new ExcelMapper("Assignments.xlsx").Fetch<Slot>();
        var volunteers = new ExcelMapper("TLBC.xlsx").Fetch<Volunteer>();
        var busIds = new ExcelMapper("BUSID.xlsx").Fetch<BusID>();
        

        trips = trips.Where(x => !x.SlotName.ToUpper().Contains("CONV".ToUpper()));

        List<string> locationsDone = new List<string>();
        program.StartWordApp();

        string currentLocation;
        string currentDay;

        foreach (BusTrip busTrip in trips)
        {
            int water = 0;
            int startBuses = 0;
            int midBuses = 0;

            currentLocation = busTrip.Location;

            // Slot may have been done before
            if (locationsDone.Contains(currentLocation))
            {
                continue;
            }

            List<BusTrip> locationTrips = trips.Where(x => x.Location.Equals(currentLocation)).ToList();

            locationTrips.Sort(new StartTimeComparer());

            currentDay = locationTrips.First().StartTime.ToString("dd/MM/yyyy");

            Console.WriteLine($"Location: {currentLocation}");
            Console.WriteLine($"Day: {currentDay}");

            program.OpenDocument("input.docx");
            program.ActivateDocument();
            program.CopyTable();

            int busIndex = 1;

            for (int i = 0; i < locationTrips.Count; i++)
            {
                if (!currentDay.Equals(locationTrips[i].StartTime.ToString("dd/MM/yyyy")))
                {
                    program.HeaderFindAndReplace("{TOTALWATER}", water.ToString());
                    program.HeaderFindAndReplace("{BUSCOUNT}", busIndex - 1);
                    program.HeaderFindAndReplace("{FIRSTBUSCOUNT}", startBuses);
                    program.HeaderFindAndReplace("{MIDBUSCOUNT}", midBuses);

                    program.SaveAs($"{currentLocation}_{currentDay.Replace("/", "")}");
                    //program.SaveAsPDF($"{currentLocation}");
                    program.CloseDocument();

                    currentDay = locationTrips[i].StartTime.ToString("dd/MM/yyyy");
                    Console.WriteLine($"Day: {currentDay}");
                    busIndex = 1;
                    water = 0;
                    startBuses = 0;
                    midBuses = 0;

                    program.OpenDocument("input.docx");
                    program.ActivateDocument();
                    program.CopyTable();
                }

                if (busIndex > 1)
                {
                    program.PasteTable();
                }

                if (((busIndex % 3) == 0) && !busIndex.Equals(program.CountTotalBus(locationTrips[i].StartTime, locationTrips)))
                {
                    program.Document.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                }


                BusTrip x = locationTrips.ElementAt(i);

                if (busIndex == 1)
                {
                    program.HeaderFindAndReplace("{LOCATION}", program.GetNameByCode(puls.ToList(), x.Location));
                    program.HeaderFindAndReplace("{DATE}", x.StartTime.ToString("dd/MM/yyyy"));
                    program.FooterFindAndReplace("{LOCATION}", program.GetNameByCode(puls.ToList(), x.Location));
                    program.FooterFindAndReplace("{DATE}", x.StartTime.ToString("dd/MM/yyyy"));
                }

                program.FindAndReplace("{LOCATION}", program.GetAddressByCode(puls.ToList(), x.Location));
                program.FindAndReplace("{SLOTNAME}", x.SlotName);
                var busid = busIds.FirstOrDefault(y => y.SlotName.Equals(x.SlotName));
                program.FindAndReplace("{BUSID}", busid == null ? "N/A" : busid.BUSID);
                program.FindAndReplace("{ACTIVITYNAME}", x.ActivityName);
                program.FindAndReplace("{LEAVETIME}", x.StartTime.ToString("dd/MM/yyyy hh:mm"));
                program.FindAndReplace("{DELEGATES}", x.Delegates);
                program.FindAndReplace("{ACTIVITYNAME}", x.ActivityName);
                program.FindAndReplace("{BUSINDEX}", busIndex);
                program.FindAndReplace("{OBS}", x.Obs);
                program.FindAndReplace("{TOURLEADER}", program.GetTLBySlot(slots.ToList(), volunteers.ToList(), x.SlotName));
                program.FindAndReplace("{BUSCAPTAIN}", program.GetBCBySlot(slots.ToList(), volunteers.ToList(), x.SlotName));


                Hotels hotel = hotelsList.ElementAt(trips.ToList().IndexOf(x));
                string hotelsCell = String.Empty;
                Type fieldsType = typeof(Hotels);
                PropertyInfo[] fields = fieldsType.GetProperties();
                for (int j = 0; j < fields.Length; j++)
                {
                    if (((int)fields[j].GetValue(hotel)) > 0)
                    {
                        hotelsCell += String.IsNullOrEmpty(hotelsCell) ? String.Empty : ",\n";
                        hotelsCell += $"{program.GetNameByCode(puls.ToList(), fields[j].Name)} ({fields[j].GetValue(hotel)})";
                    }
                }

                program.FindAndReplace("{HOTEL}", hotelsCell);

                program.CheckIndex(trips.ToList(), locationTrips[i]);

                if (program.IsFirst && program.IsLast)
                {
                    startBuses++;
                    water += 6;
                    program.FindAndReplace("{HASWATER}", "Sim");
                    program.FindAndReplace("{PREVIOUSLOCATION}", "Início de Trajeto");
                    program.FindAndReplace("{NEXTLOCATION}", "Última Paragem");
                }
                else if (program.IsFirst)
                {
                    startBuses++;
                    water += 6;
                    program.FindAndReplace("{HASWATER}", "Sim");
                    program.FindAndReplace("{PREVIOUSLOCATION}", "Início de Trajeto");
                    program.FindAndReplace("{NEXTLOCATION}", program.GetNameByCode(puls.ToList(), program.NextTripInSlot));
                }
                else if (program.IsLast)
                {
                    midBuses++;
                    program.FindAndReplace("{HASWATER}", "Não");
                    program.FindAndReplace("{NEXTLOCATION}", "Última Paragem");
                    program.FindAndReplace("{PREVIOUSLOCATION}", program.GetNameByCode(puls.ToList(), program.PreviousTripInSlot));
                }
                else
                {
                    midBuses++;
                    program.FindAndReplace("{HASWATER}", "Não");
                    program.FindAndReplace("{PREVIOUSLOCATION}", program.GetNameByCode(puls.ToList(), program.PreviousTripInSlot));
                    program.FindAndReplace("{NEXTLOCATION}", program.GetNameByCode(puls.ToList(), program.NextTripInSlot));
                }

                busIndex++;
            }

            locationsDone.Add(currentLocation);

            program.HeaderFindAndReplace("{TOTALWATER}", water.ToString());
            program.HeaderFindAndReplace("{BUSCOUNT}", busIndex - 1);
            program.HeaderFindAndReplace("{FIRSTBUSCOUNT}", startBuses);
            program.HeaderFindAndReplace("{MIDBUSCOUNT}", midBuses);

            program.SaveAs($"{currentLocation}_{currentDay.Replace("/", "")}");
            //program.SaveAsPDF($"{currentLocation}");
            program.CloseDocument();
        }

        //for (int i = 0; i < nomes.Count; i++)
        //{
        //    program.OpenDocument("input.docx");
        //    program.ActivateDocument();
        //    program.FindAndReplace("{NAME}", nomes.ElementAt(i));
        //    program.FindAndReplace("{DESIGNACAO}", "Same for everyone");
        //    program.SaveAs($"Test{i}");
        //    program.SaveAsPDF($"Test{i}");
        //    program.CloseDocument();
        //    program.SendEmail("goncalomadeiraneto@gmail.com", "basquet7GMru", emails.ElementAt(i), "Test", program.GetEmailBody(), $"Test{i}.pdf");
        //    program.DeleteFile($"Test{i}.pdf");
        //    program.DeleteFile($"Test{i}.docx");
        //}

        program.QuitWordApp();
    }

    public string GetTLBySlot(List<Slot> slots, List<Volunteer> volunteers, string slot) => GetVolunteerBySlot(slots, volunteers, slot, "AT_TL");

    public string GetBCBySlot(List<Slot> slots, List<Volunteer> volunteers, string slot) => GetVolunteerBySlot(slots, volunteers, slot, "TR_BC");

    public string GetVolunteerBySlot(List<Slot> slots, List<Volunteer> volunteers, string slotName, string usage)
    {
        var slot = slots.FirstOrDefault(x => x.SlotName.Equals(slotName) && x.Usage.Equals(usage, StringComparison.InvariantCultureIgnoreCase));

        if (usage.Equals("AT_TL") && (slotName.Equals("BB02XSEASEC-CC01") || slotName.Equals("BB27CC01.STL")))
        {
            return "Desloca-se ao local";
        }

        if (slot == null && usage.Equals("TR_BC"))
        {
            return "N/A";
        }

        if (slot == null)
        {
            return "Não há Tour Leader";
        }

        if (slotName.StartsWith("OC") && usage.Equals("AT_TL"))
        {
            return "Desloca-se ao local";
        }

        Volunteer v = volunteers.FirstOrDefault(x => x.Email.Equals(slot.Email));
        string name = $"{System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(slot.VolunteerName)} {System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(slot.VolunteerSurname)}";
        return $"{name} ({v.Mobile})";
    }

    public string GetNameByCode(List<PUL> puls, string code)
    {
        PUL p = puls.FirstOrDefault(x => x.Code.Equals(code));
        return p == null ? code : p.Name;
    }

    public string GetAddressByCode(List<PUL> puls, string code)
    {
        PUL p = puls.FirstOrDefault(x => x.Code.Equals(code));
        return p == null ? code : p.Address;
    }

    public int CountTotalBus(DateTime date, List<BusTrip> locationTrips)
    {
        int i = locationTrips.Count(x => IsSameDay(x.StartTime, date));
        return i;
    }

    public bool IsSameDay(DateTime a, DateTime b) => a.Year.Equals(b.Year) && a.Month.Equals(b.Month) && a.Day.Equals(b.Day);

    public void CheckIndex(List<BusTrip> trips, BusTrip trip)
    {
        trips = trips.Where(x => x.SlotName.Equals(trip.SlotName)).ToList();
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

    public void StartWordApp() => WordApplication = new Microsoft.Office.Interop.Word.Application();
    public void StartExcelApp() => XLApplication = new Microsoft.Office.Interop.Excel.Application();

    public void OpenDocument(string filename) => Document = WordApplication.Documents.Open(Path.Combine(Directory.GetCurrentDirectory(), filename), ReadOnly: false);

    public void ActivateDocument() => Document.Activate();

    public void SaveDocument() => Document.Save();
    public void CloseDocument() => Document.Close();

    public void DeleteFile(string filename) => File.Delete(Path.Combine(Directory.GetCurrentDirectory(), filename));

    public void SaveAs(string filename) => Document.SaveAs2(Path.Combine(Directory.GetCurrentDirectory(), "saved", $"{filename}.docx"));

    public void SaveAsPDF(string filename) => Document.SaveAs2(Path.Combine(Directory.GetCurrentDirectory(), $"{filename}.pdf"), WdSaveFormat.wdFormatPDF);

    public void QuitWordApp() => WordApplication.Quit();

    public void QuitExcelApp() => XLApplication.Quit();

    public void FindAndReplace(object findText, object replaceWithText)
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
        object replace = 2;
        object wrap = 1;

        //execute find and replace
        WordApplication.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
            ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
            ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
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
