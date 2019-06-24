using Designacoes;
using Ganss.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
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
    List<BusID> BusIDs;
    List<DropOff> DropOffs;

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
        BusIDs = new ExcelMapper("BUSID.xlsx").Fetch<BusID>().ToList();

        DOCs = Assignments.Where(x => x.Usage.Equals("AT_Drop")).ToList();
        var slotsDone = new List<string>();
        var activitiesDone = new List<string>();
        var datesDone = new List<string>();

        string date = String.Empty;

        Trips = Trips.Where(x => x.ActivityName.Contains("Beach barbecue") || x.ActivityName.Contains("Lisbon Oceanarium") || x.ActivityName.Contains("Tagus River Trip") || x.ActivityName.Contains("Bethel Tour") || x.ActivityName.Contains("Tapada de Mafra") || x.ActivityName.Contains("National Palaces")).ToList();

        DropOffs = new List<DropOff>();
        Assignments.Where(x => x.Usage.Equals("AT_Drop")).ToList().ForEach(x =>
        {
            DropOffs.Add(new DropOff($"{ToTitleCase(x.FirstName)} {ToTitleCase(x.LastName)}", TranslateActivity(x.SlotName), TranslateDay(x.SlotName), x.Email));
        });


        // for each trip // LOOP
        foreach (BusTrip trip in Trips)
        {
            if (activitiesDone.Contains(trip.ActivityName))
            {
                continue;
            }

            // Filter by trip's activity
            var tripsOfActivity = Trips.Where(x => x.ActivityName.Equals(trip.ActivityName));

            foreach (BusTrip tripOfActivity in tripsOfActivity)
            {
                if (datesDone.Contains(tripOfActivity.StartTimeDate))
                {
                    continue;
                }

                OpenDocument("input.docx");
                ActivateDocument();

                CopyTable();

                var datesOfActivity = tripsOfActivity.Where(x => x.StartTimeDate.Equals(tripOfActivity.StartTimeDate));

                foreach (BusTrip day in datesOfActivity)
                {
                    if (slotsDone.Contains(day.SlotName))
                    {
                        continue;
                    }

                    var slotsOfDay = datesOfActivity.Where(x => x.SlotName.Equals(day.SlotName)).ToList();
                    slotsOfDay.Sort(new StartTimeComparer());

                    if (slotsDone.Any())
                    {
                        Document.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);

                        PasteTable();
                    }

                    date = day.StartTimeDate;

                    HeaderFindAndReplace("{ACTIVITYNAME}", day.ActivityName);
                    HeaderFindAndReplace("{DATE}", day.StartTimeDate);
                    FooterFindAndReplace("{ACTIVITYNAME}", day.ActivityName);
                    FooterFindAndReplace("{DATE}", day.StartTimeDate);

                    FindAndReplace("{SLOTNAME}", day.SlotName);

                    int countDelegates = 0;
                    slotsOfDay.ToList().ForEach(x => countDelegates += x.Delegates);
                    FindAndReplace("{DELEGATES}", countDelegates);
                    FindAndReplace("{LASTPUL}", slotsOfDay.Last().Location);
                    FindAndReplace("{LASTPULTIME}", slotsOfDay.Last().StartTimeTime);

                    var b = BusIDs.FirstOrDefault(x => x.SlotName.Equals(day.SlotName));
                    var busid = b == null ? "N/A" : b.BUSID;
                    FindAndReplace("{BUSID}", busid);



                    var tl = Assignments.FirstOrDefault(x => x.SlotName.Equals(day.SlotName) && x.Usage.Equals("AT_TL"));
                    var name = tl == null ? "N/A" : $"{tl.FirstName} {tl.LastName}";
                    string mobile;
                    try
                    {
                        mobile = tl == null ? "N/A" : Volunteers.FirstOrDefault(x => x.Email.Equals(tl.Email)).Mobile;
                        FindAndReplace("{TOURLEADER}", $"{name} ({mobile})");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"TL Email: {tl.Email}");
                    }

                    var bc = Assignments.FirstOrDefault(x => x.SlotName.Equals(day.SlotName) && x.Usage.Equals("TR_BC"));
                    name = bc == null ? "N/A" : $"{bc.FirstName} {bc.LastName}";
                    try
                    {
                        mobile = bc == null ? "N/A" : Volunteers.FirstOrDefault(x => x.Email.Equals(bc.Email)).Mobile;
                        FindAndReplace("{BUSCAPTAIN}", $"{name} ({mobile})");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"TL Email: {bc.Email}");
                    }

                    if (tl == null)
                    {
                        Console.WriteLine(day.ActivityName + " " + day.StartTimeDate + " " + "NO TL");
                    }

                    if (bc == null)
                    {
                        Console.WriteLine(day.ActivityName + " " + day.StartTimeDate + " " + "NO BC");
                    }


                    // Delegates per slot
                    var delegatesOnSlot = Delegates.Where(x => x.SlotName.Equals(day.SlotName)).ToList();
                    // Sort by hotel
                    delegatesOnSlot.Sort(new HotelComparer());

                    // Add list of delegates to word
                    string replace = String.Empty;
                    delegatesOnSlot.ForEach(x =>
                    {
                        replace = $"{GetNameByCode(x.HotelName)} - {x.FirstName} {x.LastName} ({x.Language})\v{{DELEGATESLIST}}";
                        FindAndReplace("{DELEGATESLIST}", replace);
                    });
                    FindAndReplace("{DELEGATESLIST}", "");

                    slotsDone.Add(day.SlotName);
                }

                HeaderFindAndReplace("{BUSCOUNT}", slotsDone.Count());

                var docnames = String.Empty;

                DropOffs.Where(x => x.ActivityName.Trim().Equals(tripOfActivity.ActivityName.Trim()) && x.Date.Equals(tripOfActivity.StartTimeDate)).ToList().ForEach(x => docnames += docnames.Contains(x.VolunteerName) ? "" : docnames.Equals(String.Empty) ? x.VolunteerName : " / " + x.VolunteerName);

                HeaderFindAndReplace("{DOCNAMES}", docnames);

                DropOffs.Where(x => x.ActivityName.Trim().Equals(tripOfActivity.ActivityName.Trim()) && x.Date.Equals(tripOfActivity.StartTimeDate)).ToList().ForEach(x =>
                {
                    SaveAs($"DOC_Report");
                    SaveAsPDF("DOC_Report");

                    SendEmail("atividades.lisbon2019@gmail.com", "At@Jw2019", x.Email, $"Relatório Diário de Drop-Off {date}", GetEmailBody(date), $"DOC_Report.pdf");

                    DeleteFile("DOC_Report.pdf");
                    DeleteFile("DOC_Report.docx");

                    CloseDocument();
                });
                

                datesDone.Add(tripOfActivity.StartTimeDate);

                slotsDone.Clear();
            }

            activitiesDone.Add(trip.ActivityName);

            datesDone.Clear();
        }

        //    SendEmail("goncalomadeiraneto@gmail.com", "basquet7GMru", emails.ElementAt(i), "Test", GetEmailBody(), $"Test{i}.pdf");
        //    DeleteFile($"Test{i}.pdf");
        //    DeleteFile($"Test{i}.docx");

    }

    public string TranslateActivity(string a)
    {
        // Convention Transp.
        if (a.Contains("BB"))
            return "Beach barbecue";
        if (a.Contains("EG"))
            return "Evening Gathering";
        if (a.Contains("FS"))
            return "Field Service";
        if (a.Contains("OC"))
            return "Lisbon Oceanarium";
        if (a.Contains("NP"))
            return "National Palaces";
        if (a.Contains("TM"))
            return "Tapada de Mafra";
        if (a.Contains("BO"))
            return "Tagus River Trip";
        if (a.Contains("BETHEL"))
            return "Bethel Tour";

        return "N/A";
    }

    public string TranslateDay(string a)
    {
        if (a.Contains("24"))
            return "24/06/2019";

        if (a.Contains("25"))
            return "25/06/2019";

        if (a.Contains("26"))
            return "26/06/2019";

        if (a.Contains("27"))
            return "27/06/2019";

        if (a.Contains("28"))
            return "28/06/2019";

        if (a.Contains("29"))
            return "29/06/2019";

        if (a.Contains("30"))
            return "30/06/2019";

        if (a.Contains("01"))
            return "01/07/2019";

        if (a.Contains("02"))
            return "02/07/2019";

        return "N/A";
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
        string name = $"{System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(assignment.FirstName)} {System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(assignment.LastName)}";
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

    public string GetEmailBody(string date = "N/A") => $"<body><p style = \"font-family: verdana;font-size: 12px;\" >Prezados irmãos,</p><p style = \"font-family: verdana;font-size: 12px;\" >Enviamos em anexo a sua programação para a designação Drop-Off de dia {date} no âmbito do Congresso Interacional - Lisbon 2019. Esta informação deve ser confirmada na programação do site JW2019.org visto ser gerada automaticamente.</p><p style = \"font-family: verdana;font-size: 12px;\" >Estamos disponíveis para esclarecimentos adicionais.</p><p style = \"font-family: verdana;font-size: 12px;\" >Saudações,</p><p style = \"font-family: verdana;font-size: 12px;\" >Dept de Atividades,</p><p style = \"font-family: verdana;font-size: 12px;\" >Comissão de Hospitalidade</p></body>";


    public void SendEmail(string fromEmail, string fromPassword, string toEmail, string subject, string body, string attachment)
    {
        toEmail = "goncalomadeira.oliveira@gmail.com";
        MailAddress bcc = new MailAddress("goncalomadeiraneto@gmail.com");

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
            message.Bcc.Add(bcc);
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

    public void SaveAs(string filename) => Document.SaveAs2(Path.Combine(Directory.GetCurrentDirectory(), $"{filename}.docx"));

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
