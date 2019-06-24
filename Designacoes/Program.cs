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
    List<Assignment> TLs;
    List<Volunteer> Volunteers;
    List<Delegate> Delegates;
    List<BusID> BusIDs;
    List<Prayer> Prayers;

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
        Prayers = new ExcelMapper("Prayers.xlsx").Fetch<Prayer>().ToList();

        TLs = Assignments.Where(x => x.Usage.Equals("AT_TL")).ToList();
        var slotsDone = new List<string>() {    "NP24XQSECSEACC01",
                                                "TM27B2LMA01",
                                                "BB25LMA01",
                                                "BB27XLGOLAR01",
                                                "TM25B1LMQ01",
                                                "BB01XLGOLMQ02",
                                                "BO26LMQ02",
                                                "NP01XQLMQLMA01",
                                                "BO27XLMLLMQ01",
                                                "NP01XQLGOLAR01",
                                                "BO26LAV01",
                                                "NP27SLML01",
                                                "BO25LAV01",
                                                "NP27SLMA01",
                                                "TM26B1LGO01",
                                                "BO27CC01",
                                                "BO25LAR01",
                                                "NP26SLMQ01",
                                                "BO27LAV02",
                                                "BO01LAV01",
                                                "NP27SLMQ02",
                                                "BO26LMA.STL",
                                                "BO26SI01",
                                                "BO01LGO01",
                                                "BB26CC01.STL",
                                                "BO01LMA01",
                                                "BB25CC01.STL",
                                                "NP01XQTVSI01",
                                                "TM25B2SI01",
                                                "NP27QLBX01",
                                                "TM01B1LGO01",
                                                "BO27SEC01",
                                                "NP26QXSEASEC01",
                                                "BB02CC01.STL",
                                                "BO27LAV01",
                                                "BO27LMA03",
                                                "NP25SLMA01",
                                                "TM25B1LAV01",
                                                "BB26LBX01",
                                                "BO27LGO01",
                                                "BO26LMQ04",
                                                "BB25LGO01",
                                                "BO26LMQ01",
                                                "BB02LMQ01",
                                                "BB26SEC01",
                                                "NP26QLGO01",
                                                "TM27B1LAR01",
                                                "NP26QLBX01",
                                                "TM27B1LGO01",
                                                "BO01CA01",
                                                "NP27SLMQ04",
                                                "TM25B2CA01",
                                                "BO26LMA03",
                                                "NP27SLMQ01",
                                                "BB01LAV01",
                                                "TM26B2LBX01",
                                                "BO26TV01",
                                                "TM27B2LMQ02",
                                                "BB27LMQ01",
                                                "BO24LML01",
                                                "TM01B1LAV01",
                                                "BO26XSICA01",
                                                "NP25XQCASI01",
                                                "TM01B1LAR01",
                                                "TM27B1LMA03",
                                                "BB25LAV01",
                                                "NP26QLAV01",
                                                "TM01B2CC01",
                                                "BO25CA01",
                                                "NP27XSCASI01",
                                                "BO24XLBXLMQ01",
                                                "TM26B2TV01",
                                                "BO24LGO02",
                                                "BB27SI01",
                                                "BB27LAV01",
                                                "NP27QLAR01",
                                                "NP27XQLMALAV01"};

        for (int i = 0; i < TLs.Count; i++)
        {
            OpenDocument("input.docx");
            ActivateDocument();

            // Do it by slot 
            var currentSlot = TLs[i].SlotName;
            // If slot is already done
            if (slotsDone.Contains(currentSlot))
            {
                continue;
            }

            // List of assignments for the current slot
            var assignmentsOfSlot = TLs.Where(x => x.SlotName.Equals(currentSlot)).ToList();
            // Sort the list by start time
            assignmentsOfSlot.Sort(new StartComparer());

            var a = Trips.FirstOrDefault(x => x.SlotName.Equals(currentSlot));

            if (a == null)
                continue;

            var activityName = a.ActivityName;
            var tlName = $"{ToTitleCase(assignmentsOfSlot.First().FirstName)} {ToTitleCase(assignmentsOfSlot.First().LastName)}";

            var b = BusIDs.FirstOrDefault(x => x.SlotName.Equals(currentSlot));
            var busid = b == null ? "N/A" : b.BUSID;

            if (b == null)
            {
                Console.WriteLine($"No BUS ID for: {currentSlot}");
            }

            //Header
            HeaderFindAndReplace("{ACTIVITYNAME}", activityName);
            HeaderFindAndReplace("{DATE}", assignmentsOfSlot.First().StartDate);
            HeaderFindAndReplace("{BUSID}", busid);
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
            FindAndReplace("{BUSID}", busid);


            // Fill PULs
            var trips = Trips.Where(x => x.SlotName.Equals(currentSlot)).ToList();
            trips.Sort(new StartTimeComparer());

            int d = 0;

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

                d += trip.Delegates;

                var ppcs = Assignments.Where(x => x.Location.Equals(trip.Location) && x.StartDate.Equals(trip.StartTimeDate) && x.Usage.Equals("AT_Pick"));
                var ppc = ppcs.FirstOrDefault();
                string name = ppc == null ? "N/A" : $"{ToTitleCase(ppc.FirstName)} {ToTitleCase(ppc.LastName)}";
                var e = ppc == null ? null : Volunteers.FirstOrDefault(x => x.Email.Equals(ppc.Email));
                string mobile = "";

                if (e == null && ppc == null)
                {
                    mobile = "N/A";
                }
                else if (e == null && ppc != null)
                {
                    Console.WriteLine("Email: " + ppc.Email);
                    Console.WriteLine("Enter mobile for (Do NOT COPY): " + ppc.Email);
                    mobile = Console.ReadLine();
                }
                else
                {
                    mobile = e.Mobile;
                }

                FindAndReplace("{PPC}", name, WdReplace.wdReplaceOne);
                FindAndReplace("{PPCMOBILE}", mobile, WdReplace.wdReplaceOne);
            }

            if (trips.Any())
            {
                // Duplicate row
                WordApplication.Selection.Collapse();
                WordApplication.Selection.Find.Execute("{LOCATION}");
                WordApplication.Selection.Rows[1].Range.Copy();
                WordApplication.Selection.Rows[1].Select();
                WordApplication.Selection.Range.Paste();
                WordApplication.Selection.Collapse();

                FindAndReplace("{LOCATION}", trips.First().ActivityName, WdReplace.wdReplaceOne);
                FindAndReplace("{ARRIVAL}", "-", WdReplace.wdReplaceOne);
                FindAndReplace("{DEPARTURE}", trips.First().ReturnTimeTime, WdReplace.wdReplaceOne);
                FindAndReplace("{NDEL}", d, WdReplace.wdReplaceOne);
                FindAndReplace("{PPC}", "-", WdReplace.wdReplaceOne);
                FindAndReplace("{PPCMOBILE}", "-", WdReplace.wdReplaceOne);
            }



            FindAndReplace("{BUSCAPTAIN}", GetBCBySlot(currentSlot));

            // Prayer
            var p = Prayers.Where(x => x.SlotName.Equals(currentSlot)).ToList();

            if (p.Any())
            {
                var p1 = p.FirstOrDefault(x => x.OptionNumber.Equals(1));
                FindAndReplace("{PRAYER1}", p1 == null ? "N/A" : $"{ToTitleCase(p1.FirstName)} {ToTitleCase(p1.LastName)} ({p1.Language})");

                var p2 = p.FirstOrDefault(x => x.OptionNumber.Equals(2));
                FindAndReplace("{PRAYER2}", p2 == null ? "N/A" : $"{ToTitleCase(p2.FirstName)} {ToTitleCase(p2.LastName)} ({p2.Language})");
            }
            else
            {
                FindAndReplace("{PRAYER1}", "N/A");
                FindAndReplace("{PRAYER2}", "N/A");
            }


            FindAndReplace("{OBS}", "Enviamos os nomes de irmãos para proferirem as orações antes de almoço. Caso o irmão assinalado com (1) não possa/aceite, convide o irmão assinalado com (2).\v");


            // Delegates per slot
            var delegatesOnSlot = Delegates.Where(x => x.SlotName.Equals(currentSlot)).ToList();
            // Sort by hotel
            delegatesOnSlot.Sort(new HotelComparer());

            // Add list of delegates to word
            string replace = String.Empty;
            delegatesOnSlot.ForEach(x =>
            {
                replace = $"{GetNameByCode(x.HotelName)} - {x.FirstName} {x.LastName} ({x.Language})\v{{DELEGATESLIST}}";
                FindAndReplace("{DELEGATESLIST}", replace);
            });

            // Cleanup 
            int count = CountOccurrences("PPCMOBILE");
            for (int j = 0; j < count; j++)
            {
                DeleteRowContainsExpression("{PPCMOBILE}");
            }
            FindAndReplace("{DELEGATESLIST}", "");

            slotsDone.Add(currentSlot);
            Console.WriteLine("Done: " + currentSlot);
            Console.WriteLine("Count: " + slotsDone.Count);

            SaveAs($"{currentSlot}");
            SaveAsPDF($"{currentSlot}");
            CloseDocument();

            SendEmail("atividades.lisbon2019@gmail.com", "At@Jw2019", assignmentsOfSlot.First().Email, $"Relatório Diário de TL {assignmentsOfSlot.First().StartDate}", GetEmailBody(assignmentsOfSlot.First().StartDate), $"{currentSlot}.pdf");

            DeleteFile($"{currentSlot}.pdf");
            DeleteFile($"{currentSlot}.docx");
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
            Console.WriteLine($"No BC for {slotName}");
            return "N/A";
        }

        if (assignment == null)
        {
            Console.WriteLine($"No TL for {slotName}");
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

    public string GetEmailBody(string date = "N/A") => $"<body><p style = \"font-family: verdana;font-size: 12px;\" >Prezados irmãos,</p><p style = \"font-family: verdana;font-size: 12px;\" >Enviamos em anexo a sua programação para a designação de Tour Leader de dia {date} no âmbito do Congresso Interacional - Lisbon 2019. Esta informação deve ser confirmada na programação do site JW2019.org visto ser gerada automaticamente.</p><p style = \"font-family: verdana;font-size: 12px;\" >Estamos disponíveis para esclarecimentos adicionais.</p><p style = \"font-family: verdana;font-size: 12px;\" >Saudações,</p><p style = \"font-family: verdana;font-size: 12px;\" >Dept de Atividades,</p><p style = \"font-family: verdana;font-size: 12px;\" >Comissão de Hospitalidade</p></body>";


    public void SendEmail(string fromEmail, string fromPassword, string toEmail, string subject, string body, string attachment)
    {
        //toEmail = "goncalomadeira.oliveira@gmail.com";
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
