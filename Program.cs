using System.Diagnostics;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Xml.XPath;
using System.Xml;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using GlossaryCompliance;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

Console.WriteLine("Check Word docs for glossary compliance");
Dictionary<string, string> glossary = new();
int totalHits = 0;
int totalMisses = 0;
ReportWriter reportWriter = new("report.txt");

if (args.Length > 1)
{
    ReadGlossaryFromExcel(args[1], ref glossary);
    reportWriter.WriteLine("Glossary read.");
    string rootPath = args[0];
    TraverseDirectory(rootPath, ProcessFile);
    reportWriter.WriteLine($"Total hits: {totalHits}\tTotal misses: {totalMisses}");
    reportWriter.Show();
    return 0;
}
else
    Console.WriteLine("ERROR: Please provide a folder path as argument.");
return 1;

void ReadGlossaryFromExcel(string excelFileName, ref Dictionary<string, string> glossary)
{
    using SpreadsheetDocument document = SpreadsheetDocument.Open(excelFileName, false);
    if (document is null)
    {
        reportWriter.WriteLine($"ERROR: Could not open {excelFileName}.");
        throw new FileNotFoundException($"ERROR: Could not open {excelFileName}.");
    }
    var sheets = document.WorkbookPart.WorksheetParts;

    // Loop through each of the sheets in the spreadsheet
    foreach (var wp in sheets)
    {
        Worksheet worksheet = wp.Worksheet;

        // Loop through each of the rows in the current sheet
        var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();
        foreach (var row in rows)
        {
            // Loop through each of the cells in the current row.
            var cells = row.Elements<Cell>();
            string rememberA = null;
            foreach (var cell in cells)
            {
                if (cell.CellValue is null) continue;
                string? value = document.WorkbookPart.SharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cell.CellValue.Text)).InnerText;
                if (cell.CellReference.InnerText.StartsWith('A') && value is not null)
                {
                    rememberA = value;
                    continue;
                }
                if (cell.CellReference.InnerText.StartsWith('B') && value is not null && rememberA is not null)
                {
                    if (!glossary.TryAdd(rememberA.Trim(), value.Trim()))
                    {
                        reportWriter.WriteLine($"Duplicate glossary entry: {rememberA} - {value}");
                    }
                    break;
                }
                reportWriter.WriteLine($"Something went wrong with glossary entry {cell.CellValue.Text}: Either source or target are empty.");
            }
        }
    }
    reportWriter.WriteLine($"{glossary.Count} glossary entries read.");
    document.Close();
}

void ProcessExcel(string excelFileName)
{
    using SpreadsheetDocument document = SpreadsheetDocument.Open(excelFileName, false);
    var sheets = document.WorkbookPart.WorksheetParts;

    // Loop through each of the sheets in the spreadsheet
    foreach (var wp in sheets)
    {
        Worksheet worksheet = wp.Worksheet;

        // Loop through each of the rows in the current sheet
        var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();
        foreach (var row in rows)
        {
            // Loop through each of the cells in the current row.
            var cells = row.Elements<Cell>();
            string rememberA = null;
            foreach (var cell in cells)
            {
                string value = document.WorkbookPart.SharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cell.CellValue.Text)).InnerText;
                if (cell.CellReference.InnerText.StartsWith('A') && value is not null)
                {
                    rememberA = value;
                    continue;
                }
                if (cell.CellReference.InnerText.StartsWith('B') && value is not null && rememberA is not null)
                {

                    break;
                }
                reportWriter.WriteLine($"Something went wrong with glossary entry {cell.CellValue.Text}: Either source or target are empty.");
            }
        }
    }
    reportWriter.WriteLine($"{glossary.Count} glossary entries read.");
    document.Close();
}



void TraverseDirectory(string path, Action<string, FileType> action)
{
    foreach (string file in Directory.GetFiles(path))
    {
        if (file.ToUpperInvariant().Contains("_EN.") && (Path.GetExtension(file).ToLowerInvariant() == ".docx"))
        {
            action(file, FileType.docx);
            continue;
        }
        if (Path.GetExtension(file).ToLowerInvariant() == ".tmx")
        {
            ProcessTMX(file);
            continue;
        }
        if (Path.GetExtension(file).ToLowerInvariant() == ".xlsx")
        {
            ProcessExcel(file);
            continue;
        }
        reportWriter.WriteLine($"File skipped:\t{Path.GetFileName(file)}");
    }

    foreach (string directory in Directory.GetDirectories(path))
    {
        TraverseDirectory(directory, action);
    }
}

void ProcessTMX(string fileName)
{
    Debug.WriteLine(fileName);
    XmlDocument xmlDoc = new();
    xmlDoc.Load(fileName);
    XPathNavigator xPathNavigator = xmlDoc.CreateNavigator();
    XmlNamespaceManager xmlNamespaceManager = new(xPathNavigator.NameTable);

    // Select all TU elements
    XmlNodeList tuNodes = xmlDoc.SelectNodes("//tu");
    Dictionary<string, CountPair> glosCountPerDoc = new();
    int segmentCounter = 0;
    // Loop through each TU element and apply the method
    foreach (XmlNode tuNode in tuNodes)
    {
        segmentCounter++;
        if ((segmentCounter % 100) == 0) Console.WriteLine("T");
        // Apply your method to the TU element here
        // For example, you could extract the source and target segments:
        XmlNode segSource = tuNode.SelectSingleNode("./tuv[@xml:lang='EN-US']/seg", xmlNamespaceManager);
        string sourceText = RemoveMarkup(segSource.InnerText);

        XmlNode segTarget = tuNode.SelectSingleNode("./tuv[@xml:lang='ES-ES']/seg", xmlNamespaceManager);
        string targetText = RemoveMarkup(segTarget.InnerText);
        foreach (string glosEntry in glossary.Keys)
        {
            //Debug.Assert(!(text.InnerText.Contains("shelters") && (glosEntry=="shelters")));
            //BUGBUG This counting misses a count when the glossary term appears multiple times in the same text segment
            if (Regex.Match(sourceText, "\\b" + glosEntry + "\\b").Success)
            {
                if (glosCountPerDoc.TryGetValue(glosEntry, out CountPair countPair))
                {
                    countPair.SourceCount++;
                    glosCountPerDoc[glosEntry] = countPair;
                }
                else
                    glosCountPerDoc.Add(glosEntry, new CountPair(1, 0));
            }
        }
        foreach (string glosEntry in glosCountPerDoc.Keys)
        {
            string[] alternatives = glossary[glosEntry].Split('/');
            foreach (string alternative in alternatives)
            {
                MatchCollection matches = Regex.Matches(targetText, "\\b" + alternative.Trim() + "\\b", RegexOptions.IgnoreCase);
                if (matches.Count > 0)
                {
                    //Debug.Assert(text.InnerText.Contains("asistencia financiera"));
                    CountPair countPair = glosCountPerDoc[glosEntry];
                    countPair.TargetCount += matches.Count;
                    glosCountPerDoc[glosEntry] = countPair;
                }
            }
        }
    }
    reportWriter.WriteLine("=================================================");
    reportWriter.WriteLine($"{Path.GetFileName(fileName)}{Path.GetExtension(fileName)}");
    reportWriter.WriteLine("-------------------------------------------------");
    int matchCount = 0;
    int missCount = 0;
    foreach (string glosEntry in glosCountPerDoc.Keys)
    {
        if (glosCountPerDoc[glosEntry].IsSatisfied())
        {
            matchCount++;
        }
        else
        {
            reportWriter.WriteLine($"Glossary mismatch:\t{glosEntry}\t{glosCountPerDoc[glosEntry].SourceCount}\t{glossary[glosEntry]}\t{glosCountPerDoc[glosEntry].TargetCount}");
            missCount++;
        }
    }
    reportWriter.WriteLine($"{matchCount} entries verified.\t{missCount} entries failed.");
    totalHits += matchCount;
    totalMisses += missCount;
}


static string RemoveMarkup(string tuv)
{
    // regular expression pattern to match XML tags
    string pattern = @"<[^>]+>";

    // remove all XML tags from the string
    string plainText = Regex.Replace(tuv, pattern, "");

    //other cleanup

    plainText = plainText.Replace("\t", " ");
    plainText = plainText.Replace("• ", "");
    plainText = plainText.StartsWith("-") ? plainText[1..] : plainText;
    plainText = plainText.StartsWith("■") ? plainText[1..] : plainText;
    plainText = plainText.StartsWith("\"") ? "\"" + plainText : plainText;
    plainText = Regex.Replace(plainText, @"\.+", ".");


    // output plain text string
    return plainText.Trim();
}


void ProcessFile(string filePath, FileType fileType)
{
    Debug.WriteLine(filePath);
    Console.Write(".");
    List<Text> textsSource = new();
    using WordprocessingDocument docSource = WordprocessingDocument.Open(filePath, false);
    var bodySource = docSource.MainDocumentPart.Document.Body;
    textsSource.AddRange(bodySource.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
            .Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));

    //count up all glossary entries in the source document
    Dictionary<string, CountPair> glosCountPerDoc = new();
    foreach (Text text in textsSource)
    {
        foreach (string glosEntry in glossary.Keys)
        {
            //Debug.Assert(!(text.InnerText.Contains("shelters") && (glosEntry=="shelters")));
            //BUGBUG This counting misses a count when the glossary term appears multiple times in the same text segment
            if (Regex.Match(text.InnerText, "\\b" + glosEntry + "\\b").Success)
            {
                if (glosCountPerDoc.TryGetValue(glosEntry, out CountPair countPair))
                {
                    countPair.SourceCount++;
                    glosCountPerDoc[glosEntry] = countPair;
                }
                else
                    glosCountPerDoc.Add(glosEntry, new CountPair(1, 0));
            }
        }
    }

    //Now count the glossary target side in the target document
    string targetFileName = filePath.Replace("_EN.", "_ES.");
    if (!File.Exists(targetFileName))
    {
        reportWriter.WriteLine($"ERROR: Target file {targetFileName} not found.");
        return;
    }
    if (filePath == targetFileName)
    {
        reportWriter.WriteLine($"ERROR: Target file {targetFileName} is the same as the source file.");
        return;
    }
    List<Text> textsTarget = new();
    using WordprocessingDocument docTarget = WordprocessingDocument.Open(targetFileName, false);
    var bodyTarget = docTarget.MainDocumentPart.Document.Body;
    textsTarget.AddRange(bodyTarget.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
            .Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));

    foreach (Text text in textsTarget)
    {
        foreach (string glosEntry in glosCountPerDoc.Keys)
        {
            string[] alternatives = glossary[glosEntry].Split('/');
            foreach (string alternative in alternatives)
            {
                MatchCollection matches = Regex.Matches(text.InnerText, "\\b" + alternative.Trim() + "\\b", RegexOptions.IgnoreCase);
                if (matches.Count > 0)
                {
                    //Debug.Assert(text.InnerText.Contains("asistencia financiera"));
                    CountPair countPair = glosCountPerDoc[glosEntry];
                    countPair.TargetCount += matches.Count;
                    glosCountPerDoc[glosEntry] = countPair;
                }
            }
        }
    }
    reportWriter.WriteLine("=================================================");
    reportWriter.WriteLine($"{Path.GetFileName(filePath)}\t{Path.GetFileName(targetFileName)}");
    reportWriter.WriteLine("-------------------------------------------------");
    int matchCount = 0;
    int missCount = 0;
    foreach (string glosEntry in glosCountPerDoc.Keys)
    {
        if (glosCountPerDoc[glosEntry].IsSatisfied())
        {
            matchCount++;
        }
        else
        {
            reportWriter.WriteLine($"Glossary mismatch:\t{glosEntry}\t{glosCountPerDoc[glosEntry].SourceCount}\t{glossary[glosEntry]}\t{glosCountPerDoc[glosEntry].TargetCount}");
            missCount++;
        }
    }
    reportWriter.WriteLine($"{matchCount} entries verified.\t{missCount} entries failed.");
    totalHits += matchCount;
    totalMisses += missCount;
}


