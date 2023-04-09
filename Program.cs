using System.Diagnostics;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

Console.WriteLine("Check Word docs for glossary compliance");
Dictionary<string, string> glossary = new();

if (args.Length > 1)
{
    ReadGlossaryFromExcel(args[0] + Path.DirectorySeparatorChar + args[1], ref glossary);
    Console.WriteLine("Glossary read.");
    string rootPath = args[0];
    TraverseDirectory(rootPath, ProcessFile);
    return 0;
}
else
    Console.WriteLine("ERROR: Please provide a folder path as argument.");
return 1;

void ReadGlossaryFromExcel(string excelFileName, ref Dictionary<string, string> glossary)
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
                    if (!glossary.TryAdd(rememberA.Trim(), value.Trim()))
                    {
                        Console.WriteLine($"Duplicate glossary entry: {rememberA} - {value}");
                    }
                    break;
                }
                Console.WriteLine($"Something went wrong with glossary entry {cell.CellValue.Text}: Either source or target are empty.");
            }
        }
    }
    Console.WriteLine($"{glossary.Count} glossary entries read.");
    document.Close();
}



void TraverseDirectory(string path, Action<string> action)
{
    foreach (string file in Directory.GetFiles(path))
    {
        if (file.ToUpperInvariant().Contains("_EN."))
        {
            action(file);
        }
    }

    foreach (string directory in Directory.GetDirectories(path))
    {
        TraverseDirectory(directory, action);
    }
}

void ProcessFile(string filePath)
{
    Debug.WriteLine(filePath);
    List<Text> textsSource = new();
    using WordprocessingDocument docSource = WordprocessingDocument.Open(filePath, false);
    var bodySource = docSource.MainDocumentPart.Document.Body;
    textsSource.AddRange(bodySource.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
            .Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));

    //count up all glossary entries in the source document
    Dictionary<string, int> glosCountPerSourceDoc = new();
    foreach (Text text in textsSource)
    {
        foreach (string glosEntry in glossary.Keys)
        {
            if (Regex.Match(text.InnerText, "\\b" + glosEntry + "\\b").Success)
            {
                if (glosCountPerSourceDoc.TryGetValue(glosEntry, out int value))
                    value++;
                else
                    glosCountPerSourceDoc.Add(glosEntry, 1);
            }
        }
    }

    //Now count the glossary target side in the target document
    string targetFileName = filePath.Replace("_EN.", "_ES.");
    if (!File.Exists(targetFileName))
    {
        Console.WriteLine($"ERROR: Target file {targetFileName} not found.");
        return;
    }
    List<Text> textsTarget = new();
    using WordprocessingDocument docTarget = WordprocessingDocument.Open(targetFileName, false);
    var bodyTarget = docTarget.MainDocumentPart.Document.Body;
    textsTarget.AddRange(bodyTarget.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
            .Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));

    Dictionary<string, int> glosCountPerTargetDoc = new();
    foreach (Text text in textsTarget)
    {
        foreach (string glosEntry in glosCountPerSourceDoc.Keys)
        {
            
            if (Regex.Match(text.InnerText , "\\b" + glossary[glosEntry] + "\\b", RegexOptions.IgnoreCase).Success)
            {
                if (glosCountPerTargetDoc.TryGetValue(glosEntry, out int value))
                    value++;
                else
                    glosCountPerTargetDoc.Add(glosEntry, 1);
            }
        }
    }
    Console.WriteLine("=================================================");
    Console.WriteLine($"{Path.GetFileName(filePath)}\t{Path.GetFileName(targetFileName)}");
    Console.WriteLine("-------------------------------------------------");
    int matchCount = 0;
    foreach (string glosEntry in glosCountPerSourceDoc.Keys)
    {
        if (glosCountPerTargetDoc.TryGetValue(glosEntry, out int value) && (value == glosCountPerTargetDoc[glosEntry]))
        {
            matchCount++;
        }
        else
        {
            int glosTargetCount = 0;
            if (glosCountPerTargetDoc.ContainsKey(glosEntry)) glosTargetCount = glosCountPerTargetDoc[(glosEntry)];
            Console.WriteLine($"Glossary mismatch:\t{glosEntry}\t{glosCountPerSourceDoc[glosEntry]}\t{glosTargetCount}");
        }
    }
    Console.WriteLine($"{matchCount} entries verified.\t{glosCountPerSourceDoc.Count - matchCount} entries failed.");
}


