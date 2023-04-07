using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

Console.WriteLine("Check Word docs for glossary compliance");
Dictionary<string, string> glossary = new()
{
    { "American Red Cross", "Cruz Roja Americana" }
};

if (args.Length > 0)
{
    string rootPath = args[0];
    TraverseDirectory(rootPath, ProcessFile);
    return 0;
}
else
    Console.WriteLine("ERROR: Please provide a folder path as argument.");
return 1;


void TraverseDirectory(string path, Action<string> action)
{
    foreach (string file in Directory.GetFiles(path))
    {
        if (file.ToLowerInvariant().Contains("_EN."))
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
    List<Text> textsSource = new();
    using WordprocessingDocument docSource = WordprocessingDocument.Open(filePath, true);
    var bodySource = docSource.MainDocumentPart.Document.Body;
    textsSource.AddRange(bodySource.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
            .Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));

    //count up all glossary entries in thge source document
    Dictionary<string, int> glosCountPerSourceDoc = new();
    foreach (Text text in textsSource)
    {
        foreach (string glosEntry in glossary.Keys)
        {
            if (text.InnerText.Contains(glosEntry))
            {
                glosCountPerSourceDoc[glosEntry]++;
            }
        }
    }

    //Now count the glossary target side in the target document

    List<Text> textsTarget = new();
    using WordprocessingDocument docTarget = WordprocessingDocument.Open(filePath.Replace("_EN.", "_ES."), true);
    var bodyTarget = docTarget.MainDocumentPart.Document.Body;
    textsSource.AddRange(bodyTarget.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
            .Where(text => !String.IsNullOrEmpty(text.Text) && text.Text.Length > 0));

    Dictionary<string, int> glosCountPerTargetDoc = new();
    foreach (Text text in textsTarget)
    {
        foreach(string glosEntry in glosCountPerSourceDoc.Keys)
        {
            if (text.InnerText.Contains(glossary[glosEntry]))
            {
                glosCountPerTargetDoc[glosEntry]++;
            }
        }
    }
    Console.WriteLine("==========================");
    Console.WriteLine(filePath);
    Console.WriteLine("--------------------------");
    int matchCount = 0;
    foreach (string glosEntry in glosCountPerSourceDoc.Keys)
    {
        if (glosCountPerSourceDoc[glosEntry] == glosCountPerTargetDoc[glosEntry])
        {
            matchCount++;
        }
        else
        {
            Console.WriteLine($"Glossary mismatch:\t{glosEntry}\t{glosCountPerSourceDoc[glosEntry]}\t{glosCountPerTargetDoc[glosEntry]}");
        }
    }
    Console.WriteLine($"{matchCount} entries verified.\t{glosCountPerSourceDoc.Count - matchCount} entries failed.");
}


