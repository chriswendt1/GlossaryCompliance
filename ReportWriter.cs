using System.Diagnostics;

namespace GlossaryCompliance
{
    internal class ReportWriter
    {
        public string FileName { get; }
        public ReportWriter(string filename)
        {
            FileName = filename;
        }

        public void WriteLine(string text)
        {
            File.AppendAllText(FileName, text + "\r\n");
        }

        public void Show()
        {
            using Process process = new();
            process.StartInfo.FileName = "explorer";
            process.StartInfo.Arguments = FileName;
            process.Start();
        }
    }
}
