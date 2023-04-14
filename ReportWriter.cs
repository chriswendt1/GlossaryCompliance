using System.Diagnostics;

namespace GlossaryCompliance
{
    internal class ReportWriter
    {
        public string FileName { get; }
        public ReportWriter(string filename)
        {
            FileName = filename;
            File.WriteAllText(FileName, $"{DateTime.Now.ToShortDateString()} {DateTime.Now.ToShortTimeString()}\r\n");
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
