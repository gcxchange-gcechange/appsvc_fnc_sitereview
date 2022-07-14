using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SiteReview
{
    static class Helpers
    {
        public static List<List<string>> GenerateCSV(string csvData)
        {
            var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(GenerateStreamFromString(csvData));
            parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
            parser.SetDelimiters(new string[] { "," });

            var report = new List<List<string>>();
            while (!parser.EndOfData)
            {
                var row = parser.ReadFields().ToList();
                report.Add(row);
            }

            return report;
        }

        private static Stream GenerateStreamFromString(string str)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(str);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }
    }
}
