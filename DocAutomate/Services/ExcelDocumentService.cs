using ClosedXML.Excel;
using System;
using System.IO;

namespace DocAutomate.Services
{
    internal sealed class ExcelDocumentService
    {
        public static bool IsValidNamePart(string part)
        {
            return !string.IsNullOrWhiteSpace(part)
                && part.IndexOfAny(Path.GetInvalidFileNameChars()) < 0
                && !part.EndsWith(".", StringComparison.Ordinal)
                && !part.EndsWith(" ", StringComparison.Ordinal);
        }

        public string GetDestination(string sourceFile, string[] nameParts, DateTime selectedDate)
        {
            return Path.Combine(Path.GetDirectoryName(sourceFile), string.Join("_", nameParts) + "_" + selectedDate.ToString("yyMMdd", System.Globalization.CultureInfo.InvariantCulture) + ".xlsx");
        }

        public void Generate(string sourceFile, string destination)
        {
            using (var workbook = new XLWorkbook(sourceFile))
            using (var generatedDocument = new MemoryStream())
            {
                workbook.SaveAs(generatedDocument);
                generatedDocument.Position = 0;
                using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write))
                    generatedDocument.CopyTo(output);
            }
        }
    }
}


