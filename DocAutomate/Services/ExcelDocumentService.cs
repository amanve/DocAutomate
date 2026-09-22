using System;
using System.IO;

namespace DocAutomate.Services
{
    internal sealed class ExcelDocumentService
    {
        public static bool IsSupportedExtension(string extension)
        {
            return string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsValidNamePart(string part)
        {
            return !string.IsNullOrWhiteSpace(part)
                && part.IndexOfAny(Path.GetInvalidFileNameChars()) < 0
                && !part.EndsWith(".", StringComparison.Ordinal)
                && !part.EndsWith(" ", StringComparison.Ordinal);
        }

        public string GetDestination(string sourceFile, string[] nameParts, DateTime selectedDate)
        {
            string extension = GetSupportedExtension(sourceFile);
            return Path.Combine(Path.GetDirectoryName(sourceFile), string.Join("_", nameParts) + "_" + selectedDate.ToString("yyMMdd", System.Globalization.CultureInfo.InvariantCulture) + extension);
        }

        public void Generate(string sourceFile, string destination, string powerpointFile = null, string part = null, string crc = null, string model = null)
        {
            string extension = GetSupportedExtension(sourceFile);
            if (!string.Equals(extension, Path.GetExtension(destination), StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(AppText.Get("The generated file must use the same Excel format as the source."), "destination");

            if (part == null && crc == null && model == null && string.IsNullOrWhiteSpace(powerpointFile) && string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(sourceFile, destination, false);
                return;
            }
            if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(AppText.Get("Workbook updates require .xlsx or .xlsm. Save the .xls workbook in a modern Excel format first."));

            byte[] updated = new SoftwareTableService().Apply(sourceFile, powerpointFile, part, crc, model);
            using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write))
                output.Write(updated, 0, updated.Length);
        }

        private static string GetSupportedExtension(string sourceFile)
        {
            string extension = Path.GetExtension(sourceFile);
            if (!IsSupportedExtension(extension))
                throw new ArgumentException(AppText.Get("Please select an .xlsx, .xlsm, or .xls file."), "sourceFile");

            return extension;
        }
    }
}
