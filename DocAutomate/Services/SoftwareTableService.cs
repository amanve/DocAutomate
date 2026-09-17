using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace DocAutomate.Services
{
    internal sealed class SoftwareTableService
    {
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private sealed class Change
        {
            public string Group, Setting, Before, After;
        }

        private static XDocument Read(PackagePart part)
        {
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read))
                return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        }

        private static string Text(XElement element, XNamespace ns)
        {
            return string.Concat(element.Descendants(ns + "t").Select(t => t.Value));
        }

        private static List<Change> ReadChanges(string path)
        {
            if (!string.Equals(Path.GetExtension(path), ".pptx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Please select a .pptx PowerPoint file.");
            var changes = new List<Change>();
            using (Package package = Package.Open(path, FileMode.Open, FileAccess.Read))
            {
                foreach (PackagePart part in package.GetParts().Where(p => p.ContentType == "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"))
                {
                    XDocument slide = Read(part);
                    bool software = slide.Descendants(P + "sp").Any(shape =>
                        string.Equals(Text(shape, A).Trim(), "Software", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(Text(shape, A).Trim(), "Software Changes", StringComparison.OrdinalIgnoreCase));
                    if (!software) continue;
                    foreach (XElement table in slide.Descendants(A + "tbl"))
                    {
                        string group = null;
                        foreach (XElement row in table.Elements(A + "tr"))
                        {
                            var cells = row.Elements(A + "tc").ToList();
                            if (cells.Count != 4)
                                throw new InvalidOperationException("Software tables must have four columns: group, setting, As Is, To Be.");
                            string[] values = cells.Select(c => Text(c, A)).ToArray();
                            if (values.All(string.IsNullOrWhiteSpace)) continue;
                            if (values[2].Trim().Equals("As Is", StringComparison.OrdinalIgnoreCase) &&
                                values[3].Trim().Equals("To Be", StringComparison.OrdinalIgnoreCase)) continue;
                            if (!string.IsNullOrWhiteSpace(values[0])) group = values[0];
                            else if ((string)cells[0].Attribute("vMerge") != "1" && (string)cells[0].Attribute("vMerge") != "true")
                                throw new InvalidOperationException("A blank group must be part of a vertically merged cell.");
                            if (string.IsNullOrWhiteSpace(group) || string.IsNullOrWhiteSpace(values[1]) ||
                                string.IsNullOrWhiteSpace(values[2]) || string.IsNullOrWhiteSpace(values[3]))
                                throw new InvalidOperationException("Every software row needs a group, setting, As Is and To Be value.");
                            if (changes.Any(c => c.Group == group && c.Setting == values[1]))
                                throw new InvalidOperationException("Duplicate PowerPoint setting: " + group + " / " + values[1]);
                            changes.Add(new Change { Group = group, Setting = values[1], Before = values[2], After = values[3] });
                        }
                    }
                }
            }
            if (changes.Count == 0)
                throw new InvalidOperationException("No data table found on a Software or Software Changes slide.");
            return changes;
        }

        private static string CellText(XElement cell, List<string> shared)
        {
            if (cell == null) return "";
            string value = (string)cell.Element(S + "v") ?? "";
            if ((string)cell.Attribute("t") == "s") return shared[int.Parse(value, CultureInfo.InvariantCulture)];
            if ((string)cell.Attribute("t") == "inlineStr") return Text(cell, S);
            return value;
        }

        private static int Column(string address)
        {
            int col = 0;
            foreach (char c in address.TakeWhile(char.IsLetter)) col = col * 26 + c - 'A' + 1;
            return col;
        }

        private static int Row(string address)
        {
            return int.Parse(new string(address.SkipWhile(char.IsLetter).ToArray()), CultureInfo.InvariantCulture);
        }

        private static XElement GroupCell(XDocument sheet, Dictionary<string, XElement> cells, string address)
        {
            foreach (XElement merge in sheet.Descendants(S + "mergeCell"))
            {
                string[] range = ((string)merge.Attribute("ref")).Split(':');
                if (range.Length == 2 && Column(address) >= Column(range[0]) && Column(address) <= Column(range[1]) &&
                    Row(address) >= Row(range[0]) && Row(address) <= Row(range[1]))
                {
                    XElement anchor;
                    return cells.TryGetValue(range[0], out anchor) ? anchor : null;
                }
            }
            XElement result;
            return cells.TryGetValue(address, out result) ? result : null;
        }

        private static string Address(int column, int row)
        {
            string letters = "";
            while (column > 0) { column--; letters = (char)('A' + column % 26) + letters; column /= 26; }
            return letters + row.ToString(CultureInfo.InvariantCulture);
        }

        private static void HighlightChanges(Package package, IEnumerable<XElement> cells)
        {
            var changedCells = cells.ToList();
            if (changedCells.Count == 0) return;

            var stylesPart = package.GetParts().FirstOrDefault(p =>
                p.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            XDocument styles;
            if (stylesPart == null)
            {
                stylesPart = package.CreatePart(new Uri("/xl/styles.xml", UriKind.Relative),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
                package.GetPart(new Uri("/xl/workbook.xml", UriKind.Relative)).CreateRelationship(
                    new Uri("styles.xml", UriKind.Relative), TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
                styles = new XDocument(new XElement(S + "styleSheet",
                    new XElement(S + "fonts", new XElement(S + "font")),
                    new XElement(S + "fills",
                        new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "none"))),
                        new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "gray125")))),
                    new XElement(S + "borders", new XElement(S + "border")),
                    new XElement(S + "cellStyleXfs", new XElement(S + "xf")),
                    new XElement(S + "cellXfs", new XElement(S + "xf"))));
            }
            else styles = Read(stylesPart);

            XElement fonts = styles.Root.Element(S + "fonts");
            XElement fills = styles.Root.Element(S + "fills");
            XElement formats = styles.Root.Element(S + "cellXfs");
            if (fonts == null || fills == null || formats == null)
                throw new InvalidOperationException("The workbook is missing required cell formatting definitions.");

            int yellowFillId = fills.Elements(S + "fill").Count();
            fills.Add(new XElement(S + "fill", new XElement(S + "patternFill",
                new XAttribute("patternType", "solid"),
                new XElement(S + "fgColor", new XAttribute("rgb", "FFFFFF00")),
                new XElement(S + "bgColor", new XAttribute("indexed", "64")))));
            var highlightedStyles = new Dictionary<int, int>();
            foreach (XElement cell in changedCells)
            {
                int originalStyle = (int?)cell.Attribute("s") ?? 0;
                int highlightedStyle;
                if (!highlightedStyles.TryGetValue(originalStyle, out highlightedStyle))
                {
                    // Clone shared styles so unchanged cells retain their original appearance.
                    XElement format = new XElement(formats.Elements(S + "xf").ElementAt(originalStyle));
                    int fontId = (int?)format.Attribute("fontId") ?? 0;
                    XElement font = new XElement(fonts.Elements(S + "font").ElementAt(fontId));
                    font.Elements(S + "color").Remove();
                    font.Add(new XElement(S + "color", new XAttribute("rgb", "FFFF0000")));
                    format.SetAttributeValue("fontId", fonts.Elements(S + "font").Count());
                    fonts.Add(font);
                    format.SetAttributeValue("fillId", yellowFillId);
                    format.SetAttributeValue("applyFont", "1");
                    format.SetAttributeValue("applyFill", "1");
                    highlightedStyle = formats.Elements(S + "xf").Count();
                    formats.Add(format);
                    highlightedStyles.Add(originalStyle, highlightedStyle);
                }
                cell.SetAttributeValue("s", highlightedStyle);
            }
            foreach (XElement collection in styles.Root.Elements().Where(e =>
                e.Name == S + "fonts" || e.Name == S + "fills" || e.Name == S + "borders" ||
                e.Name == S + "cellStyleXfs" || e.Name == S + "cellXfs"))
                collection.SetAttributeValue("count", collection.Elements().Count());
            using (Stream output = stylesPart.GetStream(FileMode.Create, FileAccess.Write)) styles.Save(output);
        }

        public byte[] Apply(string excelPath, string powerpointPath)
        {
            var changes = ReadChanges(powerpointPath);
            using (var buffer = new MemoryStream())
            {
                using (Stream input = File.OpenRead(excelPath)) input.CopyTo(buffer);
                buffer.Position = 0;
                using (Package package = Package.Open(buffer, FileMode.Open, FileAccess.ReadWrite))
                {
                    var shared = new List<string>();
                    var sharedPart = package.GetParts().FirstOrDefault(p => p.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                    if (sharedPart != null) shared = Read(sharedPart).Descendants(S + "si").Select(s => Text(s, S)).ToList();
                    var sheets = package.GetParts().Where(p => p.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
                        .ToDictionary(p => p, p => Read(p));
                    var edits = new Dictionary<XElement, string>();
                    foreach (Change change in changes)
                    {
                        var matches = new List<XElement>();
                        foreach (XDocument sheet in sheets.Values)
                        {
                            var cells = sheet.Descendants(S + "c").ToDictionary(c => (string)c.Attribute("r"));
                            foreach (XElement cell in cells.Values.Where(c => CellText(c, shared) == change.Setting))
                            {
                                string address = (string)cell.Attribute("r");
                                int column = Column(address), row = Row(address);
                                if (column < 2 || CellText(GroupCell(sheet, cells, Address(column - 1, row)), shared) != change.Group) continue;
                                XElement target;
                                if (cells.TryGetValue(Address(column + 1, row), out target)) matches.Add(target);
                            }
                        }
                        string label = change.Group + " / " + change.Setting;
                        if (matches.Count != 1)
                            throw new InvalidOperationException("Expected one Excel match for " + label + "; found " + matches.Count + ".");
                        XElement match = matches[0];
                        if (match.Element(S + "f") != null)
                            throw new InvalidOperationException("Cannot replace a formula for " + label + ".");
                        if (CellText(match, shared) != change.Before)
                            throw new InvalidOperationException("Excel value for " + label + " does not match PowerPoint As Is (" + change.Before + ").");
                        if (change.Before != change.After) edits.Add(match, change.After);
                    }
                    foreach (var edit in edits)
                    {
                        XElement cell = edit.Key;
                        cell.Elements(S + "v").Remove();
                        cell.Elements(S + "is").Remove();
                        double number;
                        if (double.TryParse(edit.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out number) &&
                            !double.IsNaN(number) && !double.IsInfinity(number) &&
                            number.ToString("G", CultureInfo.InvariantCulture) == edit.Value)
                        {
                            cell.SetAttributeValue("t", "n");
                            cell.AddFirst(new XElement(S + "v", edit.Value));
                        }
                        else
                        {
                            cell.SetAttributeValue("t", "inlineStr");
                            cell.AddFirst(new XElement(S + "is", new XElement(S + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), edit.Value)));
                        }
                    }
                    HighlightChanges(package, edits.Keys);
                    foreach (var sheet in sheets.Where(s => s.Value.Descendants(S + "c").Any(edits.ContainsKey)))
                        using (Stream output = sheet.Key.GetStream(FileMode.Create, FileAccess.Write)) sheet.Value.Save(output);
                    // Ask Excel to refresh formulas that depend on the changed values.
                    if (edits.Count > 0)
                    {
                        var workbookPart = package.GetPart(new Uri("/xl/workbook.xml", UriKind.Relative));
                        XDocument workbook = Read(workbookPart);
                        XElement calc = workbook.Root.Element(S + "calcPr");
                        if (calc != null)
                        {
                            calc.SetAttributeValue("fullCalcOnLoad", "1");
                            calc.SetAttributeValue("forceFullCalc", "1");
                            using (Stream output = workbookPart.GetStream(FileMode.Create, FileAccess.Write)) workbook.Save(output);
                        }
                    }
                }
                return buffer.ToArray();
            }
        }
    }
}
