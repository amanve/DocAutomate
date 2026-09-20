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

        private static bool HasSoftwareChangesTitle(XDocument document)
        {
            return document.Descendants(P + "sp").Any(shape =>
                string.Equals(Text(shape, A).Trim(), "Software Changes", StringComparison.OrdinalIgnoreCase));
        }

        private static PackagePart RelatedPart(Package package, PackagePart part, string relationshipName)
        {
            const string prefix = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
            PackageRelationship relationship = part.GetRelationshipsByType(prefix + relationshipName)
                .FirstOrDefault(r => r.TargetMode == TargetMode.Internal);
            return relationship == null ? null : package.GetPart(
                PackUriHelper.ResolvePartUri(part.Uri, relationship.TargetUri));
        }

        private static bool IsSoftwareChangesSlide(Package package, PackagePart part, XDocument slide)
        {
            if (HasSoftwareChangesTitle(slide)) return true;

            // Text shown in Master View belongs to the slide's layout or its master.
            PackagePart layout = RelatedPart(package, part, "slideLayout");
            if (layout == null) return false;
            if (HasSoftwareChangesTitle(Read(layout))) return true;

            PackagePart master = RelatedPart(package, layout, "slideMaster");
            return master != null && HasSoftwareChangesTitle(Read(master));
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
                    if (!IsSoftwareChangesSlide(package, part, slide)) continue;
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
                throw new InvalidOperationException("No data table found on a Software Changes slide.");
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

        private static void SetText(XElement cell, string value)
        {
            cell.Elements(S + "f").Remove();
            cell.Elements(S + "v").Remove();
            cell.Elements(S + "is").Remove();
            cell.SetAttributeValue("t", "inlineStr");
            cell.AddFirst(new XElement(S + "is", new XElement(S + "t",
                new XAttribute(XNamespace.Xml + "space", "preserve"), value)));
        }

        private static bool IsValueHeader(string text)
        {
            text = text.Trim();
            return text.StartsWith("Fill Value from ", StringComparison.OrdinalIgnoreCase) ||
                (text.StartsWith("SAA", StringComparison.Ordinal) &&
                 text.IndexOf(" (0x", StringComparison.Ordinal) > 3 &&
                 text.EndsWith(")", StringComparison.Ordinal));
        }

        private static bool FillDocumentInputs(XDocument sheet, List<string> shared, string part, string crc, string model)
        {
            if (part == null || crc == null) return false;
            var cells = sheet.Descendants(S + "c").ToList();
            bool updated = false;
            foreach (XElement header in cells.Where(c =>
                IsValueHeader(CellText(c, shared))))
            {
                int valueColumn = Column((string)header.Attribute("r"));
                int headerRow = Row((string)header.Attribute("r"));
                foreach (XElement label in cells.Where(c =>
                    Column((string)c.Attribute("r")) == valueColumn - 1 &&
                    Row((string)c.Attribute("r")) > headerRow))
                {
                    string text = CellText(label, shared).Trim();
                    string value = text.Equals("Part", StringComparison.OrdinalIgnoreCase) ? part :
                        text.Equals("CRC", StringComparison.OrdinalIgnoreCase) ? crc : null;
                    if (value == null) continue;
                    string address = Address(valueColumn, Row((string)label.Attribute("r")));
                    XElement target = cells.FirstOrDefault(c => (string)c.Attribute("r") == address);
                    if (target == null)
                    {
                        target = new XElement(S + "c", new XAttribute("r", address));
                        XElement next = label.Parent.Elements(S + "c").FirstOrDefault(c =>
                            Column((string)c.Attribute("r")) > valueColumn);
                        if (next == null) label.Parent.Add(target);
                        else next.AddBeforeSelf(target);
                    }
                    SetText(target, value);
                }
                if (model != null && headerRow > 1)
                {
                    string modelAddress = Address(valueColumn, headerRow - 1);
                    XElement modelCell = cells.FirstOrDefault(c => (string)c.Attribute("r") == modelAddress);
                    if (modelCell == null)
                    {
                        XElement data = sheet.Root.Element(S + "sheetData");
                        XElement modelRow = data.Elements(S + "row").FirstOrDefault(r => (int)r.Attribute("r") == headerRow - 1);
                        if (modelRow == null)
                        {
                            modelRow = new XElement(S + "row", new XAttribute("r", headerRow - 1));
                            XElement nextRow = data.Elements(S + "row").FirstOrDefault(r => (int)r.Attribute("r") > headerRow - 1);
                            if (nextRow == null) data.Add(modelRow);
                            else nextRow.AddBeforeSelf(modelRow);
                        }
                        modelCell = new XElement(S + "c", new XAttribute("r", modelAddress));
                        XElement nextCell = modelRow.Elements(S + "c").FirstOrDefault(c => Column((string)c.Attribute("r")) > valueColumn);
                        if (nextCell == null) modelRow.Add(modelCell);
                        else nextCell.AddBeforeSelf(modelCell);
                        // Let Excel recompute the used range after adding the model cell.
                        sheet.Root.Elements(S + "dimension").Remove();
                    }
                    SetText(modelCell, model);
                }
                SetText(header, part + " (" + crc + ")");
                updated = true;
            }
            return updated;
        }

        private static bool ClearCheckResults(XDocument sheet, List<string> shared)
        {
            var cells = sheet.Descendants(S + "c").ToList();
            var headers = cells.Where(c =>
                new[] { "Check1", "Check2", "Check3" }.Contains(CellText(c, shared).Trim()))
                .Select(c => new { Column = Column((string)c.Attribute("r")), Row = Row((string)c.Attribute("r")) })
                .ToList();
            bool cleared = false;
            foreach (XElement cell in cells)
            {
                string address = (string)cell.Attribute("r");
                if (CellText(cell, shared) != "OK" ||
                    !headers.Any(h => h.Column == Column(address) && h.Row < Row(address))) continue;
                cell.Elements(S + "v").Remove();
                cell.Elements(S + "is").Remove();
                cell.Elements(S + "f").Remove();
                cell.SetAttributeValue("t", null);
                cleared = true;
            }
            return cleared;
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

        private static bool IsOldHighlight(XElement color, List<string> theme, List<string> indexed, bool font)
        {
            if (color == null) return false;
            string rgb = (string)color.Attribute("rgb");
            int index;
            if (rgb == null && int.TryParse((string)color.Attribute("theme"), out index) &&
                index >= 0 && index < theme.Count) rgb = theme[index];
            if (rgb == null && int.TryParse((string)color.Attribute("indexed"), out index) &&
                index >= 0 && index < indexed.Count) rgb = indexed[index];
            uint value;
            if (rgb == null || !uint.TryParse(rgb, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value))
                return false;
            double r = (value >> 16) & 255, g = (value >> 8) & 255, b = value & 255;
            double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
            double delta = max - min;
            if (delta < 1 || delta / max < 0.15) return false;
            double hue = max == r ? 60 * ((g - b) / delta) :
                max == g ? 60 * (2 + (b - r) / delta) : 60 * (4 + (r - g) / delta);
            if (hue < 0) hue += 360;
            // Include light and dark shades of red, orange and yellow.
            return font ? hue <= 15 || hue >= 345 : hue >= 20 && hue <= 65;
        }

        private static void ClearOldHighlights(Package package)
        {
            var theme = new List<string>();
            PackagePart themePart = package.GetParts().FirstOrDefault(p =>
                p.ContentType == "application/vnd.openxmlformats-officedocument.theme+xml");
            if (themePart != null)
            {
                XElement scheme = Read(themePart).Descendants(A + "clrScheme").FirstOrDefault();
                if (scheme != null)
                {
                    // Spreadsheet theme indexes use light before dark.
                    foreach (string name in new[] { "lt1", "dk1", "lt2", "dk2", "accent1", "accent2",
                        "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" })
                    {
                        XElement entry = scheme.Element(A + name);
                        XElement color = entry == null ? null : entry.Elements().FirstOrDefault();
                        theme.Add(color == null ? null : (string)color.Attribute("lastClr") ?? (string)color.Attribute("val"));
                    }
                }
            }
            var indexed = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
                "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF " +
                "800000 008000 000080 808000 800080 008080 C0C0C0 808080 " +
                "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF " +
                "000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF " +
                "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 " +
                "3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 " +
                "003366 339966 003300 333300 993300 993366 333399 333333").Split(' ').ToList();
            PackagePart stylesPart = package.GetParts().FirstOrDefault(p =>
                p.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            if (stylesPart != null)
            {
                XDocument styles = Read(stylesPart);
                XElement custom = styles.Descendants(S + "indexedColors").FirstOrDefault();
                if (custom != null) indexed = custom.Elements(S + "rgbColor").Select(c => (string)c.Attribute("rgb")).ToList();
            }

            foreach (PackagePart part in package.GetParts().Where(p =>
                p.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" ||
                p.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" ||
                p.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml").ToList())
            {
                XDocument document = Read(part);
                bool changed = false;
                foreach (XElement fill in document.Descendants(S + "patternFill"))
                {
                    if (!IsOldHighlight(fill.Element(S + "fgColor"), theme, indexed, false) &&
                        !IsOldHighlight(fill.Element(S + "bgColor"), theme, indexed, false)) continue;
                    fill.RemoveNodes();
                    fill.SetAttributeValue("patternType", "none");
                    changed = true;
                }
                foreach (XElement color in document.Descendants(S + "color").Where(c =>
                    c.Parent.Name == S + "font" || c.Parent.Name == S + "rPr").ToList())
                {
                    if (!IsOldHighlight(color, theme, indexed, true)) continue;
                    color.ReplaceWith(new XElement(S + "color", new XAttribute("auto", "1")));
                    changed = true;
                }
                if (changed)
                    using (Stream output = part.GetStream(FileMode.Create, FileAccess.Write)) document.Save(output);
            }
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

        public byte[] Apply(string excelPath, string powerpointPath, string part = null, string crc = null, string model = null)
        {
            var changes = string.IsNullOrWhiteSpace(powerpointPath) ? new List<Change>() : ReadChanges(powerpointPath);
            using (var buffer = new MemoryStream())
            {
                using (Stream input = File.OpenRead(excelPath)) input.CopyTo(buffer);
                buffer.Position = 0;
                using (Package package = Package.Open(buffer, FileMode.Open, FileAccess.ReadWrite))
                {
                    ClearOldHighlights(package);
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
                            // Step headers span an ID column and a label column.
                            var stepHeaders = cells.Values.Where(c =>
                                CellText(c, shared).Trim().Equals("2Step", StringComparison.OrdinalIgnoreCase)).ToList();
                            foreach (XElement cell in cells.Values.Where(c => CellText(c, shared) == change.Setting))
                            {
                                string address = (string)cell.Attribute("r");
                                int column = Column(address), row = Row(address);
                                int groupColumn = column - 1, valueColumn = column + 1;
                                if (stepHeaders.Count > 0)
                                {
                                    XElement header = stepHeaders.Where(h => Row((string)h.Attribute("r")) < row)
                                        .OrderByDescending(h => Row((string)h.Attribute("r"))).FirstOrDefault();
                                    if (header == null) continue;
                                    int headerRow = Row((string)header.Attribute("r"));
                                    XElement settingHeader = cells.Values.FirstOrDefault(h =>
                                        Row((string)h.Attribute("r")) == headerRow &&
                                        CellText(h, shared).Trim().Equals("4Step", StringComparison.OrdinalIgnoreCase));
                                    XElement valueHeader = cells.Values.FirstOrDefault(h =>
                                        Row((string)h.Attribute("r")) == headerRow &&
                                        IsValueHeader(CellText(h, shared)));
                                    if (settingHeader == null || valueHeader == null)
                                        throw new InvalidOperationException("The step table requires 4Step and a value-column header.");
                                    if (column != Column((string)settingHeader.Attribute("r")) + 1) continue;
                                    groupColumn = Column((string)header.Attribute("r")) + 1;
                                    valueColumn = Column((string)valueHeader.Attribute("r"));
                                }
                                if (groupColumn < 1 ||
                                    CellText(GroupCell(sheet, cells, Address(groupColumn, row)), shared) != change.Group) continue;
                                XElement target;
                                if (cells.TryGetValue(Address(valueColumn, row), out target)) matches.Add(target);
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
                    var clearedSheets = new HashSet<XDocument>(sheets.Values.Where(s => ClearCheckResults(s, shared)));
                    clearedSheets.UnionWith(sheets.Values.Where(s => FillDocumentInputs(s, shared, part, crc, model)));
                    foreach (var sheet in sheets.Where(s => clearedSheets.Contains(s.Value) || s.Value.Descendants(S + "c").Any(edits.ContainsKey)))
                        using (Stream output = sheet.Key.GetStream(FileMode.Create, FileAccess.Write)) sheet.Value.Save(output);
                    // Ask Excel to refresh formulas that depend on the changed values.
                    if (edits.Count > 0 || clearedSheets.Count > 0)
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
