// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

/// <summary>
/// Keeps a chart's embedded data workbook (ppt|word/embeddings/*.xlsx — the
/// package PowerPoint's "Edit Data" opens) in step with the numbers the chart
/// itself caches.
///
/// A series that came from PowerPoint stores its data as
/// &lt;c:val&gt;&lt;c:numRef&gt;&lt;c:f&gt;Sheet1!$B$2:$B$4&lt;/c:f&gt;&lt;c:numCache&gt;…
/// — the cache paints the slide, the formula names the workbook cells it was
/// built from. Rewriting only the cache (what a literal `values=` set used to
/// do) leaves the two disagreeing: the slide shows the new numbers while
/// "Edit Data" still shows the old ones, and any edit made there regenerates
/// the chart from the stale range, silently reverting the change.
///
/// So when a series still carries its formula *and* the chart carries an
/// embedded workbook, a literal `values=` set is applied to the workbook cells
/// the formula resolves to, the reference is preserved, and only the cache is
/// refreshed. When either is missing (a chart built by `add --type chart` has
/// no workbook at all) the caller falls back to its previous literal-only
/// behaviour.
/// </summary>
internal static class ChartEmbeddedWorkbook
{
    // Sheet1!$B$2:$B$4 · 'My Sheet'!B2:B4 · Sheet1!$B$2 (single cell)
    private static readonly Regex RangePattern = new(
        @"^\s*(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^'!]+))!\s*" +
        @"\$?(?<c1>[A-Za-z]{1,3})\$?(?<r1>\d+)" +
        @"(?::\s*\$?(?<c2>[A-Za-z]{1,3})\$?(?<r2>\d+))?\s*$",
        RegexOptions.Compiled);

    /// <summary>
    /// Mirror <paramref name="values"/> into the chart's embedded workbook and
    /// refresh the series' cache in place, keeping the &lt;c:numRef&gt;.
    /// Returns false — having touched nothing — when this series does not
    /// qualify (no formula, formula is not a single-column cell range, the
    /// range length differs from the value count, or no embedded workbook), so
    /// the caller can apply its literal rewrite exactly as before.
    /// </summary>
    internal static bool TryWriteSeriesValues(ChartPart chartPart, C.Values valEl, double[] values)
    {
        var numRef = valEl.GetFirstChild<C.NumberReference>();
        var formula = numRef?.GetFirstChild<C.Formula>()?.Text;
        if (numRef == null || string.IsNullOrWhiteSpace(formula)) return false;

        var m = RangePattern.Match(formula!);
        if (!m.Success) return false;

        var sheetName = m.Groups["quoted"].Success
            ? m.Groups["quoted"].Value.Replace("''", "'")
            : m.Groups["plain"].Value.Trim();
        var col1 = ColumnToIndex(m.Groups["c1"].Value);
        var row1 = int.Parse(m.Groups["r1"].Value, CultureInfo.InvariantCulture);
        var col2 = m.Groups["c2"].Success ? ColumnToIndex(m.Groups["c2"].Value) : col1;
        var row2 = m.Groups["r2"].Success
            ? int.Parse(m.Groups["r2"].Value, CultureInfo.InvariantCulture)
            : row1;
        if (col1 < 1 || col2 != col1 || row2 < row1) return false;
        if (row2 - row1 + 1 != values.Length) return false;

        var embedded = FindEmbeddedWorkbook(chartPart);
        if (embedded == null) return false;

        // Best effort: an unreadable workbook must not turn a working `set`
        // into a failure — report false so the caller keeps its literal path.
        if (!TryWriteCells(embedded, sheetName, col1, row1, values)) return false;

        numRef.RemoveAllChildren<C.NumberingCache>();
        numRef.AppendChild(BuildNumberingCache(values));
        return true;
    }

    internal static C.NumberingCache BuildNumberingCache(double[] values)
    {
        var cache = new C.NumberingCache(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Length });
        for (int i = 0; i < values.Length; i++)
        {
            cache.AppendChild(new C.NumericPoint(
                new C.NumericValue(values[i].ToString("G", CultureInfo.InvariantCulture)))
            { Index = (uint)i });
        }
        return cache;
    }

    private static EmbeddedPackagePart? FindEmbeddedWorkbook(ChartPart chartPart)
    {
        var relId = chartPart.ChartSpace?.GetFirstChild<C.ExternalData>()?.Id?.Value;
        if (string.IsNullOrEmpty(relId)) return null;
        try
        {
            // Only an .xlsx-shaped package can hold the cells; anything else
            // (a legacy .xls, an OLE blob) is left alone.
            return chartPart.GetPartById(relId!) as EmbeddedPackagePart;
        }
        catch (ArgumentOutOfRangeException)
        {
            return null; // c:externalData names a relationship that is gone
        }
    }

    private static bool TryWriteCells(
        EmbeddedPackagePart part, string sheetName, int column, int startRow, double[] values)
    {
        try
        {
            byte[] source;
            using (var read = part.GetStream(FileMode.Open, FileAccess.Read))
            using (var buffer = new MemoryStream())
            {
                read.CopyTo(buffer);
                source = buffer.ToArray();
            }

            using var package = new MemoryStream();
            package.Write(source, 0, source.Length);
            package.Position = 0;

            using (var doc = SpreadsheetDocument.Open(package, true))
            {
                var workbookPart = doc.WorkbookPart;
                if (workbookPart?.Workbook == null) return false;
                var sheet = workbookPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet?.Id?.Value == null) return false;
                if (workbookPart.GetPartById(sheet.Id!.Value) is not WorksheetPart sheetPart) return false;
                var sheetData = sheetPart.Worksheet?.GetFirstChild<SheetData>();
                if (sheetData == null) return false;

                for (int i = 0; i < values.Length; i++)
                    SetNumericCell(sheetData, $"{IndexToColumn(column)}{startRow + i}", values[i]);
                sheetPart.Worksheet!.Save();
            }

            package.Position = 0;
            using var write = part.GetStream(FileMode.Create, FileAccess.Write);
            package.CopyTo(write);
            return true;
        }
        catch (Exception ex) when (ex is IOException
                                   or InvalidDataException
                                   or OpenXmlPackageException
                                   or FileFormatException
                                   or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    /// Write one numeric cell, keeping the cell's style. Rows and cells are
    /// created on demand and spliced in index order — a spreadsheet whose
    /// &lt;x:row&gt;/&lt;x:c&gt; children are out of order is one Excel refuses to open.
    /// </summary>
    private static void SetNumericCell(SheetData sheetData, string cellRef, double value)
    {
        var rowIndex = (uint)RowOf(cellRef);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row == null)
        {
            row = new Row { RowIndex = rowIndex };
            var preceding = sheetData.Elements<Row>()
                .LastOrDefault(r => (r.RowIndex?.Value ?? 0) < rowIndex);
            if (preceding != null) preceding.InsertAfterSelf(row);
            else sheetData.PrependChild(row);
        }

        var column = ColumnOf(cellRef);
        var cell = row.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
        if (cell == null)
        {
            cell = new Cell { CellReference = cellRef };
            var preceding = row.Elements<Cell>()
                .LastOrDefault(c => ColumnOf(c.CellReference?.Value ?? "A1") < column);
            if (preceding != null) preceding.InsertAfterSelf(cell);
            else row.PrependChild(cell);
        }

        // A pre-existing shared-string (t="s") or inline-string cell must lose
        // its type, or the number would be read back as a string-table index.
        cell.DataType = null;
        cell.RemoveAllChildren<InlineString>();
        cell.CellValue = new CellValue(value.ToString("G", CultureInfo.InvariantCulture));
    }

    private static int ColumnToIndex(string letters)
    {
        var index = 0;
        foreach (var ch in letters.ToUpperInvariant())
        {
            if (ch < 'A' || ch > 'Z') return -1;
            index = index * 26 + (ch - 'A' + 1);
        }
        return index;
    }

    private static string IndexToColumn(int index)
    {
        var name = string.Empty;
        while (index > 0)
        {
            var rem = (index - 1) % 26;
            name = (char)('A' + rem) + name;
            index = (index - 1) / 26;
        }
        return name;
    }

    private static int ColumnOf(string cellRef)
    {
        var letters = new string(cellRef.TakeWhile(char.IsLetter).ToArray());
        return letters.Length == 0 ? 1 : ColumnToIndex(letters);
    }

    private static int RowOf(string cellRef)
    {
        var digits = new string(cellRef.SkipWhile(char.IsLetter).TakeWhile(char.IsDigit).ToArray());
        return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row) ? row : 1;
    }
}
