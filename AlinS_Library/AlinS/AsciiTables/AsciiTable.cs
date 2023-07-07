namespace AlinS.AsciiTables;

public class AsciiTable
{
    /// <summary>
    /// Converts each object into a string, then the string matrix is normalized and each value is a cell in the ASCII table.
    /// New-line and tab characters are properly handled. Zero-width characters and other special Unicode characters are NOT.
    /// The characters '-', '|' and '+' are used for the borders, with '=' instead of '-' for the header.
    /// The result will have at least one row and one column.
    /// </summary>
    /// <param name="hasHeader"> If true, the first line will be separated with '=' instead of '-'. </param>
    public static string Build(IEnumerable<IEnumerable<object?>> table, bool hasHeader = false)
    {
        // Get a matrix where each cell is a string[], representing the cell content split by new-lines.
        var strMatrix = table.Select(
            row => row.Select(
                cell => cell?.ToString()?.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None) ?? Array.Empty<string>()
            ).ToArray()
        ).ToArray();

        // Replace each "\t" with "    " for accurate allignment.
        foreach (var row in strMatrix)
            foreach (var cell in row)
                for (int i = 0; i < cell.Length; i++)
                    cell[i] = cell[i].Replace("\t", "    ");

        // Every table must have at least one row and one column.
        if (strMatrix.Length == 0)
            strMatrix = new[] { new[] { new[] { "" } } };
        if (strMatrix[0].Length == 0)
            strMatrix[0] = new[] { new[] { "" } };

        // Find the number of columns;
        int columnCount = strMatrix.Max(row => row.Length);

        // Normalize the matrix.
        for (int i = 0; i < strMatrix.Length; i++)
        {
            var r = strMatrix[i];
            if (r.Length < columnCount)
                strMatrix[i] = r.Concat(Enumerable.Range(0, columnCount - r.Length).Select(_ => Array.Empty<string>())).ToArray();
        }

        // Normalize the number of lines for each row.
        foreach (var row in strMatrix)
        {
            int lineCount = Math.Max(1, row.Max(cell => cell.Length));
            for (int i = 0; i < row.Length; i++)
            {
                var c = row[i];
                if (c.Length < lineCount)
                    row[i] = c.Concat(Enumerable.Range(0, lineCount - c.Length).Select(_ => "")).ToArray();
            }
        }

        // Normalize the line length for each column.
        for (int i = 0; i < columnCount; ++i)
        {
            int lineLength = strMatrix.SelectMany(r => r[i]).Max(l => l.Length);
            foreach (var row in strMatrix)
                for (int j = 0; j < row[i].Length; ++j)
                    row[i][j] = row[i][j].PadRight(lineLength);
        }

        // Get the result.
        var separatorLine = "+" + string.Join("+", Enumerable.Range(0, columnCount)
            .Select(i => new string('-', strMatrix[0][i][0].Length))) + "+";

        var nl = Environment.NewLine;
        var lineStrings = strMatrix.Select(r => r[0].Select((_, i) =>
            "|" + string.Join("|", r.Select(c => c[i])) + "|").ToArray()).ToArray();
        var rowStrings = lineStrings.Select(r => string.Join(nl, r)).ToArray();

        string result;

        if (hasHeader)
        {
            var headerSepLine = separatorLine.Replace('-', '=');
            result = headerSepLine + nl + rowStrings[0] + nl + headerSepLine;

            if (rowStrings.Length >= 2)
                result += nl + string.Join(nl + separatorLine + nl, rowStrings.Skip(1)) + nl + separatorLine;
        }
        else
        {
            result = separatorLine + nl + string.Join(nl + separatorLine + nl, rowStrings) + nl + separatorLine;
        }

        return result;
    }
}
