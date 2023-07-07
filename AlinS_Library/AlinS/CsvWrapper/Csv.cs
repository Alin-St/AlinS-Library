using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;

namespace AlinS.CsvWrapper;

public static class Csv
{
    /// <summary>
    /// Parse a CSV-formatted string and get each field into a string. The header is included in the result.
    /// </summary>
    /// <param name="newLine"> The new-line separator. If null, Enviorment.NewLine is used. </param>
    public static List<string[]> Parse(string csv, string separator = ",", string? newLine = null)
    {
        // Open the CsvParser.
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = separator,
            HasHeaderRecord = false,
            NewLine = newLine ?? Environment.NewLine,
            BadDataFound = null,
        };

        using var stringReader = new StringReader(csv);
        using var csvParser = new CsvParser(stringReader, config);

        // Add each row to a list.
        var result = new List<string[]>();

        while (csvParser.Read())
            result.Add(csvParser.Record ?? Array.Empty<string>());

        return result;
    }

    /// <summary>
    /// Serialize a table of strings into a csv.
    /// </summary>
    /// <returns> The CSV-formatted string. </returns>
    /// <param name="newLine"> The new-line separator. If null, Enviorment.NewLine is used. </param>
    public static string Serialize(IEnumerable<IEnumerable<string>> values, string separator = ",", string? newLine = null)
    {
        // Open the CsvWriter
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = separator,
            HasHeaderRecord = false,
            NewLine = newLine ?? Environment.NewLine,
            BadDataFound = null,
        };

        using var stringWriter = new StringWriter();
        using var csvWriter = new CsvWriter(stringWriter, config);

        // Write each field to the string.
        foreach (var row in values)
        {
            foreach (var cell in row)
                csvWriter.WriteField(cell);

            csvWriter.NextRecord();
        }

        csvWriter.Flush();
        return stringWriter.ToString();
    }

    /// <summary>
    /// Parse a TSV-formatted string (tab-separated values) from an Excel selection.
    /// </summary>
    /// <returns> A list of lists representing the cells in text form. </returns>
    public static List<List<string>> FromExcel(string tsv)
    {
        return Parse(tsv, "\t").Select(r => r.ToList()).ToList();
    }

    /// <summary>
    /// Serialize a table of objects into a tsv (tab-separated values) that can be pasted into Excel. <br/>
    /// Objects are converted to strings. If null they are converted to empty strings.
    /// </summary>
    /// <returns> The TSV-formatted string. </returns>
    public static string ToExcel(IEnumerable<IEnumerable<object?>> table)
    {
        return Serialize(table.Select(r => r.Select(c => c?.ToString() ?? "")), "\t");
    }
}
