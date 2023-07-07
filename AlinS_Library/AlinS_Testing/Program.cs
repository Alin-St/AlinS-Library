using AlinS.AsciiTables;
using AlinS.CsvWrapper;
using AlinS.ExcelHelper;

namespace AlinS_Testing;

internal class Program
{
    static void Main()
    {
        Directory.SetCurrentDirectory("../../../_Run");

        var x = Excel.Read("0_test.xlsx");
        Console.WriteLine(AsciiTable.Build(x));

        x = x.Select(r => r.Select(c => c.ToUpper()).ToList()).ToList();
        Excel.Save("1_test.xlsx", x);

        x = x.Select(r => r.Select(c => c + " X").ToList()).ToList();
        Excel.Save("1_test.xlsx", x, "Sheet2");
    }
}
