// Dumps a spreadsheet to a canonical text form using the original C#
// ExcelDataReader, so the Go port can be diffed against it.
//
// The format is deliberately unambiguous. Numbers are emitted as the raw IEEE
// 754 bit pattern rather than as text, because C# and Go disagree about how to
// render a double as a string and that disagreement would show up as a false
// difference on every non-integral value.

using System.Globalization;
using System.Text;
using ExcelDataReader;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

if (args.Length < 1)
{
    Console.Error.WriteLine("usage: edrdump <file> [...]");
    return 1;
}

foreach (var path in args)
{
    using var stream = File.Open(path, FileMode.Open, FileAccess.Read);
    using var reader = ExcelReaderFactory.CreateReader(stream);

    int sheetIndex = 0;
    do
    {
        string sheetName = reader.Name ?? "";
        int rowIndex = 0;
        while (reader.Read())
        {
            for (int col = 0; col < reader.FieldCount; col++)
            {
                object? v = reader.GetValue(col);
                if (v == null) continue;

                string tag, text;
                switch (v)
                {
                    case string s:
                        tag = "S";
                        text = Escape(s);
                        break;
                    case bool b:
                        tag = "B";
                        text = b ? "true" : "false";
                        break;
                    case DateTime d:
                        tag = "D";
                        text = d.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
                        break;
                    case double dbl:
                        tag = "N";
                        text = BitConverter.DoubleToInt64Bits(dbl).ToString("X16");
                        break;
                    case int i:
                        tag = "N";
                        text = BitConverter.DoubleToInt64Bits((double)i).ToString("X16");
                        break;
                    case decimal m:
                        tag = "N";
                        text = BitConverter.DoubleToInt64Bits((double)m).ToString("X16");
                        break;
                    default:
                        tag = "S";
                        text = Escape(Convert.ToString(v, CultureInfo.InvariantCulture) ?? "");
                        break;
                }

                Console.WriteLine(string.Join("\t",
                    sheetIndex.ToString(CultureInfo.InvariantCulture),
                    Escape(sheetName),
                    rowIndex.ToString(CultureInfo.InvariantCulture),
                    col.ToString(CultureInfo.InvariantCulture),
                    tag,
                    text));
            }
            rowIndex++;
        }
        sheetIndex++;
    } while (reader.NextResult());
}

return 0;

static string Escape(string s) =>
    s.Replace("\\", "\\\\").Replace("\t", "\\t").Replace("\r", "\\r").Replace("\n", "\\n");
