using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelReader
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("Excel To CSV\n");
                Console.WriteLine("ExcelToCsv.exe <filename> [/output:<filename>] [/sheet:<number>]");
                return 0;
            }

            var inFile = args[0];

            if (!File.Exists(inFile))
            {
                Console.WriteLine($"File doesn't exist: {inFile}");
                return -1;
            }

            var outFile = Path.ChangeExtension(inFile, ".csv");
            var sheet = 0;

            if (args.Length > 1)
                for (int i = 1; i < args.Length; i++)
                {
                    if (args[i].ToLower().StartsWith("/sheet:"))
                    {
                        Int32.TryParse(args[1].Substring("/sheet:".Length), out sheet);
                    }
                    if (args[i].ToLower().StartsWith("/output:"))
                    {
                        outFile = args[1].Substring("/output:".Length);
                    }
                }

            var sb = new StringBuilder();
            using (var stream = File.Open(inFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    foreach (DataRow row in result.Tables[sheet].Rows)
                    {
                        foreach (var item in row.ItemArray)
                            sb.Append(item.ToString() + "\t");
                        sb.Replace("\t", Environment.NewLine, sb.Length - 1, 1);
                    }
                }
            }

            try
            {
                File.WriteAllText(outFile, sb.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return -2;
            }
            Console.WriteLine($"File {outFile} generated.");

            return 1;
        }
    }
}
