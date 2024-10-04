using ReportLockerAPI;
using System;
using System.IO;

namespace ReportLockerApp
{
    internal class Program
    {
        static void PrintUsage()
        {
            Console.WriteLine("Usage ReportLocker <--help|--lock|--unlock|--check|--sign> <file|folder|topic> [key|signature] [sheet] [row] [column]");
        }
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                PrintUsage();
                return;
            }

            if (args.Length == 2 && args[0] == "--help")
            {
                switch (args[1])
                {
                    case "lock":
                        Console.WriteLine("Lock xlsx report/reports with key provided");
                        break;
                    case "unlock":
                        Console.WriteLine("Unlock xlsx report/reports");
                        break;
                    case "check":
                        Console.WriteLine("Unlock xlsx report/reports");
                        break;
                    case "sign":
                        Console.WriteLine("Sign xlsx report/reports with signature at specified sheet, row, column coordinate");
                        break;
                    default:
                        Console.WriteLine("Unknown topic");
                        break;
                }
            }

            uint row = 0;

            switch (args[0])
            {
                case "--lock" when args.Length < 3:
                    Console.WriteLine("Use of --lock require a key value");
                    return;
                case "--sign" when args.Length < 6:
                    Console.WriteLine("Use of --sign require signature, sheet, row and column values");
                    return;
                case "--sign" when !uint.TryParse(args[4], out row):
                    Console.WriteLine("Wrong format for row parameter");
                    return;
            }

            string extension = Path.GetExtension(args[1]).ToLower();

            var dir = new DirectoryInfo(args[1]);

            bool folder = dir.Exists;

            if (!folder && !extension.Equals(".xlsx"))
            {
                Console.WriteLine("{0} files are not allowed (only xlsx)", extension);
                return;
            }

            if (!folder && !File.Exists(args[1]))
            {
                Console.WriteLine("File {0} must exist", args[1]);
                return;
            }

            string[] files;

            if (folder)
                files = Directory.GetFiles(args[1], "*.xlsx");
            else
            {
                files = new string[1];
                files[0] = args[1];
            }

            var count = 0;

            var reportLocker = new ReportLocker();

            foreach (string file in files)
            {
                switch (args[0])
                {
                    case "--lock":
                        if (reportLocker.Lock(file, args[2]))
                        {
                            count++;
                            Console.WriteLine("Report {0} locked", file);
                        }
                        else
                            Console.WriteLine("Report {0} lock failed", file);
                        break;

                    case "--unlock":
                        if (reportLocker.Unlock(file))
                        {
                            count++;
                            Console.WriteLine("Report {0} unlocked", file);
                        }
                        else
                            Console.WriteLine("Report {0} unlock failed", file);
                        break;

                    case "--check":
                        ReportLocker.Protection result = reportLocker.GetProtection(file);

                        Console.WriteLine("Report {0} is {1}", file, result.ToString());
                        break;

                    case "--sign":
                        if (reportLocker.SignReport(file, args[2], args[3], row, args[5]))
                            Console.WriteLine("Report {0} signed", file);
                        else
                            Console.WriteLine("Report {0} not signed", file);
                        break;
                }
            }

            Console.WriteLine("{0} files {1} processed ", files.Length, count);
        }
    }
}
