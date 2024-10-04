using ReportLockerAPI;
using System;
using System.IO;

namespace ReportLockerApp
{
    internal class Program
    {
        static void PrintUsage()
        {
            Console.WriteLine($"Usage {System.Reflection.Assembly.GetExecutingAssembly().GetName().Name} <--help|--lock|--unlock|--check|--sign> <file|folder|topic> [key|signature] [sheet] [row] [column]");
        }
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                PrintUsage();
                return;
            }

            string @switch = args[0].ToLower().TrimStart('-');

            if (args.Length == 2 && @switch == "help")
            {
                switch (args[1].ToLower())
                {
                    case "lock":
                        Console.WriteLine("Lock xlsx report/reports with key provided");
                        break;
                    case "unlock":
                        Console.WriteLine("Unlock xlsx report/reports");
                        break;
                    case "check":
                        Console.WriteLine("Check protection on report/reports");
                        break;
                    case "sign":
                        Console.WriteLine("Sign xlsx report/reports with signature at specified sheet, row, column coordinate");
                        break;
                    default:
                        Console.WriteLine("Unknown topic");
                        break;
                }

                return;
            }

            uint row = 0;

            switch (@switch)
            {
                case "lock" when args.Length < 3:
                    Console.WriteLine("Use of --lock require a key value");
                    return;
                case "sign" when args.Length < 6:
                    Console.WriteLine("Use of --sign require signature, sheet, row and column values");
                    return;
                case "sign" when !uint.TryParse(args[4], out row):
                    Console.WriteLine("Wrong format for row parameter");
                    return;
            }

            var dir = new DirectoryInfo(args[1]);

            if (!dir.Exists && !File.Exists(args[1]))
            {
                Console.WriteLine($"File {args[1]} must exist");
                return;
            }

            string[] files;

            if (dir.Exists)
                files = Directory.GetFiles(args[1], "*.xlsx");
            else
            {
                string extension = Path.GetExtension(args[1]).ToLower();

                if (string.IsNullOrEmpty(extension))
                {
                    Console.WriteLine("Cannot process files without extension");
                    return;
                }

                if (!extension.Equals(".xlsx"))
                {
                    Console.WriteLine($"{extension} files are not allowed (only xlsx)");
                    return;
                }

                files = new string[1];
                files[0] = args[1];
            }

            var count = 0;

            var reportLocker = new ReportLocker();

            foreach (string file in files)
            {
                switch (@switch)
                {
                    case "lock":
                        if (reportLocker.Lock(file, args[2]))
                        {
                            count++;
                            Console.WriteLine($"Report {file} locked");
                        }
                        else
                            Console.WriteLine($"Report {file} lock failed");
                        break;

                    case "unlock":
                        if (reportLocker.Unlock(file))
                        {
                            count++;
                            Console.WriteLine($"Report {file} unlocked");
                        }
                        else
                            Console.WriteLine($"Report {file} unlock failed");
                        break;

                    case "check":
                        ReportLocker.Protection result = reportLocker.GetProtection(file);

                        Console.WriteLine($"Report {file} is {result}");
                        break;

                    case "sign":
                        if (reportLocker.SignReport(file, args[2], args[3], row, args[5]))
                            Console.WriteLine($"Report {file} signed");
                        else
                            Console.WriteLine($"Report {file} not signed");
                        break;

                    //TODO: check signature

                    default:
                        Console.WriteLine($"Unknown switch: {@switch}");
                        break;
                }
            }

            Console.WriteLine($"{files.Length} files, {count} processed");
        }
    }
}
