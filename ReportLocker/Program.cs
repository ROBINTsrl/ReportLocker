using ReportLockerAPI;
using System;
using System.IO;

namespace ReportLockerApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage ReportLocker <--lock|--unlock|--check|--sign> <file|folder> [key|signature] [sheet] [row] [column]");
                return;
            }

            if (args[0] == "--lock" && args.Length < 3)
            {
                Console.WriteLine("Use of --lock require a key value");
                return;
            }

            if (args[0] == "--sign" && args.Length < 6)
            {
                Console.WriteLine("Use of --sign require signature, sheet, row and column values");
                return;
            }

            string extension = Path.GetExtension(args[1]).ToLower();

            DirectoryInfo dir = new DirectoryInfo(args[1]);

            bool folder = dir.Exists;

            if (!folder && !extension.Equals(".xlsx"))
            {
                Console.WriteLine("{0} file are not allowed (only xlsx)", extension);
                return;
            }

            if (!folder && !File.Exists(args[1]))
            {
                Console.WriteLine("File {0} must exist", args[1]);
                return;
            }

            ReportLocker reportLocker = new ReportLocker();

            if (!folder)
            {
                if (args[0] == "--lock")
                {
                    if (reportLocker.Lock(args[1], args[2]))
                        Console.WriteLine("Report {0} locked", args[1]);
                    else
                        Console.WriteLine("Report {0} lock failed", args[1]);
                }
                else if (args[0] == "--unlock")
                {
                    if (reportLocker.Unlock(args[1]))
                        Console.WriteLine("Report {0} unlocked", args[1]);
                    else
                        Console.WriteLine("Report {0} unlock failed", args[1]);
                }
                else if (args[0] == "--check")
                {
                    ReportLocker.Protection result = reportLocker.GetProtection(args[1]);

                    Console.WriteLine("Report {0} is {1}", args[1], result.ToString());
                }
                else if (args[0] == "--sign")
                {
                    if (!uint.TryParse(args[4], out uint row))
                        Console.WriteLine("Wrong format for row parameter");
                    else if (reportLocker.SignReport(args[1], args[2], args[3], row, args[5]))
                        Console.WriteLine("Report {0} signed", args[1]);
                    else
                        Console.WriteLine("Report {0} not signed", args[1]);
                }
            }
            else 
            {
                string[] files = Directory.GetFiles(args[1], "*.xlsx");
                int count = 0;

                foreach (string file in files)
                {
                    if (args[0] == "--lock")
                    {
                        if (reportLocker.Lock(file, args[2]))
                        {
                            count++;
                            Console.WriteLine("Report {0} locked", file);
                        }
                        else
                            Console.WriteLine("Report {0} lock failed", file);
                    }
                    else if (args[0] == "--unlock")
                    {
                        if (reportLocker.Unlock(file))
                        {
                            count++;
                            Console.WriteLine("Report {0} unlocked", file);
                        }
                        else
                            Console.WriteLine("Report {0} unlock failed", file);
                    }
                    else if (args[0] == "--check")
                    {
                        ReportLocker.Protection result = reportLocker.GetProtection(file);

                        Console.WriteLine("Report {0} is {1}", file, result.ToString());
                    }
                }

                Console.WriteLine("{0} files {1} processed ", files.Length, count);
            }
        }
    }
}
