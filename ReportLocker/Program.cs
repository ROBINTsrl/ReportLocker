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
                Console.WriteLine("Usage ReportLocker <--lock|--unlock|--check> <file|folder> [key]");
                return;
            }

            if (args[0] == "--lock" && args.Length < 3)
            {
                Console.WriteLine("Use of --lock require a key value");
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
                    if(reportLocker.Lock(args[1], args[2]))
                        Console.WriteLine("Report {0} locked", args[1]);
                    else
                        Console.WriteLine("Report {0} lock failed", args[1]);
                }
                else if (args[0] == "--unlock")
                {
                    if(reportLocker.Unlock(args[1]))
                        Console.WriteLine("Report {0} unlocked", args[1]);
                    else
                        Console.WriteLine("Report {0} unlock failed", args[1]);
                }
                else if (args[0] == "--check")
                {
                    ReportLocker.Protection result = reportLocker.GetProtection(args[1]);

                    Console.WriteLine("Report {0} is {1}", args[1], result.ToString());
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
