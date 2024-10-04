using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace ReportLockerAPI
{
    public class ReportLocker
    {
        public enum Protection {
            Unlocked = 0,
            Locked = 1,
            Crypted = 2
        }
        public ReportLocker()
        {
        }
        protected string HexPasswordConversion(string password)
        {
            byte[] passwordCharacters = System.Text.Encoding.ASCII.GetBytes(password);

            int hash = 0;

            if (passwordCharacters.Length > 0)
            {
                int charIndex = passwordCharacters.Length;

                while (charIndex-- > 0)
                {
                    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                    hash ^= passwordCharacters[charIndex];
                }

                // Main difference from spec, also hash with charcount
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters.Length;
                hash ^= (0x8000 | ('N' << 8) | 'K');
            }

            return Convert.ToString(hash, 16).ToUpperInvariant();
        }
        public bool Lock(string report, string key)
        {
            if (string.IsNullOrEmpty(report))
                throw new ArgumentException("Report file required");

            if (!System.IO.File.Exists(report))
                throw new ArgumentException("Report file must exist");

            string hexKey = HexPasswordConversion(key);

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(report, true))
                {
                    foreach (var part in document.WorkbookPart.WorksheetParts)
                    {
                        SheetProtection protection = new SheetProtection()
                        {
                            Sheet = true,
                            Objects = true,
                            Scenarios = true,
                            Password = hexKey
                        };

                        part.Worksheet.InsertAfter(protection, part.Worksheet.Descendants<SheetData>().LastOrDefault());
                        part.Worksheet.Save();
                    }
                }
            }
            catch (System.IO.FileFormatException)
            {
                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return true;
        }
        public bool Unlock(string report)
        {
            if (string.IsNullOrEmpty(report))
                throw new ArgumentException("Report file required");

            if (!System.IO.File.Exists(report))
                throw new ArgumentException("Report file must exist");

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(report, true))
                {
                    foreach (var part in document.WorkbookPart.WorksheetParts)
                    {
                        part.Worksheet.RemoveAllChildren<SheetProtection>();
                        part.Worksheet.Save();
                    }
                }
            }
            catch (System.IO.FileFormatException)
            {
                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return true;
        }
        public Protection GetProtection(string report)
        {
            if (string.IsNullOrEmpty(report))
                throw new ArgumentException("Report file required");

            if (!System.IO.File.Exists(report))
                throw new ArgumentException("Report file must exist");

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(report, false))
                {
                    foreach (var part in document.WorkbookPart.WorksheetParts)
                    {
                        var protections = part.Worksheet.Elements<SheetProtection>();

                        if (protections.Any())
                            return Protection.Locked;
                    }
                }
            }
            catch (System.IO.FileFormatException) {
                return Protection.Crypted;
            }

            return Protection.Unlocked;
        }
        public string GetProtectionString(string report)
        {
            return GetProtection(report).ToString();
        }
        public bool SignReport(string report, string signature, int row, int column)
        {
            if (string.IsNullOrEmpty(report))
                throw new ArgumentException("Report file required");

            if (!System.IO.File.Exists(report))
                throw new ArgumentException("Report file must exist");

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(report, true))
                {
                    //TODO: place signature at row,column
                }
            }
            catch (System.IO.FileFormatException)
            {
                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return true;
        }
        public bool IsSigned(string report, string signature, int row, int column)
        {
            if (string.IsNullOrEmpty(report))
                throw new ArgumentException("Report file required");

            if (!System.IO.File.Exists(report))
                throw new ArgumentException("Report file must exist");

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(report, false))
                {
                    //TODO get specific cell value
                }
            }
            catch (System.IO.FileFormatException)
            {
                return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return false;
        }
        public static ReportLocker Create() { return new ReportLocker(); }
    }
}
