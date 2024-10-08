using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace ReportLockerAPI
{
    public class ReportLocker
    {
        public enum Protection
        {
            Unlocked = 0,
            Locked = 1,
            Crypted = 2
        }
        public ReportLocker()
        {
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
                    if (document.WorkbookPart.Workbook.FileSharing != null)
                    {
                        document.WorkbookPart.Workbook.FileSharing.Remove();
                        document.WorkbookPart.Workbook.Save();
                    }
                    
                    foreach (var part in document.WorkbookPart.WorksheetParts)
                    {
                        part.Worksheet.RemoveAllChildren<SheetProtection>();
                        part.Worksheet.Save();
                    }

                    document.Save();
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
                    if (document.WorkbookPart.Workbook.FileSharing != null)
                    {
                        return Protection.Locked;
                    }

                    foreach (var part in document.WorkbookPart.WorksheetParts)
                    {
                        var protections = part.Worksheet.Elements<SheetProtection>();

                        if (protections.Any())
                            return Protection.Locked;
                    }
                }
            }
            catch (System.IO.FileFormatException)
            {
                return Protection.Crypted;
            }

            return Protection.Unlocked;
        }
        public string GetProtectionString(string report)
        {
            return GetProtection(report).ToString();
        }
        public bool SignReport(string report, string signature, string sheet, uint row, string column)
        {
            if (string.IsNullOrEmpty(report))
                throw new ArgumentException("Report file required");

            if (!System.IO.File.Exists(report))
                throw new ArgumentException("Report file must exist");

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(report, true))
                {
                    var workbookPart = document.WorkbookPart;
                    var _sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheet);
                    //TODO _sheet == null ?
                    var worksheetPart = workbookPart.GetPartById(_sheet.Id.Value) as WorksheetPart;
                    //var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    var cell = GetCell(worksheetPart.Worksheet, column, row);

                    cell.CellValue = new CellValue(signature);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);

                    worksheetPart.Worksheet.Save();
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
        public bool IsSigned(string report, string signature, string sheet, uint row, string column)
        {
            if (string.IsNullOrEmpty(report))
                throw new ArgumentException("Report file required");

            if (!System.IO.File.Exists(report))
                throw new ArgumentException("Report file must exist");

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(report, false))
                {
                    var workbookPart = document.WorkbookPart;
                    var _sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheet);
                    //TODO _sheet == null ?
                    var worksheetPart = workbookPart.GetPartById(_sheet.Id.Value) as WorksheetPart;
                    //var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    var cell = GetCell(worksheetPart.Worksheet, column, row);

                    var value = cell.CellValue;

                    if (value.ToString() == signature)
                        return true;
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
        protected static Cell GetCell(Worksheet worksheet, string column, uint row)
        {
            var _row = GetRow(worksheet, row);

            if (_row == null) return null;

            var FirstRow = _row.Elements<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, column + row, true) == 0);

            return FirstRow ?? null;
        }
        protected static Row GetRow(Worksheet worksheet, uint row)
        {
            Row _row = worksheet.GetFirstChild<SheetData>()
                .Elements<Row>()
                .FirstOrDefault(r => r.RowIndex == row) ?? throw new ArgumentException(string.Format("No row with index {0} found in spreadsheet", row));

            return _row;
        }
    }
}
