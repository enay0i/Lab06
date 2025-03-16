using System;
using System.IO;
using System.Data;
using OfficeOpenXml;
using ExcelDataReader;

namespace Test_QuocCuong
{
    internal class ExcelProvider
    {
        private static DataTable _excelDataTable;

        public static DataTable ReadExcel(string filePath)
        {
            if (_excelDataTable != null)
                return _excelDataTable;

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet dataSet = reader.AsDataSet();
                    _excelDataTable = dataSet.Tables[0];
                    return _excelDataTable;
                }
            }
        }

        public static void WriteResultToExcel(string filePath, string sheetName, int rowIndex, string actualResult, string status)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName] ?? package.Workbook.Worksheets[0];

                    int colActualResult = 7;
                    int colStatus = 8;

                    int totalRows = worksheet.Dimension?.Rows ?? 0;

                    worksheet.Cells[rowIndex, colActualResult].Value = actualResult;
                    worksheet.Cells[rowIndex, colStatus].Value = status;

                    package.Save();
                    Console.WriteLine($"excel bri");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi ghi vào Excel: " + ex.Message);
            }
        }

        public static bool ValidateExcelData(string filePath, string sheetName, int expectedTotal, int expectedCurrentMonthRevenue)
        {
            try
            {
                DataTable excelData = ReadExcel(filePath);
                if (excelData == null) return false;

                int sum = 0;
                int currentMonth = DateTime.Now.Month;
                int currentMonthRevenue = 0;

                for (int i = 0; i < 12; i++)
                {
                    if (excelData.Rows.Count > i + 1 && excelData.Columns.Count > 1)
                    {
                        object cellValue = excelData.Rows[i + 1][1];
                        if (cellValue != null && int.TryParse(cellValue.ToString(), out int value))
                        {
                            sum += value;
                            if (i + 1 == currentMonth) currentMonthRevenue = value;
                        }
                    }
                }
                return sum == expectedTotal && currentMonthRevenue == expectedCurrentMonthRevenue;
            }
            catch(Exception ex) {
            
                Console.WriteLine(ex.Message);
                return false;
            }
        }



        public static bool ValidateCustomerData(string filePath, string sheetName, int totalRow, string firstCus, string lastCus)
        {
            try
            {
                DataTable excelData = ReadExcel(filePath);
                if (excelData == null)
                {
                    return false;
                }
                int dataRowCount = excelData.Rows.Count-1;
                if (dataRowCount != totalRow)
                {
                    return false;
                }

                if (dataRowCount > 0)
                {
                    string firstCustomer = excelData.Rows[1][0].ToString().Trim();
                    if (!firstCustomer.Equals(firstCus, StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    string lastCustomer = excelData.Rows[dataRowCount - 1][0].ToString().Trim();
                    if (!lastCustomer.Equals(lastCus, StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        public static bool ValidateEmptyCustomerData(string filePath, string sheetName, string[] expectedCot)
        {
            try
            {
                DataTable excelData = ReadExcel(filePath);

                if (excelData == null || excelData.Columns.Count < expectedCot.Length)
                {
                    return false;
                }

                for (int i = 0; i < expectedCot.Length; i++)
                {
                    string actualHeader = excelData.Columns[i].ColumnName.Trim();
                    if (!actualHeader.Equals(expectedCot[i], StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }
                if (excelData.Rows.Count > 0)
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi đọc file Excel: " + ex.Message);
                return false;
            }
        }
    }
}