using ImportExcelDocToSQLApp.Services.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.SqlClient;
using Dapper;

namespace ImportExcelDocToSQLApp.Services.Implementations
{
    public class DataReaderService : IDataReaderService
    {
        public DataReaderService()
        {

        }

        public DataTable ReadExcelToDataTable(string filePath)
        {
            DataTable dataTable = new DataTable();

            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                // Add columns to the DataTable based on the Excel file's columns
                for (int col = 1; col <= xlRange.Columns.Count; col++)
                {
                    string header = (string)(xlRange.Cells[4, col] as Excel.Range).Text;
                    if (string.IsNullOrWhiteSpace(header)) header = $"Column{col}";
                    dataTable.Columns.Add(header);
                }

                // Populate data starting from the row after the header
                for (int row = 5; row <= xlRange.Rows.Count; row++)
                {
                    string firstCellValue = (string)(xlRange.Cells[row, 1] as Excel.Range).Text;
                    if (firstCellValue == "Derivatives" || firstCellValue.StartsWith("Total for") || firstCellValue.StartsWith("Grand Total"))
                    {
                        continue; // Skip the "Derivatives", "Total for", and "Grand Total" rows
                    }

                    bool isEmptyRow = true;
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= xlRange.Columns.Count; col++)
                    {
                        var cellValue = xlRange.Cells[row, col] as Excel.Range;
                        if (!string.IsNullOrWhiteSpace((string)cellValue.Text))
                        {
                            isEmptyRow = false;
                        }
                        if (dataTable.Columns[col - 1].ColumnName == "Date" && DateTime.TryParseExact((string)cellValue.Text, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                        {
                            dataRow[col - 1] = parsedDate;
                        }
                        else
                        {
                            dataRow[col - 1] = cellValue.Text;
                        }
                    }

                    if (!isEmptyRow)
                    {
                        dataTable.Rows.Add(dataRow);
                    }
                }

                // Cleanup
                xlWorkbook.Close(false);
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                // Handle COM exceptions that may occur when interacting with Excel
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }
            return dataTable;
        }
        public void DisplayDataTable(DataTable dataTable)
        {
            Console.WriteLine($"Excel Data (Rows: {dataTable.Rows.Count}, Columns: {dataTable.Columns.Count})");

            // Display column headers
            Console.WriteLine(string.Join("\t", dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName)));

            // Display rows
            foreach (DataRow row in dataTable.Rows)
            {
                Console.WriteLine(string.Join("\t", row.ItemArray));
            }

            Console.WriteLine();
        }

        public bool IsFileProcessed(FileInfo fileInfo, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "SELECT COUNT(1) FROM ProcessedFiles WHERE FileName = @FileName AND LastModified = @LastModified";
                return connection.QuerySingle<int>(query, new { FileName = fileInfo.Name, LastModified = fileInfo.LastWriteTime }) > 0;
            }
        }

        public void LogProcessedFile(FileInfo fileInfo, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "INSERT INTO ProcessedFiles (FileName, LastModified) VALUES (@FileName, @LastModified)";
                connection.Execute(query, new { FileName = fileInfo.Name, LastModified = fileInfo.LastWriteTime });
            }
        }

        public void SaveToDatabase(DataTable dataTable, Dictionary<string, string> columnMapping, string connectionString, Dictionary<string, Func<object>> additionalColumns, string logFolder)
        {
            string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                    foreach (DataRow row in dataTable.Rows)
                    {
                        var parameters = new DynamicParameters();

                        // Map Excel columns to SQL columns
                        foreach (var map in columnMapping)
                        {
                            if (!dataTable.Columns.Contains(map.Key))
                            {
                                LogMessage(logFolder, datetime, $"Column '{map.Key}' does not exist in DataTable. Skipping...");
                                continue;
                            }

                            object value = row[map.Key] == DBNull.Value ? (object)DBNull.Value : row[map.Key]?.ToString();

                            // Handle float conversion for specific columns
                            if (map.Value == "MTMYield" || map.Value == "MarkPrice" || map.Value == "SpotRate" || map.Value == "PreviousMTM" || map.Value == "PreviousPrice" || map.Value == "DeltaValue"
                                || map.Value == "Strike" || map.Value == "PremiumOnOption" || map.Value == "Volatility" || map.Value == "ContractsTraded" || map.Value == "OpenInterest" || map.Value == "Delta")
                            {
                                string floatString = value.ToString().Replace(",", ""); // Remove any commas

                                // Handle negative values in parentheses for "Delta" column
                                if (map.Value == "Delta" && floatString.StartsWith("(") && floatString.EndsWith(")"))
                                {
                                    floatString = "-" + floatString.Trim('(', ')');
                                }

                                if (value != DBNull.Value && decimal.TryParse(floatString, out decimal decimalValue))
                                {
                                    value = decimalValue;
                                }
                                else
                                {
                                    value = 0;
                                }
                            }

                            // Handle Date conversion for date columns
                            if (map.Value == "ExpiryDate" && value != DBNull.Value)
                            {
                                if (DateTime.TryParseExact(value.ToString(), "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dateValue))
                                {
                                    value = dateValue;
                                }
                                else
                                {
                                    value = DBNull.Value; // Set to DBNull.Value if conversion fails
                                }
                            }

                            parameters.Add($"@{map.Value}", value);
                        }

                        // Add additional columns
                        foreach (var additionalColumn in additionalColumns)
                        {
                            parameters.Add($"@{additionalColumn.Key}", additionalColumn.Value());
                        }

                        var allColumns = columnMapping.Values.Concat(additionalColumns.Keys);
                        var parameterNames = allColumns.Select(col => $"@{col}");

                        string sql = $"INSERT INTO DailyMTM ({string.Join(", ", allColumns)}) VALUES ({string.Join(", ", parameterNames)})";

                        try
                        {
                            connection.Execute(sql, parameters);
                        }
                        catch (SqlException ex)
                        {
                            LogError(logFolder, datetime, ex);
                        }
                        catch (Exception ex)
                        {
                            LogError(logFolder, datetime, ex);
                        }

                    }
                }
                catch (SqlException ex)
                {
                    LogError(logFolder, datetime, ex);
                }
                catch (Exception ex)
                {
                    LogError(logFolder, datetime, ex);
                }
            }
        }

        static void LogError(string logFolder, string datetime, Exception ex)
        {
            string logFilePath = Path.Combine(logFolder, $"ErrorLog_{datetime}.log");
            using (StreamWriter sw = new StreamWriter(logFilePath, true))
            {
                sw.WriteLine($"{DateTime.Now} - Error: {ex.Message}");
                sw.WriteLine(ex.ToString());
            }
        }

        static void LogMessage(string logFolder, string datetime, string message)
        {
            string logFilePath = Path.Combine(logFolder, $"Log_{datetime}.log");
            using (StreamWriter sw = new StreamWriter(logFilePath, true))
            {
                sw.WriteLine($"{DateTime.Now} - Message: {message}");
            }
        }
    }
}
