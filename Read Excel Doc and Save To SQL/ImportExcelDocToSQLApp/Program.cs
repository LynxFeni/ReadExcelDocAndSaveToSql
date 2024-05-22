using ImportExcelDocToSQLApp.Services.Implementations;
using ImportExcelDocToSQLApp.Services.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExcelDocToSQLApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string excelFilesDirectory = ConfigurationManager.AppSettings["ExcelDirectory"]; //your excel directory 
            //retrieve a list of all Excel files in the specified directory
            var excelFiles = Directory.GetFiles(excelFilesDirectory, "*.xls");

            string connectionString = ConfigurationManager.ConnectionStrings["myDabaseConnectionString"].ConnectionString; //your connection string here
            string logFolder = ConfigurationManager.AppSettings["logFolder"]; // your log folder directory


            // Use the service provider to resolve the dependencies
            var serviceProvider = ConfigureServices();
            var dataReaderService = serviceProvider.GetService<IDataReaderService>();

            

            foreach (var excelFilePath in excelFiles)
            {
                FileInfo fileInfo = new FileInfo(excelFilePath);
                if (!dataReaderService.IsFileProcessed(fileInfo, connectionString))
                {
                    DataTable dataTable = dataReaderService.ReadExcelToDataTable(excelFilePath);

                    // Define column mappings
                    var columnMapping = new Dictionary<string, string>
                    {
                        { "Contract Details", "Contract" },
                        { "Column3", "ExpiryDate" },
                        { "Column4", "Classification" },
                        { "Strike", "Strike" },
                        { "Call /Put", "CallPut" },
                        { "MTM Yield", "MTMYield" },
                        { "Mark Price", "MarkPrice" },
                        { "Spot Rate", "SpotRate" },
                        { "Previous MTM", "PreviousMTM" },
                        { "Previous Price", "PreviousPrice" },
                        { "Premium On Option", "PremiumOnOption" },
                        { "Volatility", "Volatility" },
                        { "Delta", "Delta" },
                        { "Delta Value", "DeltaValue" },
                        { "ContractsTraded", "ContractsTraded" },
                        { "Open Interest", "OpenInterest" }
                    };

                    // Define additional columns that are not present in the Excel file
                    var additionalColumns = new Dictionary<string, Func<object>>
                    {
                        { "FileDate", () => DateTime.Now } // Adding current date
                    };

                    // Save valid rows to SQL database
                    dataReaderService.SaveToDatabase(dataTable, columnMapping, connectionString, additionalColumns, logFolder);

                    // Log the processed file
                    dataReaderService.LogProcessedFile(fileInfo, connectionString);
                }
            }
            Console.WriteLine("Press any key to continue...");
            Console.ReadLine();
        }
        private static ServiceProvider ConfigureServices()
        {
            var services = new ServiceCollection();
            // Register the interfaces and implementations with the DI container
            services.AddTransient<IDataReaderService, DataReaderService>();
            return services.BuildServiceProvider();
        }
 
    }
}
