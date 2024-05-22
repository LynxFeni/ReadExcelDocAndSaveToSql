using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportExcelDocToSQLApp.Services.Interfaces
{
    public interface IDataReaderService
    {
        DataTable ReadExcelToDataTable(string filePath);
        void DisplayDataTable(DataTable dataTable);//use this method for testing perpose

        bool IsFileProcessed(FileInfo fileInfo, string connectionString);
        void LogProcessedFile(FileInfo fileInfo, string connectionString);

        void SaveToDatabase(DataTable dataTable, Dictionary<string, string> columnMapping, string connectionString, Dictionary<string, Func<object>> additionalColumns, string logFolder);

    }
}
