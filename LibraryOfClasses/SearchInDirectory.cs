using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Serilog;
using Serilog.Sinks.File;
using System.Data;

namespace LibraryOfClasses
{
    /// <summary>
    /// Ищет в *.xls файлах слово
    /// </summary>
    public class SearchInDirectory
    {
        public static ConcurrentBag<string> cbResultSearch { get; set; }

        public static void Begin( string wordFound )
        {
            Log.Logger = new LoggerConfiguration( )
                 .WriteTo.File( "logExcel.txt", rollingInterval: RollingInterval.Infinite )
                 .CreateLogger( );
            string[] files = Directory.GetFiles( Directory.GetCurrentDirectory( ), "*.xls", SearchOption.TopDirectoryOnly );

            List<Task> ltask = new List<Task>( );

            foreach (string cfile in files)
            {
                ltask.Add( Task.Run( ( ) =>
                {
                    Log.Information( $"Файл-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                    ExtractExcelFiletoDateSet Excel = new ExtractExcelFiletoDateSet(
                        ConfigurationManager.ConnectionStrings[ "ExcelODBC" ].ConnectionString, cfile);
                    Log.Information( $"Импорт данных из файла-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                    // resultSearch.AddRange( Excel.SearchWordInDataSet( wordFound ) );
                    Log.Information( $"Поиск завершен-< {Path.GetFileNameWithoutExtension( cfile )} >" );
                } ) );
            }
            Task.WaitAll( ltask.ToArray<Task>( ) );
            //return cbResultSearch;
        }

        public static List<string> SearchWordInDataSet( string wordFound )
        {
            //List<string> resultList = new List<string>( );
                  cbResultSearch = new ConcurrentBag<string>( );
         /* 
                   foreach (DataTable curDt in ExcelFileDataSet.Tables)
                   {
                           foreach (DataRow dataRow in curDt.Rows)
                           {
                               foreach (DataColumn item in curDt.Columns)
                               {
                                   if (dataRow[ item ].ToString( ).Trim( ).ToLowerInvariant( ).Equals( wordFound.Trim( ).ToLowerInvariant( )
                                       , StringComparison.InvariantCultureIgnoreCase ))
                                   {
                                       //cbResultSearch.Add( PrintDataRow( dataRow, curDt, ExcelFileDataSet ) );
                                       Console.WriteLine($"{PrintDataRow( dataRow, curDt, ExcelFileDataSet )}");
                                   }
                               }
                           }
                   }

        */
            return cbResultSearch.ToList<string>();
            
        }

        private static string PrintDataRow( DataRow dataRow, DataTable curDt, DataSet excelFileDataSet )
        {
            string resultStr = null;

            if (curDt.TableName.Contains( "список" ))// Для первого листа "Список"
            {
                resultStr = $" [ {curDt.TableName} ] --->\t" +
                $"{dataRow[ "Обозначение" ]}\t" +
                $"{dataRow[ "Наименование" ]}";
            }
            else if (curDt.Columns.Contains( "Инв# № подл#" ))// Для групповой спецификации
            {
                resultStr = $" [ {curDt.TableName} ] --->\t" +
               $"{dataRow[ "F7" ]}\t" +
               $"{dataRow[ "Инв# № дубл#" ]}\t";
            }
            else  // Для одиночной спецификации
            {
                resultStr = $"[ {curDt.TableName} ] --->\t" +
               $"{dataRow[ "Обозначение" ]}\t" +
               $"{dataRow[ "Наименование" ]}\t" +
               $"{dataRow[ "Кол#" ]}";
            }
            return resultStr;
        }
    }
}
