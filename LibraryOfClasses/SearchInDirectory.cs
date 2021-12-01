using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Serilog;
using Serilog.Sinks.File;


namespace LibraryOfClasses
{
    /// <summary>
    /// Ищет в *.xls файлах слово
    /// </summary>
    public class SearchInDirectory
    {
        public static List<string> Begin( string wordFound )
        {
            Log.Logger = new LoggerConfiguration( )
                 .WriteTo.File( "logExcel.txt", rollingInterval: RollingInterval.Infinite )
                 .CreateLogger( );
            List<string> resultSearch = new List<string>( ); //Список найденных строк
            DirectoryInfo rootDir = Directory.CreateDirectory( Directory.GetCurrentDirectory( ) );
            ExtractExcelFiletoDateSet Excel = null;// Объект для отображения файла Excel на DataSet

            foreach (FileInfo cfile in rootDir.GetFiles( "*.xls", SearchOption.TopDirectoryOnly ))
            {
                Log.Information( $"Файл-< {cfile.Name} >" );
                Excel = new ExtractExcelFiletoDateSet(
                    ConfigurationManager.ConnectionStrings[ "Excel16" ].ConnectionString, cfile.FullName );
                Log.Information( $"Импорт данных из файла-< {cfile.Name} >" );
                resultSearch.AddRange( Excel.SearchWordInDataSet( wordFound ) );
                Log.Information( $"Поиск завершен-< {cfile.Name} >" );
            }
            return resultSearch;
        }
    }
}
