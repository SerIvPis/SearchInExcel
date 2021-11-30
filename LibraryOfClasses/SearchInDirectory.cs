using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibraryOfClasses
{
    /// <summary>
    /// Ищет в *.xls файлах слово
    /// </summary>
    public class SearchInDirectory
    {
        public static List<string> Begin( string wordFound )
        {
            List<string> resultSearch = new List<string>( ); //Список найденных строк
            DirectoryInfo rootDir = Directory.CreateDirectory( Directory.GetCurrentDirectory( ) );
            ExtractExcelFiletoDateSet Excel = null;// Объект для отображения файла Excel на DataSet

            foreach (FileInfo cfile in rootDir.GetFiles( "*.xls", SearchOption.TopDirectoryOnly ))
            {
                Excel = new ExtractExcelFiletoDateSet(
                    ConfigurationManager.ConnectionStrings[ "ExcelHome" ].ConnectionString, cfile.FullName );

                resultSearch.AddRange( Excel.SearchWordInDataSet( wordFound ) );
            }
            return resultSearch;
        }
    }
}
