using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibraryOfClasses
{
    public class SearchInDirectory
    {
        public static List<string> Begin( string wordFound )
        {
            List<string> resultSearch = new List<string>( ); //Список найденных строк
            DirectoryInfo rootDir = Directory.CreateDirectory( Directory.GetCurrentDirectory( ) );
            ExtractExcelFiletoDateSet Excel = null;

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
