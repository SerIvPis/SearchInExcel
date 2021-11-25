using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LibraryOfClasses;
using System.Configuration;
using System.IO;
using System.Data;

namespace SearchInExcel
{
    class Program
    {
        static void Main( string[] args )
        {
            string wordForSearch = null;
            string stopWord = "й";
            bool isWork = true;
            while (isWork)
            {
                Console.WriteLine( new string( '-', 100 ) );
                Console.WriteLine( "Введите децимальный номер для поиска:" );
                wordForSearch = Console.ReadLine();
                List<string> resultSearch = SearchWords( wordForSearch );
                Console.WriteLine($"Нажмите \"{stopWord}\" чтобы завершить работу или любую клавишу чтобы продолжить...");
                if (Console.ReadLine( ).ToLowerInvariant() == stopWord)
                {
                    isWork = false;
                }
            }
        }

        private static List<string> SearchWords(string wordFound )
        {
            List<string> resultSearch = new List<string>( );
            DirectoryInfo rootDir = Directory.CreateDirectory( Directory.GetCurrentDirectory( ) );
            rootDir.Create( );

            try
            {
                foreach (FileInfo cfile in rootDir.GetFiles( "*.xls", SearchOption.TopDirectoryOnly ))
                {
                    //Console.WriteLine( $"{cfile.FullName}" );
                    ExtractExcelFiletoDateSet Excel = new ExtractExcelFiletoDateSet( 
                        ConfigurationManager.ConnectionStrings[ "Excel16" ].ConnectionString, cfile.FullName);

                    Excel.SearchWordInDataSet( wordFound );
                }
            }
            catch (Exception e )
            {
                Console.WriteLine(e.Message);
            }
            
            return resultSearch;
        }

        
    }
}
