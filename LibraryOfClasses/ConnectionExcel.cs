using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace LibraryOfClasses
{
    /// <summary>
    /// Класс взаимодействуе с файлом Excel
    /// вытягивает из него данные в обьект DataSet
    /// </summary>
    public class ExtractExcelFiletoDateSet
    {
        public string ConnectionString { get; private set; }
        public DataSet ExcelFileDataSet { get; private set; }


        public ExtractExcelFiletoDateSet( string connectionString, string fileName )
        {
            ConnectionString = connectionString;
            ExcelFileDataSet = new DataSet( fileName );
            Parse( fileName );
        }

        private  void Parse( string fileName )
        {
            //string connectionString = string.Format( "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';" );
            OleDbConnectionStringBuilder build = new OleDbConnectionStringBuilder( ConnectionString );
            build[ "Data Source" ] = fileName;

            //DataSet excelDataSet = new DataSet( );

            foreach (var sheetName in GetExcelSheetNames( build.ConnectionString ))
            {
                using (OleDbConnection con = new OleDbConnection( build.ConnectionString ))
                {
                    var listExcelDT = new DataTable( sheetName );
                    string query = $"SELECT * FROM [{sheetName}]";
                    con.Open( );
                    OleDbDataAdapter adapter = new OleDbDataAdapter( query, con );
                    adapter.Fill( listExcelDT );
                    ExcelFileDataSet.Tables.Add( listExcelDT );
                }
            }

            //return excelDataSet;
        }

        /// <summary>
        /// Получить список листов в файле EXcel        /// 
        /// </summary>
        /// <param name="_connectionString"></param>
        /// <returns></returns>
        private List<string> GetExcelSheetNames( string _connectionString )
        {
            OleDbConnection con = null;
            DataTable dt = null;
            con = new OleDbConnection( _connectionString );
            con.Open( );
            dt = con.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null );

            if (dt == null)
            {
                return null;
            }

            List<string> excelSheetNames = new List<string>();
            //int i = 0;

            foreach (DataRow row in dt.Rows)
            {
                if (!row[ "TABLE_NAME" ].ToString( ).Contains( "Print_Area" ))
                {
                    excelSheetNames.Add( row[ "TABLE_NAME" ].ToString( ) );
                }
                //excelSheetNames[ i ] = row[ "TABLE_NAME" ].ToString( );
                //i++;
            }

            return excelSheetNames;
        }


        /// <summary>
        /// Поиск слова в DataSet
        /// </summary>
        /// <param name="wordFound"></param>
        public void SearchWordInDataSet( string wordFound )
        {
            foreach (DataTable curDt in ExcelFileDataSet.Tables)
            {
                foreach (DataRow dataRow in curDt.Rows)
                {
                    foreach (var item in dataRow.ItemArray)
                    {
                        if (item.ToString().Trim().ToLowerInvariant().Equals(wordFound.Trim().ToLowerInvariant()))
                        {
                            PrintDataRow( dataRow, curDt, ExcelFileDataSet );
                        }
                    }
                }
            }
        }

        private void PrintDataRow( DataRow dataRow, DataTable curDt, DataSet excelFileDataSet )
        {
            Console.Write( $"[ {Path.GetFileName( excelFileDataSet.DataSetName)} ] [ {curDt.TableName} ] --->" );
            foreach (var item in dataRow.ItemArray)
            {
                Console.Write( $"{item.ToString( )} " );
            }
            Console.WriteLine( );
        }

        //private void PrintDataRow( DataRow dataRow )
        //{
        //    Console.WriteLine( "------" );
        //    foreach (var item in dataRow.ItemArray)
        //    {
        //        Console.Write( $"{item.ToString( )} " );
        //    }
            
        //}
    }
}
