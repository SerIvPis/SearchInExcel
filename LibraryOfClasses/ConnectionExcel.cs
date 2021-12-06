using Serilog.Debugging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Serilog;
using System.Diagnostics;

namespace LibraryOfClasses
{
    /// <summary>
    /// Класс взаимодействуе с файлом Excel
    /// вытягивает из него данные в обьект DataSet
    /// </summary>
    public class ExtractExcelFiletoDateSet
    {
        public DataSet ExcelFileDataSet { get; set; }

        public string ConnectionString { get; private set; }
       
        public string FName { get; set; }

        public ExtractExcelFiletoDateSet( string connectionString, string fileName )
        {
            ConnectionString = connectionString;
            ExcelFileDataSet = new DataSet(Path.GetFileNameWithoutExtension(fileName));
            FName = fileName;
            Parse( );
        }

        private void Parse( )
        {
            //string connectionString = string.Format( "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';" );
            OdbcConnectionStringBuilder build = new OdbcConnectionStringBuilder( ConnectionString )
            {
                [ "Data Source" ] = FName
            };

            using (OdbcConnection con = new OdbcConnection( build.ConnectionString ))
            {
                con.Open( );
                Log.Information( $"Соединение с файлом {Path.GetFileNameWithoutExtension( FName )} открыто " );
                OdbcCommand cmd = new OdbcCommand( );
                cmd.Connection = con;

                DataTable dtSheet = con.GetSchema( "Tables" );//"TABLES"
               // DisplayData( dtSheet );

                Log.Information( $"Получена схема таблиц файла {Path.GetFileNameWithoutExtension( FName )}" );
                //foreach (DataRow dr in dtSheet.Rows)
                //{
                //string sheetName = dr[ "TABLE_NAME" ].ToString( );    
                string sheetName = @"список$" ;

                    //if (!(sheetName.EndsWith( "$" ) | (sheetName.EndsWith( "$'" ))))
                    //{
                    //    Log.Information( $"\tПропущено\t< {sheetName} >" );
                    //    continue;
                    //}

                    cmd.CommandText = $"SELECT * FROM [{sheetName}]";

                    DataTable dt = new DataTable( );
                    dt.TableName = $"{Path.GetFileNameWithoutExtension( FName )}_{sheetName}";

                    OdbcDataAdapter da = new OdbcDataAdapter( cmd );
                    da.Fill( dt );
                    Log.Information( $"\tДобавлена\t< {sheetName} >" );
                    string List_Columns = "";
                    foreach (DataColumn item in dt.Columns)
                    {
                        List_Columns += $"< {item.ColumnName} >";
                    }
                    Log.Information( List_Columns );
                    ExcelFileDataSet.Tables.Add( dt );
                    con.Close( );
                //}
                Console.WriteLine( $"{FName}" );
            }
        }


        private static void DisplayData( System.Data.DataTable table )
        {
            foreach (System.Data.DataRow row in table.Rows)
            {
                foreach (System.Data.DataColumn col in table.Columns)
                {
                    Console.WriteLine( "{0} = {1}", col.ColumnName, row[ col ] );
                }
                Console.WriteLine( "============================" );
            }
        }
        //Log.Information( $"Соединение с файлом {Path.GetFileNameWithoutExtension( FName )} закрыто " );
        /// <summary>
        /// Поиск слова в DataSet
        /// </summary>
        /// <param name="wordFound"></param>
        public List<string> SearchWordInDataSet( string wordFound )
        {
            List<string> resultList = new List<string>( );
            foreach (DataTable curDt in ExcelFileDataSet.Tables)
            {
                foreach (DataRow dataRow in curDt.Rows)
                {
                    foreach (DataColumn item in curDt.Columns)
                    {
                        if (dataRow[item].ToString( ).Trim( ).ToLowerInvariant( ).Equals( wordFound.Trim( ).ToLowerInvariant( )
                            , StringComparison.InvariantCultureIgnoreCase ))
                        {
                            resultList.Add( PrintDataRow( dataRow, curDt, ExcelFileDataSet ) );
                            //Console.WriteLine($"{PrintDataRow( dataRow, curDt, ExcelFileDataSet )}");
                        }
                    }
                }
            }
            return resultList;
        }

        private string PrintDataRow( DataRow dataRow, DataTable curDt, DataSet excelFileDataSet )
        {
            string resultStr = null;

            if (curDt.TableName.Contains("список"))// Для первого листа "Список"
            {
               resultStr = $"[ {Path.GetFileName( excelFileDataSet.DataSetName )} ] [ {curDt.TableName} ] --->\t" +
               $"{dataRow[ "Обозначение" ]}\t" +
               $"{dataRow[ "Наименование" ]}";
            }
            else if (curDt.Columns.Contains("Инв# № подл#" ))// Для групповой спецификации
            {
                resultStr = $"[ {Path.GetFileName( excelFileDataSet.DataSetName )} ] [ {curDt.TableName} ] --->\t" +
               $"{dataRow[ "F7" ]}\t" +
               $"{dataRow[ "Инв# № дубл#" ]}\t";
            }
            else  // Для одиночной спецификации
            {
                resultStr = $"[ {Path.GetFileName( excelFileDataSet.DataSetName )} ] [ {curDt.TableName} ] --->\t" +
               $"{dataRow[ "Обозначение" ]}\t" +
               $"{dataRow[ "Наименование" ]}\t" +
               $"{dataRow[ "Кол#" ]}";
            }

            return resultStr;
        }
    }
}

     
