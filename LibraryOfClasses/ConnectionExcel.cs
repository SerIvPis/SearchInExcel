using Serilog.Debugging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Serilog;


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
        public string FName { get; set; }

        public ExtractExcelFiletoDateSet( string connectionString, string fileName )
        {
            ConnectionString = connectionString;
            ExcelFileDataSet = new DataSet( fileName );
            FName = fileName;
            Parse( );
        }

        private void Parse( )
        {
            //string connectionString = string.Format( "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';" );
            OleDbConnectionStringBuilder build = new OleDbConnectionStringBuilder( ConnectionString );
            build[ "Data Source" ] = FName;
            
            //DataSet excelDataSet = new DataSet( );
            using (OleDbConnection con = new OleDbConnection( build.ConnectionString ))
            {
                con.Open( );
                Log.Information( $"Соединение с файлом {Path.GetFileNameWithoutExtension( FName )} открыто " );
                OleDbCommand cmd = new OleDbCommand( );
                cmd.Connection = con;

                DataTable dtSheet = con.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null );
                //Log.Information( $"Получена схема таблиц файла {Path.GetFileNameWithoutExtension( fileName )}" );

                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr[ "TABLE_NAME" ].ToString( );
                    
                    if ( !( sheetName.EndsWith( "$" ) | (sheetName.EndsWith( "$'" ) ))) 
                    {
                        Log.Information( $"\tПропущено\t< {sheetName} >" );
                        continue;
                    }

                    cmd.CommandText = $"SELECT * FROM [{sheetName}]";

                    DataTable dt = new DataTable(  );
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter( cmd );
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
                }
            }
            Log.Information( $"Соединение с файлом {Path.GetFileNameWithoutExtension( FName )} закрыто " );
        }

        /// <summary>
        /// Поиск слова в DataSet
        /// </summary>
        /// <param name="wordFound"></param>
        public /*List<string>*/ void SearchWordInDataSet( string wordFound )
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
                            //resultList.Add( PrintDataRow( dataRow, curDt, ExcelFileDataSet ) );
                            Console.WriteLine($"{PrintDataRow( dataRow, curDt, ExcelFileDataSet )}");
                        }
                    }
                }
            }
           // return resultList;
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
