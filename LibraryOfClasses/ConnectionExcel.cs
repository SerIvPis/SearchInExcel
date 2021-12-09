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
        public string world = "";

        public ExtractExcelFiletoDateSet( string connectionString, string fileName )
        {
            ConnectionString = connectionString;
            FName = fileName;
            Parse( );
        }

        public ExtractExcelFiletoDateSet( string connectionString, string fileName, string world )
        {
            ConnectionString = connectionString;
            FName = fileName;
            this.world = world;
            Parse( );
        }

        //private DataSet Parse( )
        private  void Parse( )
        {
            DataSet result = new DataSet( Path.GetFileNameWithoutExtension( FName ));

            OdbcConnectionStringBuilder build = new OdbcConnectionStringBuilder( ConnectionString )
            {
                [ "Dbq" ] = FName
            };

            try
            {
                using (OdbcConnection con = new OdbcConnection( build.ConnectionString ))
                {
                    using (var cmd = con.CreateCommand( ))
                    {
                        Log.Information( $"перед con.Open() Connection state = {con.State}" );
                        con.ConnectionTimeout = 0;
                        //await con.OpenAsync( );
                        con.Open( );
                        Log.Information( $"после con.Open() Connection state = {con.State}" );

                        Log.Information( $"Соединение с файлом {Path.GetFileNameWithoutExtension( FName )} открыто " );

                        //OdbcCommand cmd = new OdbcCommand( );

                        var dtSheet = con.GetSchema( OdbcMetaDataCollectionNames.Tables );

                        //DisplayData( dtSheet );

                        Log.Information( $"Получена схема таблиц файла {Path.GetFileNameWithoutExtension( FName )}" );
                        foreach (DataRow dr in dtSheet.Rows)
                        {
                            string sheetName = dr[ "TABLE_NAME" ].ToString( );

                            if (!(sheetName.EndsWith( "$" ) | (sheetName.EndsWith( "$'" ))))
                            {
                                Log.Information( $"\tПропущено\t< {sheetName} >" );
                                continue;
                            }

                            cmd.CommandText = $"SELECT * FROM [{sheetName}]";

                            using (OdbcDataReader reader = cmd.ExecuteReader( ))
                            {
                                while (reader.Read( ))
                                {
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        if (reader.GetValue( i ).ToString( ).Trim( ).ToLowerInvariant( ).Equals( world.Trim( ).ToLowerInvariant( )
                                                 , StringComparison.InvariantCultureIgnoreCase ))
                                        {
                                            Console.WriteLine( PrintRow( reader, sheetName ) );
                                        }

                                    }
                                }

                            }

                            //OdbcDataAdapter da = new OdbcDataAdapter( cmd );
                            //DataTable dt = new DataTable( );
                            //dt.TableName = $"{Path.GetFileNameWithoutExtension( FName )}_{sheetName}";
                            //da.Fill( dt );
                            Log.Information( $"\tДобавлена\t< {sheetName} >" );
                            //string List_Columns = "";
                            //foreach (DataColumn item in dt.Columns)
                            //{
                            //    List_Columns += $"< {item.ColumnName} >";
                            //}
                            //Log.Information( List_Columns );
                            //  result.Tables.Add( dt );
                        }

                    }

                }

                Console.WriteLine( $"{FName}" );
                ExcelFileDataSet = result;

            }
            catch (Exception e)
            {

                Log.Fatal( $"Ошибка в {e.StackTrace} - {e.Message}" );
                Console.WriteLine( e.Message );
                //Console.WriteLine( e.StackTrace );
            }
            
            //return result;
        }

       

        private string PrintRow( OdbcDataReader reader, string sheetName )
        {
            StringBuilder result = new StringBuilder( );
            for (int i = 0; i < reader.FieldCount; i++)
            {
                result.Append( reader.GetValue( i ).ToString() );
            }
            
            //if (sheetName.Contains( "список" ))// Для первого листа "Список"
            //{
            //    resultStr = $"[ {Path.GetFileName( excelFileDataSet.DataSetName )} ] [ {curDt.TableName} ] --->\t" +
            //    $"{dataRow[ "Обозначение" ]}\t" +
            //    $"{dataRow[ "Наименование" ]}";
            //}
            //else if (sheetName.Contains( "Инв# № подл#" ))// Для групповой спецификации
            //{
            //    resultStr = $"[ {Path.GetFileName( excelFileDataSet.DataSetName )} ] [ {curDt.TableName} ] --->\t" +
            //   $"{dataRow[ "F7" ]}\t" +
            //   $"{dataRow[ "Инв# № дубл#" ]}\t";
            //}
            //else  // Для одиночной спецификации
            //{
            //    resultStr = $"[ {Path.GetFileName( excelFileDataSet.DataSetName )} ] [ {curDt.TableName} ] --->\t" +
            //   $"{dataRow[ "Обозначение" ]}\t" +
            //   $"{dataRow[ "Наименование" ]}\t" +
            //   $"{dataRow[ "Кол#" ]}";
            //}

            return result.ToString();
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
            try
            {
                foreach (DataTable curDt in ExcelFileDataSet.Tables)
                {
                    foreach (DataRow dataRow in curDt.Rows)
                    {
                        foreach (DataColumn item in curDt.Columns)
                        {
                            if (dataRow[ item ].ToString( ).Trim( ).ToLowerInvariant( ).Equals( wordFound.Trim( ).ToLowerInvariant( )
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
            catch (Exception e)
            {
                
                Log.Fatal( $"Ошибка в {e.StackTrace} - {e.Message}" );
                Console.WriteLine($"{e.Message} метод {e.TargetSite}" );
                return resultList;

            }
           
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

     
