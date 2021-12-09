using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibraryOfClasses
{
    public class ExcelDAL
    {
        private OdbcConnection odbcConnection = null;
        private string fName = null ;

        public ExcelDAL( )
        {
            //Log.Logger = new LoggerConfiguration( )
            //    .WriteTo.File( "logExcel.txt", rollingInterval: RollingInterval.Minute )
            //    .CreateLogger( );
        }
        public void OpenConnection( string connectionString )
        {
            odbcConnection = new OdbcConnection( );
            odbcConnection.ConnectionString = connectionString;
            odbcConnection.Open( );
        }

        public void CloseConnection( )
        {
            odbcConnection.Close( );
        }

        public List<DataTable> GetAllSheets( string fileName )
        {
            fName = fileName;
            List<DataTable> listDataTable = new List<DataTable>( );
            var dtSheet = odbcConnection.GetSchema( OdbcMetaDataCollectionNames.Tables );
            //Log.Information( $"Получена схема таблиц файла {Path.GetFileNameWithoutExtension( FName )}" );

            foreach (DataRow dr in dtSheet.Rows)
            {
                string sheetName = dr[ "TABLE_NAME" ].ToString( );

                if (!(sheetName.EndsWith( "$" ) | (sheetName.EndsWith( "$'" ))))
                {
                    //Log.Information( $"\tПропущено\t< {sheetName} >" );
                    continue;
                }
                listDataTable.Add( GetTableFromSheets( sheetName ) );
                Log.Information( $"\tДобавлен\t< {sheetName} >" );
            }
            return listDataTable;
        }

        private DataTable GetTableFromSheets( string sheetName )
        {
            DataTable dataTable = new DataTable( );
            string sqlQuery = $"SELECT * FROM [{sheetName}]";

            using (var cmd = odbcConnection.CreateCommand( ))
            {
                cmd.CommandText = sqlQuery;
                OdbcDataReader dataReader = cmd.ExecuteReader( );
                dataTable.Load( dataReader );
                dataTable.TableName = $"{Path.GetFileName(fName)}_{sheetName}";
                dataReader.Close( );
            }
            //Console.WriteLine($"\t добавлена --> {sheetName}");
            return dataTable;
        }


    }
}
