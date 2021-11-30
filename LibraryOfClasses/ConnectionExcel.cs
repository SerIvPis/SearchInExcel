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
            using (OleDbConnection con = new OleDbConnection( build.ConnectionString ))
            {
                con.Open( );
                OleDbCommand cmd = new OleDbCommand( );
                cmd.Connection = con;

                DataTable dtSheet = con.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null );

                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr[ "TABLE_NAME" ].ToString( );

                    if ( !( sheetName.EndsWith( "$" ) | (sheetName.EndsWith( "$'" ) ))) 
                    {
                        continue;
                    }

                  

                    cmd.CommandText = $"SELECT * FROM [{sheetName}]";

                    DataTable dt = new DataTable(  );
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter( cmd );
                    da.Fill( dt );

                    ExcelFileDataSet.Tables.Add( dt );
                    con.Close( );
                }
            }
        }

       


        /// <summary>
        /// Поиск слова в DataSet
        /// </summary>
        /// <param name="wordFound"></param>
       /* public List<string> SearchWordInDataSet( string wordFound )
        {
            List<string> resultList = new List<string>( );
            foreach (DataTable curDt in ExcelFileDataSet.Tables)
            {
                foreach (DataRow dataRow in curDt.Rows)
                {
                    foreach (var item in dataRow.ItemArray)
                    {
                        if (item.ToString().Trim().ToLowerInvariant().Equals(wordFound.Trim().ToLowerInvariant()
                            ,StringComparison.InvariantCultureIgnoreCase))
                        {
                            resultList.Add(PrintDataRow( dataRow, curDt, ExcelFileDataSet ));
                        }
                    }
                }
            }
            return resultList;
        }*/

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
                        }
                    }
                }
            }
            return resultList;
        }

        private string PrintDataRow( DataRow dataRow, DataTable curDt, DataSet excelFileDataSet )
        {
            string resultStr = null;

            if (curDt.TableName.Contains("список"))
            {
               resultStr = $"[ {Path.GetFileName( excelFileDataSet.DataSetName )} ] [ {curDt.TableName} ] --->\t" +
               $"{dataRow[ "Обозначение" ]}\t" +
               $"{dataRow[ "Наименование" ]}";
            }
            else
            {
                resultStr = $"[ {Path.GetFileName( excelFileDataSet.DataSetName )} ] [ {curDt.TableName} ] --->\t" +
               $"{dataRow[ "Обозначение" ]}\t" +
               $"{dataRow[ "Наименование" ]}\t" +
               $"{dataRow[ "Кол#" ]}";
            }
            
            //string resultStr =  $"[ {Path.GetFileName( excelFileDataSet.DataSetName)} ] [ {curDt.TableName} ] --->\t" +
            //    $"{dataRow.ItemArray[6]}\t" +
            //    $"{dataRow.ItemArray[ 15 ]}\t" +
            //    $"{dataRow.ItemArray[ 21 ]}\t" +
            //    $"{dataRow.ItemArray[ 23 ]}" ;
            //foreach (var item in dataRow.ItemArray)
            //{
            //    //resultStr += $"{item.ToString( )} ";
            //}
            return resultStr;
        }

        
    }
}
