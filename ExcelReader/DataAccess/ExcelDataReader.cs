using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.DataAccess
{
    public class ExcelDataReader
    {
        private string connectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Spoyi.com.xlsx;" +
                                            "Extended Properties = 'Excel 8.0;HDR=YES'";

        public DataTable GetAll()
        {
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                conn.Open();
                OleDbDataAdapter objDA = new OleDbDataAdapter
                ("select * from [termos$] where Desc is null", conn);
                DataSet excelDataSet = new DataSet();
                objDA.Fill(excelDataSet);
                return excelDataSet.Tables[0];
            }
        }
    }
}
