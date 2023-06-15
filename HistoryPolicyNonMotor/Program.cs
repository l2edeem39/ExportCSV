using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HistoryPolicyNonMotor
{
    public static class Program
    {
        public static SqlConnection _connection;
        static void Main(string[] args)
        {
            String conn = ConfigurationManager.AppSettings["ConnectionString"].ToString();
            String pathFile = ConfigurationManager.AppSettings["FileCsv"].ToString();
            String spName = ConfigurationManager.AppSettings["spName"].ToString();
            String paramYear = ConfigurationManager.AppSettings["year"].ToString();
            String paramMaxrang = ConfigurationManager.AppSettings["maxrang"].ToString();
            Console.WriteLine("Please Wait...");
            var dict = new Dictionary<string, int>();
            dict.Add("@year", Int32.Parse(paramYear));
            dict.Add("@maxrang", Int32.Parse(paramMaxrang));
            var data = GetDataTable(spName, dict, conn);
            ToCSV(data, pathFile);
            Console.WriteLine("Completed...");
        }
        public static DataTable GetDataTable(string storedName, Dictionary<string, int> param, String conn)
        {
            try
            {
                using (_connection = new SqlConnection(conn))
                {
                    _connection.Open();

                    var result = new DataTable();

                    using (SqlCommand command = new SqlCommand(storedName, _connection))
                    {
                        command.CommandType = System.Data.CommandType.StoredProcedure;
                        foreach (var item in param)
                        {
                            command.Parameters.AddWithValue(item.Key.Trim(), item.Value);
                        }

                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataSet dataSet = new DataSet();
                            adapter.Fill(dataSet);
                            DataTable myTable = dataSet.Tables[0];
                            result = myTable;
                        }
                    }
                    _connection.Close();

                    return result;
                }
            }
            catch (Exception ex)
            {
                if (_connection != null && _connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
                throw ex;
            }
        }
        public static void ToCSV(this DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers    
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(","))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value.Replace("\n", "").Replace("\r\n", "").Replace("\r", ""));
                        }
                        else
                        {
                            sw.Write(dr[i].ToString().Replace("\n", "").Replace("\r\n", "").Replace("\r", ""));
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
    }
}
