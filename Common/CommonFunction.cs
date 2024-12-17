using Microsoft.Data.SqlClient;
using System.Data;

namespace PCS_SYSTEMS.Common
{
    public class CommonFunction
    {
        public static readonly string connectionString = "Data Source=10.201.21.84,50150;Initial Catalog=PCS;Persist Security Info=True;User ID=cimitar2;Password=TFAtest1!2!;Trust Server Certificate=True";
        public static readonly string SUCCESS = "SUCCESS";
        public static readonly string FAIL = "FAIL";
        public static readonly string ERROR = "ERROR";
        private readonly IHostEnvironment _environment;
        public CommonFunction(IHostEnvironment environment)
        {
            _environment = environment;
        }
        public static List<Dictionary<string, object>> GetDataFromProcedure(SqlDataReader reader)
        {
            List<Dictionary<string, object>> dictionary = new List<Dictionary<string, object>>();
            while (reader.Read())
            {
                var row = new Dictionary<string, object>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var columnName = reader.GetName(i);
                    var columnValue = reader.GetValue(i);
                    row.Add(columnName, columnValue);
                }
                dictionary.Add(row);
            }
            return dictionary;
        }
        public static void LogInfo(string DefaultConnection, string idCard, string info, string typeLog, string function)
        {
            using var connection = new SqlConnection(DefaultConnection);
            using var command = new SqlCommand("AddLogInfo", connection) { CommandType = CommandType.StoredProcedure };

            // Thêm các tham số cho stored procedure (nếu cần)
            command.Parameters.AddWithValue("@idCard", idCard);
            command.Parameters.AddWithValue("@info", info);
            command.Parameters.AddWithValue("@typeLog", typeLog);
            command.Parameters.AddWithValue("@function", function);

            connection.Open();
            var reader = command.ExecuteReader();
        }
    }
}
