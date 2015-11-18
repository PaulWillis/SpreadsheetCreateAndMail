using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

    public class DAL
    {
        private string _connectionString;

        static private DAL _instance = null;
        static public DAL Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DAL();
                return _instance;
            }
        }

        private DAL()
        {
            try
            {
                ConnectionStringSettings conStringObject = ConfigurationManager.ConnectionStrings["ConnectionString"];
                if (conStringObject != null)
                {
                    _connectionString = conStringObject.ConnectionString;
                }
                else
                {
                    throw new ApplicationException("Connection string not found.");
                }
            }
            catch (Exception ex)
            {

                throw;
            }

        }
         
        public DataSet GetDataset(string Cycle)
        {
            string stmt = @"mgview.usp_tvc_TrendingTowardWriteOffs";

            var ds = new DataSet();

            using (SqlConnection cnn = new SqlConnection(_connectionString))
            {
                var command = new SqlCommand(stmt, cnn);
                command.CommandTimeout = 30000;
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.AddWithValue("@CycleNum", Cycle);

                var adapter = new SqlDataAdapter(command);

                adapter.Fill(ds);
            }

            return ds;
        } 

    }

