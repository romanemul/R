using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace RRD
{
    public class ExcelConnector:DataTable
    {
        // command Sample
        // select * from [sheet1$];
        
        // project deployment musi byt pro platformu x64 jinak ACE.OLEDB. not installed on this machine.

        private string _ConnectionString;
        private string _FilePath;
        private bool _IsHDR;
        private string _ConnectionStringIsHDR;

        public ExcelConnector()
        {
            
        }
        public ExcelConnector(string Filepath, bool IsHDR)
        {
            _FilePath = Filepath;
            
            if (IsHDR) 
            {
                _ConnectionStringIsHDR = "YES";
            }
            else 
            {
                _ConnectionStringIsHDR = "NO";
            }
            _ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                                            "Data Source=" + Filepath + ";" +                                                            
                                                            "Extended Properties=\"Excel 12.0 Macro;" +
                                                            "HDR="+_ConnectionStringIsHDR+"\";";
        }


        public void ExcelConnectorNOHDR(string Filepath)
        {
            _FilePath = Filepath;
            _ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                                                            "Data Source=" + Filepath + ";" +
                                                            "Extended Properties=\"Excel 12.0 Macro;" +
                                                            "HDR=NO\";";
        }

        public string ConnectionString
        {
            get => _ConnectionString;
            set => _ConnectionString = value;
        }

        

        public DataTable Select(string Command)
        {
            DataTable dt = new DataTable();
            return dt = Connect(Command);
        }

        private DataTable Connect(string Command)
        {
            DataSet ds = new DataSet();

            using (var odConnection = new OleDbConnection(_ConnectionString))
            {
                odConnection.Open();

                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = odConnection;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = Command;

                    using (OleDbDataAdapter ExcelAdapter = new OleDbDataAdapter(cmd))
                    {
                        ExcelAdapter.Fill(ds);
                    }
                }
                odConnection.Close();
                return ds.Tables[0];
            }
        }
    }
}
