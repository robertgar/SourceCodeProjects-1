using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Text;

namespace connection {
    public class SQL {
        internal class Connection {
            public SqlConnection getConnection() {
                return new SqlConnection("User ID=JoshuaS;Database=GuatemalaDigital;PASSWORD=malacuac;server=remoto.guatemaladigital.org;Connect Timeout=30");
            }
        }

        private SqlConnection connection = new Connection().getConnection();
        private DataTable _buffer = new DataTable();
        private StringBuilder _query = new StringBuilder();

        public void fillTable(ref StringBuilder query, ref DataTable table) {
            table.Rows.Clear();
            table.Columns.Clear();

            connection.Open();

            try {
                SqlDataAdapter adapter = new SqlDataAdapter(query.ToString(), connection);
                adapter.Fill(table);
            }catch(Exception){}

            connection.Close();
        }

        public string getValue(ref StringBuilder query, [Optional] string Error) {
            fillTable(ref query, ref _buffer);

            try {
                return _buffer.Rows[0][0].ToString();
            } catch(Exception) {
                return Error;
            }
        }

        public int getNat(ref StringBuilder query, [Optional] int Error) {
            fillTable(ref query, ref _buffer);

            try {
                return int.Parse(_buffer.Rows[0][0].ToString());
            } catch (Exception) {
                return Error;
            }
        }
    }
}