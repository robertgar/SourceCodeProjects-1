using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;

namespace connection{
    public class Execute{
        private SqlConnection connector = new ConnectToSQL().getStringConnection();
        private DataTable Buffer = new DataTable();

        public String getValue(ref StringBuilder query, String Error = ""){
            fillTable(ref query, ref Buffer);

            try {
                return Buffer.Rows[0][0].ToString();
            }catch (Exception ex) {
                return Error;
            }
        }

        public void fillTable(ref StringBuilder query, ref DataTable Tablita) {
            Tablita.Clear();

            try{
                connector.Open();
                SqlDataAdapter reader = new SqlDataAdapter(query.ToString(), connector);
                reader.Fill(Tablita);
                connector.Close();
            }
            catch (Exception e) { }
        }

        public int getNat(ref StringBuilder query, [Optional] int Error){
            try{
                fillTable(ref query, ref Buffer);

                return int.Parse(Buffer.Rows[0][0].ToString());
            }
            catch (Exception e) {
                return Error;
            }
        }

        public void executeQuery(ref StringBuilder query) {
            try {
                connector.Open();
                SqlCommand command = new SqlCommand(query.ToString(), connector);
                command.CommandType = CommandType.Text;
                command.ExecuteNonQuery();
                connector.Close();
            }catch (Exception e) { }
        }
    }
}
