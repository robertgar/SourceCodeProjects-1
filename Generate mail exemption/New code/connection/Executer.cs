using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;

namespace connection{
    public class Execute{
        private SqlConnection connector = new ConnectToSQL().getStringConnection();
        private DataTable Buffer = new DataTable();
        private Boolean isSimulation;

        public void setSimulation(ref Boolean deAMentis) {
            this.isSimulation = deAMentis;
        }

        public String getValue(ref StringBuilder query, [Optional] String Error){
            fillTable(ref query, ref Buffer);

            try {
                return Buffer.Rows[0][0].ToString();
            }catch (Exception) {
                return Error;
            }
        }

        public void fillTable(ref StringBuilder query, ref DataTable Tablita) {
            Tablita.Clear();
            Tablita.Rows.Clear();
            Tablita.Columns.Clear();

            connector.Open();
            try{
                SqlDataAdapter reader = new SqlDataAdapter(query.ToString(), connector);
                reader.Fill(Tablita);
            } catch (Exception) {}
            connector.Close();
        }

        public int getNat(ref StringBuilder query, [Optional] int Error){
            try{
                fillTable(ref query, ref Buffer);
                return int.Parse(Buffer.Rows[0][0].ToString());
            }catch (Exception) {
                return Error;
            }
        }

        public void executeQuery(ref StringBuilder query, [Optional] String msg) {
            if (isSimulation) {
                Console.WriteLine(msg);
                return;
            }

            try {
                connector.Open();
                SqlCommand command = new SqlCommand(query.ToString(), connector);
                command.CommandType = CommandType.Text;
                command.ExecuteNonQuery();
                connector.Close();
            }catch (Exception) { }
        }

        public String getParameter(int ParameterCode, [Optional] String error) {
            StringBuilder query = new StringBuilder();
            query.AppendLine(" select");
            query.AppendLine("    Valor");
            query.AppendLine(" from");
            query.AppendLine("    Parametro");
            query.AppendLine(" where");
            query.Append("    CodigoParametro = ").Append(ParameterCode);
            
            return getValue(ref query, error);
        }

        public String getChanel(int Chanel) {
            StringBuilder query = new StringBuilder();
            query.AppendLine(" select");
            query.AppendLine("    w.Url");
            query.AppendLine(" from");
            query.AppendLine("    Alerta as a");
            query.AppendLine("    inner join Webhook as w on w.CodigoWebhook = a.CodigoWebhook");
            query.AppendLine(" where");
            query.AppendLine("    a.Activo = 1");
            query.Append("    and a.CodigoAlerta = ").Append(Chanel);

            return getValue(ref query);
        }
    }
}
