using System.Data.SqlClient;

namespace connection{
    internal class ConnectToSQL{
        public SqlConnection getStringConnection() {
            return new SqlConnection("User ID=JoshuaS;Database=GuatemalaDigital;PASSWORD=malacuac;server=remoto.guatemaladigital.org;Connect Timeout=30");
        }
    }
}