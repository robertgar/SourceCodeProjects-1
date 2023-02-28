using System.Data;
using System.Text;

namespace common {
    public class UseCommon {
        public StringBuilder query = new StringBuilder();
        public DataTable table = new DataTable();

        public void Clear() {
            query.Clear();
            table.Rows.Clear();
            table.Columns.Clear();
        }

    }
}