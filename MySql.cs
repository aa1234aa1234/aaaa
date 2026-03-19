using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod
{
    internal class MySql
    {
        private static MySql instance;
        private MySqlConnection conn;

        public MySql()
        {
            conn = new MySqlConnection(string.Format("Server={0};Port={1};Database={2};Uid={3};Pwd={4}", "pcall.kr", 3306, "stock", "stock", "tmxkr1234"));
            conn.Open();
        }

        ~MySql()
        {
            conn.Close();
        }

        public static MySql GetInstance()
        {
            if(instance == null) instance = new MySql();
            return instance;
        }

        public DataRow PollLatestRequest()
        {
            DataSet ds = new DataSet();
            string sql = "SELECT * FROM stock.order WHERE bought=0 ORDER BY idx DESC;";
            MySqlDataAdapter adpt = new(sql, conn);
            adpt.Fill(ds, "order");
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0].Rows[0];
        }

        public void DeleteAllRows()
        {
            MySqlCommand cmd = new MySqlCommand("DELETE FROM stock.order;", conn);
            cmd.ExecuteNonQuery();
        }

        public void UpdateRow(int idx, string type, int done)
        {
            MySqlCommand cmd = new MySqlCommand("UPDATE stock.order SET " + type + "=" + done.ToString() + " WHERE idx=" + idx.ToString() + ";", conn);
            cmd.ExecuteNonQuery();
        }
    }
}
