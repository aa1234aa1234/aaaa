using MySql.Data.MySqlClient;
using MySqlX.XDevAPI;
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

        public void UpdateWhitelist(string uuid, string stockcode, int bought=-1, int sold=-1)
        {
            MySqlCommand cmd;
            string sql;
            if (bought > 0 && sold > 0) sql = "UPDATE stock.whitelist SET bought" + "=" + bought + ", sold=" + sold + " WHERE  user_uuid = '%s'and stockcode='" + stockcode + "';";
            else if (bought > 0) sql = "UPDATE stock.whitelist SET bought" + "=" + bought + " WHERE  user_uuid = '%s' and stockcode='" + stockcode + "';";
            else sql = "UPDATE stock.whitelist SET sold=" + sold + " WHERE  user_uuid = '%s' and stockcode='" + stockcode + "';";
            sql = string.Format(sql, uuid);
            cmd = new(sql, conn);

            cmd.ExecuteNonQuery();
        }

        public DataRowCollection PollWhitelist(string uuid)
        {
            DataSet ds = new DataSet();
            string sql = "SELECT * FROM stock.whitelist WHERE user_uuid = '%s' and  bought=0 ORDER BY idx DESC;";
            sql = string.Format(sql, uuid);
            MySqlDataAdapter adpt = new(sql, conn);
            adpt.Fill(ds, "whitelist");
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0].Rows;
        }

        public DataRowCollection PollWhitelist2(string uuid)
        {
            DataSet ds = new DataSet();
            string sql = "SELECT * FROM stock.whitelist WHERE user_uuid = '%s' and bought=1 and sold=0 ORDER BY idx DESC;";
            sql = string.Format(sql, uuid);
            MySqlDataAdapter adpt = new(sql, conn);
            adpt.Fill(ds, "whitelist");
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0].Rows;
        }

        public DataTable PollEntireWhitelist(string uuid)
        {
            DataTable ds = new();
            string sql = "SELECT * FROM stock.whitelist WHERE user_uuid = '" + uuid + "';";
            MySqlDataAdapter adpt = new(sql, conn);
            adpt.Fill(ds);
            return ds;
        }


        public DataRowCollection GetAllPendingBuyRequest()
        {
            DataSet ds = new DataSet();
            string sql = "SELECT * FROM stock.order WHERE bought=0 ORDER BY idx DESC;";
            MySqlDataAdapter adpt = new(sql, conn);
            adpt.Fill(ds, "order");
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0].Rows;
        }

        public DataRowCollection GetAllPendingSellRequest()
        {
            DataSet ds = new DataSet();
            string sql = "SELECT * FROM stock.order WHERE sold=0 ORDER BY idx DESC;";
            MySqlDataAdapter adpt = new(sql, conn);
            adpt.Fill(ds, "order");
            if (ds.Tables[0].Rows.Count == 0) return null;
            return ds.Tables[0].Rows;
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

        public void Insert(string sql)
        {
            MySqlCommand cmd = new MySqlCommand(sql, conn);
            cmd.ExecuteNonQuery();
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


        public int PollStockPrice(string stockcode)
        {
            DataSet ds = new DataSet();
            string sql = "SELECT * FROM chatting.stockprice WHERE stockcode='" + stockcode + "' ORDER BY idx DESC;";
            MySqlDataAdapter adpt = new(sql, conn);
            adpt.Fill(ds, "price");
            return int.Parse(ds.Tables[0].Rows[0].ToString());
        }
    }
}
