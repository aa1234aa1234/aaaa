using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod.Detector
{
    internal class MySqlDetector : IDetector
    {
        private MySql sql;

        public MySqlDetector(MySql sql)
        {
            this.sql = sql;
        }

        public void Update()
        {
            System.Data.DataRow row;
            if((row=sql.PollLatestRequest()) != null)
            {
                switch(row["type"].ToString())
                {
                    case "BUY":
                        EventSystem.GetInstance().DispatchEvent(new Event("SETBUYORDER"), row["code"], row["vol"], row["price"]);
                        break;
                    case "SELL":
                        break;
                }
            }
        }
    }
}
