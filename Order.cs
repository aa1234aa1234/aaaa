using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod
{
    class Order
    {
        public string stockcode;
        public int size;
        public int price, resellprice;
        public int idx;
        public int isOrderUp=0;

        public Order(string stockcode, int size, int price, int resellprice = -1, int idx = -1, int order = -1) { 
            this.stockcode = stockcode; this.size = size; this.price = price; this.resellprice = (resellprice == -1 ? price : resellprice); 
            this.idx = idx;
            this.isOrderUp = order;
        }

        public static bool operator==(Order left, string stockcode)
        {
            return left.stockcode == stockcode;
        }

        public static bool operator !=(Order left, string stockcode)
        {
            return left.stockcode != stockcode;
        }

        public static bool operator ==(Order left, Order right)
        {
            return left.stockcode == right.stockcode;
        }

        public static bool operator !=(Order left, Order right)
        {
            return left.stockcode != right.stockcode;
        }
    }
}
