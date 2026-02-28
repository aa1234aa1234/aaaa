using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod.EventListeners
{
    internal class BuyOrderListener : EventListener
    {
        private TradingEngine engine;
        public BuyOrderListener(string eventId, TradingEngine engine) : base(eventId)
        {
            this.engine = engine;
        }

        public override void OnEvent(params object[] args)
        {
            engine.SetBuyOrder(args[0].ToString(), (int)args[1], (int)args[2]);
        }
    }
}
