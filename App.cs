using ohmygod.EventListeners;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod
{
    internal class App
    {
        private TradingEngine engine;
        private StartUpListener StartUpListener;
        private BuyOrderListener BuyOrderListener;

        public App() {
            engine = new();
            StartUpListener = new("STARTUP", engine);
            BuyOrderListener = new("SETBUYORDER", engine);
        }

        public void run(Form1 form)
        {
            engine.run(form);
        }
    }
}
