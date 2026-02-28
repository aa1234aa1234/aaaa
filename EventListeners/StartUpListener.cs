using FlaUI.Core.Identifiers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod.EventListeners
{
    internal class StartUpListener : EventListener
    {
        private TradingEngine engine;
        public StartUpListener(string eventId, TradingEngine engine) : base(eventId)
        {
            this.engine = engine;
        }

        public override void OnEvent(params object[] args)
        {
            if(engine.GetWindowHandle() != (IntPtr)args[0] && TradingContext.windowNames[(int)args[1]] != TradingContext.windowNames[engine.GetWindow()])
            {
                Console.WriteLine("fewajfejfjaejfwjeajfjefjwaejf");
                Thread.Sleep(1000);
                engine.SetWindowHandle((IntPtr)args[0]);
                engine.SetWindow((int)args[1]);
                engine.StartUp((int)args[1]);
            }
        }
    }
}
