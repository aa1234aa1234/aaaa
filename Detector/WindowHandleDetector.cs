using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod.Detector
{
    internal class WindowHandleDetector : IDetector
    {
        public WindowHandleDetector() { }

        public void Update()
        {
            IntPtr wnd;
            if((wnd=WindowEventController.FindWindow(null, TradingContext.windowNames[0])) != IntPtr.Zero )
            {
                EventSystem.GetInstance().DispatchEvent(new Event("STARTUP"), wnd, 0);
            }
            else if ((wnd = WindowEventController.FindWindow(null, TradingContext.windowNames[1])) != IntPtr.Zero)
            {
                EventSystem.GetInstance().DispatchEvent(new Event("STARTUP"), wnd, 1);
            }
        }
    }
}
