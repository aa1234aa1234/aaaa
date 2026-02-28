using ohmygod.Detector;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod
{
    internal class TradingEngine
    {
        private TradingContext tradingContext;
        private MySql mysql;
        private List<IDetector> detectors = new();
        public TradingEngine() {
            mysql = new();
            tradingContext = new TradingContext(TradingContext.TradingSystem.KIWOOM);
            detectors.Add(new WindowHandleDetector());
            detectors.Add(new MySqlDetector(mysql));
        }

        public void run(Form1 form)
        {
            System.Windows.Forms.Timer timer = new();
            timer.Interval = 1000;
            timer.Tick += (sender, e) =>
            {
                foreach (var p in detectors) 
                    p.Update();
            };
            timer.Start();
        }

        public void SetBuyOrder(string stockcode, int size, int price)
        {
            tradingContext.SetBuyOrder(stockcode, size, price);
        }

        public void SetWindowHandle(IntPtr handle)
        {
            tradingContext.SetWindowHandle(handle);
        }

        public IntPtr GetWindowHandle() { return tradingContext.GetWindowHandle(); }

        public void StartUp(int title)
        {
            tradingContext.StartUp(title);
        }

        public void SetWindow(int a) {  tradingContext.SetWindow(a); }

        public int GetWindow() {  return tradingContext.GetWindow(); }

    }
}
