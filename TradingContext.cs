using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tesseract;

namespace ohmygod
{
    internal class TradingContext
    {
        public enum TradingSystem
        {
            KIWOOM,
            IMERITZ
        }

        private WindowEventController windowController;
        private TesseractEngine tesseractEngine;
        private TradingSystem tradingSystem;
        private Bitmap screenBitmap;
        public static string[] windowNames;
        private int window;
        private IntPtr wnd;

        public TradingContext(TradingSystem tradingsystem) {
            windowController = new();
            tradingSystem = tradingsystem;
            screenBitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            tesseractEngine = new TesseractEngine(@"./tessdata", "kor", EngineMode.Default);
            Graphics.FromImage(screenBitmap).CopyFromScreen(new Point(0,0),Point.Empty,new Size(screenBitmap.Width, screenBitmap.Height));
            switch(tradingSystem)
            {
                case TradingSystem.KIWOOM:
                    windowNames = new string[2]{ "영웅문4 Login", "영웅문4"};
                    break;
            }
        }

        ~TradingContext()
        {
            tesseractEngine.Dispose();
        }

        public void SetBuyOrder(string stockcode, int size, int price)
        {
            ClearScreen();
            windowController.Click(new Point(110, 36));
            Thread.Sleep(100);
            Rectangle a = ImageFinder.FindSingleImage(@"../../../images/kiwoom/ae.png", screenBitmap);
            if (a != Rectangle.Empty)
            {
                windowController.Click(new Point(a.X + a.Width / 2, a.Y + a.Height / 2));
            }
            Thread.Sleep(500);
            switch (tradingSystem)
            {
                case TradingSystem.KIWOOM:
                    Point[] offset = { new Point(80, 20), new Point(0, 20), new Point(0, 60), new Point(0, 100), new Point(0, 160) };
                    string[] input = { "690201", stockcode, size.ToString(), price.ToString() };
                    Bitmap buy = new Bitmap(@"../../../images/kiwoom/buy1.png");
                    Rectangle rect;
                    rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/eee.png", screenBitmap);
                    if (rect == Rectangle.Empty) return;
                    windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
                    for (int i = 0; i < offset.Length; i++)
                    {
                        windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2), offset[i]);
                        if (i < input.Length) windowController.typeString(input[i]);
                        Thread.Sleep(50);
                    }
                    windowController.Click(new Point(550, 419));
                    break;
                case TradingSystem.IMERITZ:
                    break;
            }
        }

        public void SetSellOffer(string stockcode, int size, int price)
        {
            ClearScreen();
            windowController.Click(new Point(110, 36));
            Thread.Sleep(100);
            Rectangle a = ImageFinder.FindSingleImage(@"../../../images/kiwoom/ae.png", screenBitmap);
            if (a != Rectangle.Empty)
            {
                windowController.Click(new Point(a.X + a.Width / 2, a.Y + a.Height / 2));
            }
            Thread.Sleep(500);
            switch (tradingSystem)
            {
                case TradingSystem.KIWOOM:
                    Point[] offset = { new Point(350, 30), new Point(270, 70), new Point(270, 117), new Point(270, 150), new Point(270, 205) };
                    string[] input = { "690201", stockcode, size.ToString(), price.ToString() };
                    Bitmap buy = new Bitmap(@"../../../images/kiwoom/buy.png");
                    Rectangle rect;
                    rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/eee.png", screenBitmap);
                    if (rect == Rectangle.Empty) return;
                    windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2), new Point(70, 0));
                    for (int i = 0; i < offset.Length; i++)
                    {
                        windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2), offset[i]);
                        if (i < input.Length) windowController.typeString(input[i]);
                        Thread.Sleep(50);
                    }
                    windowController.Click(new Point(550, 419));
                    break;
                case TradingSystem.IMERITZ:
                    break;
            }
        }

        private void ClearScreen()
        {
            //foreach(var p in ImageFinder.FindImage(@"../../../images/kiwoom/a.png", screenBitmap))
            //{
            //    windowController.Click(new Point(p.X + p.Width / 2, p.Y + p.Height / 2));
            //}
            Rectangle rect;
            windowController.Click(new Point(350, 58));
            Thread.Sleep(500);
            rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/purge.png", screenBitmap);
            if (rect != Rectangle.Empty) windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
            Thread.Sleep(50);
            windowController.typeString(" ");
        }

        public void SetWindowHandle(IntPtr hwnd) { wnd = hwnd; }

        public IntPtr GetWindowHandle() { return wnd; }

        public int GetWindow() { return window; }

        public void SetWindow(int a) { window = a; }

        public void StartUp(int windowTitle)
        { 
            windowController.Click(new Point(1, 0));
            windowController.SetWindowPos(wnd, 0, 0, 1200, 800);
            windowController.FocusWindow(wnd);
            switch (windowTitle)
            {
                case 1:
                    Rectangle rect;
                    rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/b.png", screenBitmap);
                    if (rect != Rectangle.Empty)
                    {
                        rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/e.png", screenBitmap);
                        if (rect != Rectangle.Empty) windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
                        else windowController.typeString(" ");
                    }
                    else
                    {
                        rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/ee.png", screenBitmap);
                        if (rect != Rectangle.Empty) windowController.Click(new Point(rect.X, rect.Y));
                        else windowController.Click(new Point(600, 400));
                            windowController.typeString(" ");
                    }
                    ClearScreen();
                    List<Rectangle>? a=new();
                    do
                    {
                        Console.WriteLine(a.Count);
                        if(a.Count >= 2) Console.WriteLine(a[0] + "\n" + a[1]);
                        Thread.Sleep(50);
                    } while ((a=ImageFinder.FindImage(@"../../../images/kiwoom/d.png", screenBitmap)) != null && a.Count > 1);
                    windowController.Click(new Point(57, 88));
                    windowController.typeString("0345", true);
                    Thread.Sleep(1000);
                    rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/account.png", screenBitmap);
                    windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
                    break;
            }
        }
    }
}
