
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
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
        private Dictionary<string, int> stockVolume = new();
        private Stack<Order> buyOrders = new(), sellOffers = new();
        private Stack<string> buyOrderID = new(), sellOfferID = new();
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


        private void BuyOrder(string stockcode, int size, int price)
        {
            Point[] offset = { new Point(80, -20), new Point(0, 20), new Point(0, 60), new Point(0, 100), new Point(0, 160) };
            string[] input = { "690201", stockcode, size.ToString(), price.ToString() };
            Bitmap buy = new Bitmap(@"../../../images/kiwoom/buy1.png");
            Rectangle rect;
            rect = ImageFinder.FindImage(@"../../../images/kiwoom/e3.png", screenBitmap).MinBy(p => p.X);
            if (rect == Rectangle.Empty) return;
            rect.X -= 150;
            rect.Y += 23;
            windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
            for (int i = 0; i < offset.Length; i++)
            {
                if (i < input.Length && input[i] == "-1") continue;
                windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2), offset[i]);
                if (i < input.Length) windowController.typeString(input[i]);
                Thread.Sleep(50);
            }
            windowController.Click(new Point(550, 419));
            
        }
        public void SetBuyOrder(string stockcode, int size, int price, int resellprice = -1, int idx = -1)
        {
            //ClearScreen();
            //windowController.Click(new Point(110, 36));
            //Thread.Sleep(500);
            //Rectangle a = ImageFinder.FindSingleImage(@"../../../images/kiwoom/ae.png", screenBitmap);
            //if (a != Rectangle.Empty)
            //{
            //    windowController.Click(new Point(a.X + a.Width / 2, a.Y + a.Height / 2));
            //}
            OpenWindow("4989");
            Thread.Sleep(2000);
            
            switch (tradingSystem)
            {
                case TradingSystem.KIWOOM:
                    //Point[] offset = { new Point(80, -20), new Point(0, 20), new Point(0, 60), new Point(0, 100), new Point(0, 160) };
                    //string[] input = { "690201", stockcode, size.ToString(), price.ToString() };
                    //Bitmap buy = new Bitmap(@"../../../images/kiwoom/buy1.png");
                    //Rectangle rect;
                    //rect = ImageFinder.FindImage(@"../../../images/kiwoom/e3.png", screenBitmap).MinBy(p => p.X);
                    //if (rect == Rectangle.Empty) return;
                    //rect.X -= 150;
                    //rect.Y += 23;
                    //windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
                    //for (int i = 0; i < offset.Length; i++)
                    //{
                    //    if (i < input.Length && input[i] == "-1") continue;
                    //    windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2), offset[i]);
                    //    if (i < input.Length) windowController.typeString(input[i]);
                    //    Thread.Sleep(50);
                    //}
                    //windowController.Click(new Point(550, 419));
                    BuyOrder(stockcode, size, price);
                    MySql.GetInstance().UpdateRow(idx, "bought", 1);
                    if (!stockVolume.ContainsKey(stockcode)) stockVolume.Add(stockcode, size);
                    else stockVolume[stockcode] += size;
                    Thread.Sleep(100);
                    windowController.typeString("#VK_ESCAPE#");
                    buyOrders.Push(new Order(stockcode,size,price,resellprice));
                    break;
                case TradingSystem.IMERITZ:
                    break;
            }
        }

        private void SellOffer(string stockcode, int size, int price)
        {
            //Point[] offset = { new Point(350, 30), new Point(270, 70), new Point(270, 117), new Point(270, 150), new Point(270, 205) };
            Point[] offset = { new Point(80, -20), new Point(0, 20), new Point(0, 60), new Point(0, 100), new Point(0, 160) };
            string[] input = { "690201", stockcode, size.ToString(), price.ToString() };
            Bitmap buy = new Bitmap(@"../../../images/kiwoom/buy.png");
            Rectangle rect;
            rect = ImageFinder.FindImage(@"../../../images/kiwoom/e3.png", screenBitmap).MinBy(p => p.X);
            if (rect == Rectangle.Empty) return;
            rect.X -= 150;
            rect.Y += 23;
            windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2), new Point(70, 0));
            for (int i = 0; i < offset.Length; i++)
            {
                if (i < input.Length && input[i] == "-1") continue;
                windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2), offset[i]);
                if (i < input.Length) windowController.typeString(input[i]);
                Thread.Sleep(50);
            }
            windowController.Click(new Point(550, 419));
        }

        public void SetSellOffer(Order order)
        {
            //ClearScreen();
            //windowController.Click(new Point(110, 36));
            //Thread.Sleep(500);
            //Rectangle a = ImageFinder.FindSingleImage(@"../../../images/kiwoom/ae.png", screenBitmap);
            //if (a != Rectangle.Empty)
            //{
            //    windowController.Click(new Point(a.X + a.Width / 2, a.Y + a.Height / 2));
            //}
            OpenWindow("4989");
            Thread.Sleep(2000);
            switch (tradingSystem)
            {
                case TradingSystem.KIWOOM:
                    //Point[] offset = { new Point(350, 30), new Point(270, 70), new Point(270, 117), new Point(270, 150), new Point(270, 205) };
                    //Point[] offset = { new Point(80, -20), new Point(0, 20), new Point(0, 60), new Point(0, 100), new Point(0, 160) };
                    //string[] input = { "690201", stockcode, stockVolume[stockcode].ToString(), price.ToString() };
                    //Bitmap buy = new Bitmap(@"../../../images/kiwoom/buy.png");
                    //Rectangle rect;
                    //rect = ImageFinder.FindImage(@"../../../images/kiwoom/e3.png", screenBitmap).MinBy(p => p.X);
                    //if (rect == Rectangle.Empty) return;
                    //rect.X -= 150;
                    //rect.Y += 23;
                    //windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2), new Point(70, 0));
                    //for (int i = 0; i < offset.Length; i++)
                    //{
                    //    windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2), offset[i]);
                    //    if (i < input.Length) windowController.typeString(input[i]);
                    //    Thread.Sleep(50);
                    //}
                    //windowController.Click(new Point(550, 419));
                    SellOffer(order.stockcode, order.size, order.price);
                    Thread.Sleep(100);
                    windowController.typeString("#VK_ESCAPE#");
                    sellOffers.Push(order);
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

        public void OpenWindow(string windowCode)
        {
            windowController.Click(new Point(57, 88));
            windowController.typeString(windowCode, true);
        }


        private Bitmap ocr(Vector2 start, Vector2 size)
        {
            Vector2 screenSize = size;
            Bitmap screen = new Bitmap((int)size.X, (int)size.Y);
            using (Graphics graphic = Graphics.FromImage(screen))
            {
                graphic.CopyFromScreen(new Point((int)start.X, (int)start.Y), Point.Empty, screen.Size);
            }
            return screen;
        }

        private bool flag = false;

        private void CheckAccount(Rectangle rect)
        {
            //rect.X += 2;
            //rect.Y += 8;
            //var orders = new Stack<string>();
            //var engine = new TesseractEngine(@"./tessdata", "kor+eng", EngineMode.Default);
            //try
            //{
            //    var a = ImageFinder.FindImage(@"../../../images/kiwoom/dd.png", screenBitmap).Count;
            //    for (int i = 0; i < a - 1; i++)
            //    {
            //        Bitmap bitmap = ocr(new Vector2(rect.X, rect.Y + 3), new Vector2(39, 16));
            //        OpenCvSharp.Mat src = OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap), dst = new OpenCvSharp.Mat();
            //        OpenCvSharp.Cv2.Resize(src, dst, new OpenCvSharp.Size(68, 28));
            //        Pix pix = PixConverter.ToPix(OpenCvSharp.Extensions.BitmapConverter.ToBitmap(dst));
            //        var result = engine.Process(pix);
            //        Console.WriteLine(result.GetText() + " fjewlakfjwlaejflawef");
            //        orders.Push(result.GetText());

            //        result.Dispose();
            //        pix.Dispose();
            //        rect.Y += 18;
            //    }
            //}
        }

        int checkedbox = -1;

        public void CheckBuyOrder()
        {
            
            Rectangle rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/da.png", screenBitmap);
            if (rect != Rectangle.Empty)
            {
                if (buyOrders.Count == 0) return;
                if (checkedbox != 0)
                {
                    windowController.Click(new Point(rect.X + 180, rect.Y + 80));
                    checkedbox = 0;
                }
                rect.X -= 41;
                rect.Y += 116;
                var orders = new Stack<string>();
                var engine = new TesseractEngine(@"./tessdata", "kor+eng", EngineMode.Default);
                try
                {
                    var a = ImageFinder.FindImage(@"../../../images/kiwoom/dd.png", screenBitmap).Count;
                    for (int i = 0; i < a - 1; i++)
                    {
                        Bitmap bitmap = ocr(new Vector2(rect.X, rect.Y + 3), new Vector2(39, 16));
                        OpenCvSharp.Mat src = OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap), dst = new OpenCvSharp.Mat();
                        OpenCvSharp.Cv2.Resize(src, dst, new OpenCvSharp.Size(68, 28));
                        Pix pix = PixConverter.ToPix(OpenCvSharp.Extensions.BitmapConverter.ToBitmap(dst));
                        var result = engine.Process(pix);
                        Console.WriteLine(result.GetText() + " fjewlakfjwlaejflawef");
                        orders.Push(result.GetText());

                        result.Dispose();
                        pix.Dispose();
                        rect.Y += 18;
                    }
                    engine.Dispose();
                }
                catch (Exception e)
                {
                    return;
                }
                
                if(orders.Count < buyOrders.Count && buyOrders.Count != 0)
                {
                    List<int> idx = new();
                    for (int i = buyOrderID.Count - 1; i >= 0; i--)
                    {
                        idx.Add(i);
                    }
                    if (idx.Count == 0) idx.Add(0);
                    OpenWindow("4989");
                    Thread.Sleep(2000);
                    var buyorders = buyOrders.ToList();
                    foreach (var p in idx)
                    {
                        SellOffer(buyorders[p].stockcode, buyorders[p].size, buyorders[p].resellprice);
                        sellOffers.Push(new Order(buyorders[p].stockcode, buyorders[p].size, buyorders[p].resellprice, buyorders[p].idx));
                        buyorders.RemoveAt(p);
                        buyOrders = new Stack<Order>(buyorders);
                        Thread.Sleep(100);
                    }
                    windowController.typeString("#VK_ESCAPE#");
                    return;
                }
                if (buyOrderID.Count == 0) { buyOrderID.Clear(); buyOrderID = orders; }
                
                else
                {
                    List<int> idx = new();
                    for (int i = buyOrderID.Count - 1; i>=0; i--)
                    {
                        if (!orders.Contains(buyOrderID.ToList()[i]))
                        {
                            idx.Add(i);
                        }
                    }
                    buyOrderID = orders;
                    if (idx.Count == 0) return;
                    OpenWindow("4989");
                    Thread.Sleep(2000);
                    var buyorders = buyOrders.ToList();
                    foreach (var p in idx)
                    {
                        SellOffer(buyorders[p].stockcode, buyorders[p].size, buyorders[p].resellprice);
                        sellOffers.Push(new Order(buyorders[p].stockcode, buyorders[p].size, buyorders[p].resellprice, buyorders[p].idx));
                        buyorders.RemoveAt(p);
                        Thread.Sleep(100);
                    }
                    buyOrders = new Stack<Order>(buyorders);
                    windowController.typeString("#VK_ESCAPE#");
                    

                }
            }
            else if(!flag) { OpenWindow("0341"); flag = true; }
        }

        public void CheckSellOffers()
        {
            
            Rectangle rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/da.png", screenBitmap);
            if (rect != Rectangle.Empty)
            {
                if (sellOffers.Count == 0) return;
                if (checkedbox != 1)
                {
                    checkedbox = 1;
                    windowController.Click(new Point(rect.X + 225, rect.Y + 80));
                }
                rect.X -= 41;
                rect.Y += 116;
                var orders = new Stack<string>();
                var engine = new TesseractEngine(@"./tessdata", "kor+eng", EngineMode.Default);
                try
                {
                    var a = ImageFinder.FindImage(@"../../../images/kiwoom/dd.png", screenBitmap).Count;
                    for (int i = 0; i < a - 1; i++)
                    {
                        Bitmap bitmap = ocr(new Vector2(rect.X, rect.Y + 3), new Vector2(39, 16));
                        OpenCvSharp.Mat src = OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap), dst = new OpenCvSharp.Mat();
                        OpenCvSharp.Cv2.Resize(src, dst, new OpenCvSharp.Size(68, 28));
                        Pix pix = PixConverter.ToPix(OpenCvSharp.Extensions.BitmapConverter.ToBitmap(dst));
                        var result = engine.Process(pix);
                        Console.WriteLine(result.GetText() + " fjewlakfjwlaejflawef");
                        orders.Push(result.GetText());

                        result.Dispose();
                        pix.Dispose();
                        rect.Y += 18;
                    }
                    engine.Dispose();
                }
                catch (Exception e)
                {
                    return;
                }

                if (orders.Count < sellOffers.Count && sellOffers.Count != 0)
                {
                    List<int> idx = new();
                    for (int i = sellOffers.Count - 1; i >= 0; i--)
                    {
                        idx.Add(i);
                    }
                    if (idx.Count == 0) idx.Add(0);
                    var selloffers = sellOffers.ToList();
                    foreach (var p in idx)
                    {
                        MySql.GetInstance().UpdateRow(selloffers[p].idx, "sold", 1);
                        selloffers.RemoveAt(p);
                        
                        Thread.Sleep(100);
                    }
                    sellOffers = new Stack<Order>(selloffers);
                    //windowController.typeString("#VK_ESCAPE#");
                    return;
                }
                if (sellOfferID.Count == 0) { sellOfferID.Clear(); sellOfferID = orders; }

                else
                {
                    List<int> idx = new();
                    for (int i = sellOfferID.Count - 1; i >= 0; i--)
                    {
                        if (!orders.Contains(sellOfferID.ToList()[i]))
                        {
                            idx.Add(i);
                        }
                    }
                    sellOfferID = orders;
                    if (idx.Count == 0) return;
                    var selloffers = sellOffers.ToList();
                    foreach (var p in idx)
                    {
                        MySql.GetInstance().UpdateRow(selloffers[p].idx, "sold", 1);
                        selloffers.RemoveAt(p);
                        
                        Thread.Sleep(100);
                    }
                    sellOffers = new Stack<Order>(selloffers);
                    //windowController.typeString("#VK_ESCAPE#");


                }
            }
            else if (!flag) { OpenWindow("0341"); flag = true; }
        }

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
                        else windowController.typeString("#VK_ESCAPE#");
                    }
                    else
                    {
                        rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/ee.png", screenBitmap);
                        if (rect != Rectangle.Empty) windowController.Click(new Point(rect.X, rect.Y));
                        else windowController.typeString("#VK_ESCAPE#");
                    }
                    ClearScreen();
                    List<Rectangle>? a=new();
                    do
                    {
                        Console.WriteLine(a.Count);
                        if(a.Count >= 2) Console.WriteLine(a[0] + "\n" + a[1]);
                        Thread.Sleep(50);
                    } while ((a=ImageFinder.FindImage(@"../../../images/kiwoom/d.png", screenBitmap)) != null && a.Count > 1);
                    OpenWindow("0345");
                    Thread.Sleep(1000);
                    rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/account.png", screenBitmap);
                    windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
                    windowController.typeString("#VK_ESCAPE#");
                    break;
            }
        }
    }
}
