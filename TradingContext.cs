
using FlaUI.Core.Tools;
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
        private List<KeyValuePair<string, Order>> stocks = new();
        private Dictionary<string, int> stockVolume = new();
        private Stack<Order> buyOrders = new(), sellOffers = new();
        private Stack<string> buyOrderID = new(), sellOfferID = new();
        private Dictionary<string, int> stockPrice = new();
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

        public void UpdateStock()
        {
            System.Data.DataRowCollection whitelist = MySql.GetInstance().PollWhitelist();
            foreach(System.Data.DataRow p in whitelist)
            {
                if (buyOrders.Any(a => (a.stockcode == p["stockcode"].ToString())))
                {
                    continue;
                }
                //BuyOrder(p["stockcode"].ToString(), 1, int.Parse(p["price"].ToString()));
                //buyOrders.Push(new Order(p["stockcode"].ToString(), 1, int.Parse(p["price"].ToString())));
                int price = int.Parse(p["price"].ToString());
                int resell = int.Parse(p["resellprice"].ToString());
                MySql.GetInstance().Insert(string.Format("INSERT INTO stock.order VALUES (0, 'BUY', '%s', 1, %d, %d, 0, 0);", p["stockcode"].ToString(), price, resell));
            }
            whitelist = MySql.GetInstance().PollWhitelist2();
            foreach (System.Data.DataRow p in whitelist)
            {
                if (sellOffers.Any(a => (a.stockcode == p["stockcode"].ToString())))
                {
                    continue;
                }
                //BuyOrder(p["stockcode"].ToString(), 1, int.Parse(p["price"].ToString()));
                //buyOrders.Push(new Order(p["stockcode"].ToString(), 1, int.Parse(p["price"].ToString())));
                int price = int.Parse(p["price"].ToString());
                int resell = int.Parse(p["resellprice"].ToString());
                sellOffers.Push(new Order(p["stockcode"].ToString(), 1, resell));
                //MySql.GetInstance().Insert(string.Format("INSERT INTO stock.order VALUES (0, 'BUY', '%s', 1, %d, %d, 0, 0);", p["stockcode"].ToString(), price, resell));
            }
        }

        public void UpdateStockPrice(MySql sql)
        {
            foreach(var p in stockPrice)
            {
                stockPrice[p.Key] = sql.PollStockPrice(p.Key);
            }
        }

        public void Sell()
        {
            foreach(var p in sellOffers)
            {
                if (p.isOrderUp == 1) continue;
                if (stockPrice[p.stockcode] > p.price * 1.1) SellOffer(p.stockcode, p.size, stockPrice[p.stockcode]);
                else if (stockPrice[p.stockcode] < p.price * 1.03) SellOffer(p.stockcode, p.size, stockPrice[p.stockcode]);
            }
        }

        private void BuyOrder(string stockcode, int size, int price)
        {
            bool a = false;
            //foreach (var p in stocks) { if (p.Key == stockcode) { a = true; break; } }
            //if (!a) { return; }
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
                    Console.WriteLine("buyorder idx: " + idx);
                    buyOrders.Push(new Order(stockcode, size, price, resellprice, idx, 1));
                    if (!stockVolume.ContainsKey(stockcode)) stockVolume.Add(stockcode, size);
                    else stockVolume[stockcode] += size;
                    Thread.Sleep(100);
                    windowController.typeString("#VK_ESCAPE#");
                    
                    break;
                case TradingSystem.IMERITZ:
                    break;
            }
        }

        private void SellOffer(string stockcode, int size, int price, Order order = null)
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
            order.isOrderUp = 1;
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

        public void CheckBuyOrder(bool onBoot = false)
        {
            
            Rectangle rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/da.png", screenBitmap);
            if (rect != Rectangle.Empty)
            {
                if (buyOrders.Count == 0 && !onBoot) return;
                if (checkedbox != 0)
                {
                    windowController.Click(new Point(rect.X + 180, rect.Y + 80));
                    checkedbox = 0;
                }
                //if(stockcode != "")
                //{
                //    windowController.Click(new Point(rect.X - 5, rect.Y+80));
                //    windowController.Click(new Point(rect.X +40, rect.Y + 80));
                //    windowController.typeString(stockcode);
                //    var a = ImageFinder.FindImage(@"../../../images/kiwoom/dd.png", screenBitmap).Count;
                //    if(a > 1) {
                //        buyOrders.Push(new Order(stockcode, 1, 0));
                //        windowController.typeString("#VK_ESCAPE#");
                //        return;
                //    }
                //}
                //else
                //{
                //    windowController.Click(new Point(rect.X - 30, rect.Y + 80));
                //}
                rect.X -= 41;
                rect.Y += 116;
                if(onBoot)
                {
                    windowController.MoveMouse(rect.X - 1, rect.Y + 15);
                    windowController.MoveRelative(1, 1);
                }
                
                
                var orders = new Stack<string>();
                var engine = new TesseractEngine(@"./tessdata", "kor+eng", EngineMode.Default);
                try
                {
                    var a = ImageFinder.FindImage(@"../../../images/kiwoom/dd.png", screenBitmap).Count;
                    for (int i = 0; i < a - 1; i++)
                    {
                        //Bitmap bitmap = ocr(new Vector2(rect.X, rect.Y + 3), new Vector2(39, 16));
                        //OpenCvSharp.Mat src = OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap), dst = new OpenCvSharp.Mat();
                        //OpenCvSharp.Cv2.Resize(src, dst, new OpenCvSharp.Size(68, 28));
                        //Pix pix = PixConverter.ToPix(OpenCvSharp.Extensions.BitmapConverter.ToBitmap(dst));
                        //var result = engine.Process(pix);
                        //Console.WriteLine(result.GetText() + " fjewlakfjwlaejflawef");
                        //orders.Push(result.GetText());

                        //result.Dispose();
                        //pix.Dispose();
                        orders.Push(ImageReader.GetInstance().ReadBitmap(engine, new Vector2(rect.X, rect.Y + 3), new Vector2(39, 16)));
                        if (onBoot)
                        {
                            windowController.MoveRelative(40 * (i > 0 ? 0 : 1), 18 * i);
                            windowController.MoveRelative(3, 3);
                            windowController.MoveRelative(-3, -3);
                            Thread.Sleep(500);
                            Point pos = windowController.GetMousePos();
                            Console.WriteLine(pos + " position");
                            var code=  ImageReader.GetInstance().ReadBitmap(engine, new Vector2(pos.X + 33, pos.Y + 19), new Vector2(38, 17));
                            Console.WriteLine(ImageReader.GetInstance().ReadBitmap(engine, new Vector2(pos.X + 268, pos.Y - 3), new Vector2(68, 18)));
                            buyOrders.Push(new Order(code, 1, -1, -1, -1, 1));
                        }
                        rect.Y += 18;
                    }
                    engine.Dispose();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
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
                        //SellOffer(buyorders[p].stockcode, buyorders[p].size, buyorders[p].resellprice);
                        sellOffers.Push(new Order(buyorders[p].stockcode, buyorders[p].size, buyorders[p].price, buyorders[p].resellprice, buyorders[p].idx));
                        MySql.GetInstance().UpdateWhitelist(buyorders[p].stockcode, 1);
                        Console.WriteLine(buyorders[p].idx + " " + buyorders[p].stockcode);
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
                        stocks.Add(new KeyValuePair<string, Order>(buyorders[p].stockcode, new Order(buyorders[p].stockcode, buyorders[p].size, buyorders[p].resellprice, buyorders[p].idx)));
                        stockPrice.Add(buyorders[p].stockcode, buyorders[p].price);
                        MySql.GetInstance().UpdateWhitelist(buyorders[p].stockcode, 1);
                        //SellOffer(buyorders[p].stockcode, buyorders[p].size, buyorders[p].resellprice);
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

        public void CheckSellOffers(bool onBoot=false)
        {
            
            Rectangle rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/da.png", screenBitmap);
            if (rect != Rectangle.Empty)
            {
                if (sellOffers.Count == 0 && !onBoot) return;
                if (checkedbox != 1)
                {
                    checkedbox = 1;
                    windowController.Click(new Point(rect.X + 225, rect.Y + 80));
                }
                rect.X -= 41;
                rect.Y += 116;
                if(onBoot)
                {
                    windowController.MoveMouse(rect.X - 1, rect.Y + 13);
                    windowController.MoveRelative(1, 1);
                }
                
                var orders = new Stack<string>();
                var engine = new TesseractEngine(@"./tessdata", "kor+eng", EngineMode.Default);
                Console.WriteLine("debug");
                try
                {
                    var a = ImageFinder.FindImage(@"../../../images/kiwoom/dd.png", screenBitmap).Count;
                    for (int i = 0; i < a - 1; i++)
                    {
                        //Bitmap bitmap = ocr(new Vector2(rect.X, rect.Y + 3), new Vector2(39, 16));
                        //OpenCvSharp.Mat src = OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap), dst = new OpenCvSharp.Mat();
                        //OpenCvSharp.Cv2.Resize(src, dst, new OpenCvSharp.Size(68, 28));
                        //Pix pix = PixConverter.ToPix(OpenCvSharp.Extensions.BitmapConverter.ToBitmap(dst));
                        //var result = engine.Process(pix);
                        //Console.WriteLine(result.GetText() + " fjewlakfjwlaejflawef");
                        //orders.Push(result.GetText());

                        //result.Dispose();
                        //pix.Dispose();
                        Console.WriteLine("debug2");
                        orders.Push(ImageReader.GetInstance().ReadBitmap(engine, new Vector2(rect.X, rect.Y + 3), new Vector2(39, 16)));
                        Console.WriteLine("orders: " + orders.Count);
                        if (onBoot)
                        {
                            windowController.MoveRelative(40 * (i > 0 ? 0 : 1), 18 * i);
                            windowController.MoveRelative(3, 3);
                            windowController.MoveRelative(-3, -3);
                            Thread.Sleep(500);
                            Point pos = windowController.GetMousePos();
                            Console.WriteLine(pos + " position");
                            var code = ImageReader.GetInstance().ReadBitmap(engine, new Vector2(pos.X + 33, pos.Y + 19), new Vector2(38, 17));
                            Console.WriteLine(ImageReader.GetInstance().ReadBitmap(engine, new Vector2(pos.X + 268, pos.Y - 3), new Vector2(68, 18)));
                            sellOffers.Push(new Order(code, 1, -1));
                            Console.WriteLine(sellOffers.Count);
                            
                        }
                        rect.Y += 18;
                    }
                    engine.Dispose();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
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
                        Console.WriteLine("idx: " + selloffers[p].idx);
                        selloffers.RemoveAt(p);
                        MySql.GetInstance().UpdateWhitelist(selloffers[p].stockcode, -1, 1);

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
                        MySql.GetInstance().UpdateWhitelist(selloffers[p].stockcode, -1, 1);
                        Console.WriteLine("idx: " + selloffers[p].idx);
                        selloffers.RemoveAt(p);
                        
                        Thread.Sleep(100);
                    }
                    sellOffers = new Stack<Order>(selloffers);
                    //windowController.typeString("#VK_ESCAPE#");


                }
            }
            else if (!flag) { OpenWindow("0341"); flag = true; }
        }

        public void CheckAccount()
        {
            OpenWindow("0345");
            Thread.Sleep(1000);
            //rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/account.png", screenBitmap);
            //windowController.Click(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
            //windowController.typeString("#VK_ESCAPE#");
            Rectangle rect = ImageFinder.FindSingleImage(@"../../../images/kiwoom/da.png", screenBitmap);
            if (rect != Rectangle.Empty)
            {
                rect.X -= 54;
                rect.Y += 116 + 16;
                windowController.MoveMouse(rect.X-1, rect.Y-1);
                windowController.MoveRelative(1, 1);
                var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default);
                try
                {
                    var a = ImageFinder.FindImage(@"../../../images/kiwoom/dd.png", screenBitmap).Count;
                    for (int i = 0; i < a - 1; i++)
                    {
                        //Bitmap bitmap = ocr(new Vector2(rect.X, rect.Y + 3), new Vector2(39, 16));
                        //OpenCvSharp.Mat src = OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap), dst = new OpenCvSharp.Mat();
                        //OpenCvSharp.Cv2.Resize(src, dst, new OpenCvSharp.Size(68, 28));
                        //Pix pix = PixConverter.ToPix(OpenCvSharp.Extensions.BitmapConverter.ToBitmap(dst));
                        //var result = engine.Process(pix);
                        //Console.WriteLine(result.GetText() + " fjewlakfjwlaejflawef");
                        //orders.Push(result.GetText());

                        //result.Dispose();
                        //pix.Dispose();
                        windowController.MoveRelative(40 * (i > 0 ? 0 : 1), 18*i);
                        windowController.MoveRelative(3, 3);
                        Thread.Sleep(100);
                        windowController.MoveRelative(-3, -3);
                        Thread.Sleep(5000);
                        Point pos = windowController.GetMousePos();
                        Console.WriteLine(pos + " position");
                        Console.WriteLine(ImageReader.GetInstance().ReadBitmap(engine, new Vector2(pos.X + 33, pos.Y + 19), new Vector2(38, 17)));
                        Console.WriteLine(ImageReader.GetInstance().ReadBitmap(engine, new Vector2(pos.X + 268, pos.Y - 3), new Vector2(68, 18)));
                        rect.Y += 18;
                    }
                    engine.Dispose();
                }
                catch (Exception e)
                {
                    return;
                }
            }
            windowController.typeString("#VK_ESCAPE#");
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
                    CheckAccount();
                    OpenWindow("0341");
                    break;
            }
        }
    }
}
