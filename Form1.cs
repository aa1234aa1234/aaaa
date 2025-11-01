using FlaUI.Core;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using Interop.UIAutomationClient;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Windows.Forms.Automation;
using Tesseract;

namespace ohmygod
{

    public partial class Form1 : Form
    {

        [DllImport("user32.dll")]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int cButtons, int dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        public static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);

        [DllImport("user32.dll")]
        internal static extern uint SendInput(uint nInputs, [MarshalAs(UnmanagedType.LPArray), In] INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte vk, byte scan, int flags, ref int extrainfo);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,   // HWND_TOP, HWND_BOTTOM, HWND_TOPMOST …
        int X,
        int Y,
        int cx,
        int cy,
        uint uFlags);

        const byte AltKey = 18; const int KEYUP = 0x0002;

        internal struct INPUT
        {
            public UInt32 Type;
            public MOUSEKEYBDHARDWAREINPUT Data;
        }
        [StructLayout(LayoutKind.Explicit)]
        internal struct MOUSEKEYBDHARDWAREINPUT
        {
            [FieldOffset(0)] public MOUSEINPUT Mouse;

            // 0 offset – keyboard (adds new fields)
            [FieldOffset(0)] public KEYBDINPUT Keyboard;

            // 0 offset – hardware (rarely used)
            [FieldOffset(0)] public HARDWAREINPUT Hardware;
        }

        internal struct MOUSEINPUT
        {
            public Int32 X;
            public Int32 Y;
            public UInt32 MouseData;
            public UInt32 Flags;
            public UInt32 Time;
            public IntPtr ExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left; public int Top; public int Right; public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public static implicit operator Point(POINT point)
            {
                return new Point(point.X, point.Y);
            }
        }

        internal struct KEYBDINPUT
        {
            public UInt16 Vk;
            public UInt16 Scan;
            public UInt32 Flags;
            public UInt32 Time;
            public IntPtr ExtraInfo;
        }
        internal struct HARDWAREINPUT
        {
            public UInt32 uMsg;
            public UInt16 wParamL;
            public UInt16 wParamH;
        }


        public const uint MOUSEEVENTF_LEFTDOWN = 0x02;
        public const uint MOUSEEVENTF_LEFTUP = 0x04;
        internal const UInt32 INPUT_MOUSE = 0;
        internal const UInt32 INPUT_KEYBOARD = 1;
        internal const UInt32 INPUT_HARDWARE = 2;

        internal const UInt32 KEYEVENTF_EXTENDEDKEY = 0x0001;
        internal const UInt32 KEYEVENTF_KEYUP = 0x0002;
        internal const UInt32 KEYEVENTF_UNICODE = 0x0004;
        internal const UInt32 KEYEVENTF_SCANCODE = 0x0008;

        IntPtr wnd = IntPtr.Zero;

        int imageIdx = 0;
        List<Bitmap> images = new List<Bitmap>();
        List<Bitmap> buttons = new List<Bitmap>();
        List<Point> buttonBounds = new List<Point>();
        Bitmap bitmap;

        public Form1()
        {
            InitializeComponent();
            Console.WriteLine("hello1");
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            images.Add(new Bitmap(@"C:\Users\sw_303\Desktop\수강생107\김정우\ohmygod\images\kiwoom\a.png"));
            //images.Add(new Bitmap(@".\images\meritz\인증서(1).png"));
            //images.Add(new Bitmap(@".\images\meritz\로그인.png"));
            //images.Add(new Bitmap(@".\images\meritz\현재가.png"));
            //images.Add(new Bitmap(@".\images\meritz\매수.png"));
            //images.Add(new Bitmap(@".\images\meritz\주식매수주문.png"));
            if (FindWindow(null, "영웅문4 Login") == IntPtr.Zero && FindWindow(null, "영웅문4") == IntPtr.Zero)
            {
                ProcessStartInfo start = new ProcessStartInfo();
                start.FileName = @"C:\KiwoomHero4\Bin\NKStarter.exe";
                start.UseShellExecute = true;
                start.Verb = "runas";
                try
                {
                    Process.Start(start);
                    while (wnd == IntPtr.Zero) { wnd = FindWindow(null, "영웅문4 Login"); Thread.Sleep(1000); }
                }
                catch (Exception ex)
                {

                }
            }
            wnd = FindWindow(null, "영웅문4 Login");
            wnd = FindWindow(null, "영웅문4") == IntPtr.Zero ? wnd : FindWindow(null, "영웅문4");
            SetWindowPos(wnd, IntPtr.Zero, 0, 0, 1000, 800, 0);
            button1.Visible = true;
            ewfaklwje();
            SetForegroundWindow(wnd);
            SetFocus(wnd);
            clearscreen();
            click(new Point(57, 88));
            typeString("0345", true);
            Thread.Sleep(100);
            Rectangle rect;
            findAndClick(out rect, new Bitmap(@"C:\Users\sw_303\Desktop\수강생107\김정우\ohmygod\images\kiwoom\account.png"));
            Thread.Sleep(100);
            Pix pix = PixConverter.ToPix(ocr(new Vector2(rect.X + rect.Width + 5, rect.Y), new Vector2(95, 16)));

            var engine = new TesseractEngine(@"./tessdata", "kor+eng", EngineMode.TesseractOnly);
            engine.SetVariable("tessedit_char_whitelist", "0123456789");
            var result = engine.Process(pix);
            Console.WriteLine(result.GetText() + " fjewlakfjwlaejflawef");
            BeginInvoke(new Action(() =>
            {
                label2.Text = result.GetText();
                result.Dispose();
                pix.Dispose();
                engine.Dispose();
            }));
            
            
            //new Thread(a).Start();
        }

        private static ushort CharToVirtualKey(char ch)
        {
            if (ch >= 'A' && ch <= 'Z') return (ushort)(ch - 'A' + 0x41);
            if (ch >= 'a' && ch <= 'z') return (ushort)(ch - 'a' + 0x41);
            if (ch >= '0' && ch <= '9') return (ushort)(ch - '0' + 0x30);
            if (ch == ' ') return 0x20;
            return 0;
        }

        private void PressKey(ushort vk, bool extended = false)
        {
            var down = new INPUT
            {
                Type = INPUT_KEYBOARD,
                Data = new MOUSEKEYBDHARDWAREINPUT
                {
                    Keyboard = new KEYBDINPUT
                    {
                        Vk = vk,
                        Scan = 0,
                        Flags = extended ? KEYEVENTF_EXTENDEDKEY : 0,
                        Time = 0,
                        ExtraInfo = IntPtr.Zero
                    }
                }
            };

            var up = new INPUT
            {
                Type = INPUT_KEYBOARD,
                Data = new MOUSEKEYBDHARDWAREINPUT
                {
                    Keyboard = new KEYBDINPUT
                    {
                        Vk = vk,
                        Scan = 0,
                        Flags = KEYEVENTF_KEYUP |
                            (extended ? KEYEVENTF_EXTENDEDKEY : 0),
                        Time = 0,
                        ExtraInfo = IntPtr.Zero
                    }
                }
            };

            INPUT[] inputs = new INPUT[] { down, up };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        }

        private void typeString(string str, bool pressenter = false)
        {
            foreach (char a in str)
            {
                Console.WriteLine(a);
                PressKey(CharToVirtualKey(a));
                Thread.Sleep(50);
            }
            if (pressenter) PressKey(0x0D);
        }

        private void findAndClick(out Rectangle rt, Bitmap image, Point offset = new Point())
        {
            RECT rect;
            Rectangle rectangle;
            GetWindowRect(wnd, out rect);
            rectangle = checkImage(new Point(rect.Left, rect.Top), new Point(rect.Right, rect.Bottom), new Vector2(rect.Right - rect.Left, rect.Bottom - rect.Top), image);
            rt = rectangle;
            if (rectangle.X == -1)
            {
                BeginInvoke(new Action(() =>
                {
                    textBox1.Text += "ewfawe\n";
                }));

                return;
            }

            click(new Point(rectangle.X + rectangle.Width / 2 + offset.X, rectangle.Y + rectangle.Height / 2 + offset.Y));
        }

        private void SetBuyOrder(string stockcode, int size, int price)
        {
            Point[] offset = { new Point(350, 30), new Point(270, 70), new Point(270, 117), new Point(270, 150), new Point(270, 205) };
            string[] input = { "690201", stockcode, size.ToString(), price.ToString() };
            Bitmap buy = new Bitmap(@"C:\Users\sw_303\Desktop\수강생107\김정우\ohmygod\images\kiwoom\buy.png");
            Rectangle rect;
            findAndClick(out rect, buy, new Point(225, 48));
            for (int i = 0; i < offset.Length; i++)
            {
                findAndClick(out rect, buy, offset[i]);
                if (i < input.Length) typeString(input[i]);
                Thread.Sleep(50);
            }
            click(new Point(452, 419));
        }

        private void SetSellOffer(string stockcode, int size, int price)
        {
            Point[] offset = { new Point(350, 30), new Point(270, 70), new Point(270, 117), new Point(270, 150), new Point(270, 205) };
            string[] input = { "690201", stockcode, size.ToString(), price.ToString() };
            Bitmap buy = new Bitmap(@"C:\Users\sw_303\Desktop\수강생107\김정우\ohmygod\images\kiwoom\buy.png");
            Rectangle rect;
            findAndClick(out rect, buy, new Point(285, 48));
            for (int i = 0; i < offset.Length; i++)
            {
                findAndClick(out rect, buy, offset[i]);
                if (i < input.Length) typeString(input[i]);
                Thread.Sleep(50);
            }
            click(new Point(452, 419));
        }

        private void a()
        {
            while (true)
            {
                POINT point;
                GetCursorPos(out point);
                this.label1.Text = point.X + ", " + point.Y;
                Thread.Sleep(50);
            }

        }

        void ewfaklwje()
        {
            Bitmap aa = new Bitmap(@"C:\Users\sw_303\Desktop\수강생107\김정우\ohmygod\images\kiwoom\awef.png");
            int[] x = { (320 + 370) / 2, (371 + 421) / 2, (422 + 465) / 2, (466 + 509) / 2, (510 + 553) / 2, (554 + 597) / 2, (598 + 641) / 2, (642 + 685) / 2, (686 + 729) / 2, (730 + 773) / 2, (774 + 818) / 2 };
            foreach (int a in x)
            {
                buttonBounds.Add(new Point(a, 69));
            }
        }

        Rectangle checkImage(Point start, Point end, Vector2 size, Bitmap bitmap)
        {
            //Console.WriteLine(end.Y - images[imageIdx].Height);
            //Console.WriteLine(imageIdx);
            Vector2 screenSize = size;
            using (Bitmap screen = new Bitmap(1920, 1080))
            {
                using (Graphics graphic = Graphics.FromImage(screen))
                {
                    graphic.CopyFromScreen(new Point(0, 0), Point.Empty, screen.Size);
                    for (int i = start.X; i < end.X - bitmap.Width; i++)
                    {
                        for (int j = start.Y; j < end.Y - bitmap.Height; j++)
                        {
                            if (check(screen, i, j, bitmap))
                            {
                                int a = imageIdx;
                                imageIdx++;
                                return new Rectangle(i, j, bitmap.Width, bitmap.Height);
                            }
                        }
                    }
                }
            }
            Console.WriteLine("fewa");
            return new Rectangle(-1, -1, -1, -1);
        }

        bool check(Bitmap screen, int x, int y, Bitmap bitmap)
        {
            for (int i = 0; i < bitmap.Width; i++)
            {
                for (int j = 0; j < bitmap.Height; j++)
                {
                    if (bitmap.GetPixel(i, j).A == 0) continue;
                    if (bitmap.GetPixel(i, j) != screen.GetPixel(i + x, j + y)) { return false; }
                }
            }
            return true;
        }

        public static void mouseclick()
        {
            var inputMouseDown = new INPUT();
            inputMouseDown.Type = 0;
            inputMouseDown.Data.Mouse.Flags = 0x0002;
            var inputMouseUp = new INPUT();
            inputMouseUp.Type = 0;
            inputMouseUp.Data.Mouse.Flags = 0x0004;
            var inputs = new INPUT[] { inputMouseDown, inputMouseUp };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        public void test()
        {
            var automation = new CUIAutomation();

            // Get the desktop root element
            IUIAutomationElement root = automation.GetRootElement();

            // 1. Find the window by its title (NameProperty)
            string windowTitle = "메리츠증권 iMeritz XII 로그인"; // <-- change to your window title
            IUIAutomationCondition windowCondition = automation.CreatePropertyCondition(
                UIA_PropertyIds.UIA_NamePropertyId,
                windowTitle
            );

            IUIAutomationElement window = automation.ElementFromHandle(wnd);

            if (window == null)
            {
                Console.WriteLine("Window not found.");
                return;
            }

            Console.WriteLine("Window found!");

            var emptyNameCondition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_NamePropertyId, "");
            var dialogTypeCondition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_WindowControlTypeId);
            var dialogCondition = automation.CreateAndCondition(emptyNameCondition, dialogTypeCondition);
            var dialog = window.FindFirst(TreeScope.TreeScope_Children, dialogCondition);

            if (dialog == null)
            {
                MessageBox.Show("Dialog not found!");
                return;
            }

            var buttonIdCondition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_AutomationIdPropertyId, "1003");
            var buttonTypeCondition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_ButtonControlTypeId);
            var buttonCondition = automation.CreateAndCondition(buttonIdCondition, buttonTypeCondition);
            var button = dialog.FindFirst(TreeScope.TreeScope_Descendants, buttonCondition);

            var invokePattern = (IUIAutomationInvokePattern)button.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId);
            invokePattern.Invoke();
            Console.WriteLine("Button clicked!");
        }

        private async Task click(Point point)
        {
            SetCursorPos(point.X, point.Y);
            Thread.Sleep(100);
            Mouse.LeftClick(point);
        }


        private async Task clearscreen()
        {
            RECT rect;
            SetForegroundWindow(wnd);
            SetFocus(wnd);
            GetWindowRect(wnd, out rect);
            Rectangle rectangle;
            while ((rectangle = checkImage(new Point(rect.Left, rect.Top), new Point(rect.Right, rect.Bottom), new Vector2(rect.Right - rect.Left, rect.Bottom - rect.Top), images[0])).X != -1)
            {
                Console.WriteLine(rectangle.X + " " + rectangle.Y);
                click(new Point(rectangle.X + rectangle.Width / 2, rectangle.Y + rectangle.Height / 2));
                //SetForegroundWindow(wnd);
                //SetFocus(wnd);
                //await click(new Point(rectangle.X + rectangle.Width / 2, rectangle.Y + rectangle.Height / 2));
                imageIdx = 0;
                Task.Delay(2000);
            }
        }

        private async void main()
        {
            RECT rect;
            Rectangle rectangle;
            SetForegroundWindow(wnd);
            SetFocus(wnd);
            GetWindowRect(wnd, out rect);
            rectangle = checkImage(new Point(rect.Left, rect.Top), new Point(rect.Right, rect.Bottom), new Vector2(rect.Right - rect.Left, rect.Bottom - rect.Top), buttons[4]);
            if (rectangle.X == -1)
            {
                textBox1.Text += "ewfawe\n";
                return;
            }
            SetForegroundWindow(wnd);
            SetFocus(wnd);

            click(new Point(rectangle.X + rectangle.Width / 2, rectangle.Y + rectangle.Height / 2));
        }


        private void login(int a, string pin = "")
        {
            SetForegroundWindow(wnd);
            SetFocus(wnd);
            click(new Point(407 + a * 58, 40));
            switch (a)
            {
                case 0:
                case 1:
                    click(new Point(493, 90));
                    Thread.Sleep(100);
                    click(new Point(500, 485));
                    typeString(pin);
                    Thread.Sleep(100);
                    click(new Point(600, 525));
                    break;
                case 2:
                case 3:
                    break;
            }
            wnd = FindWindow(null, "영웅문4");
            while (wnd == IntPtr.Zero) { wnd = FindWindow(null, "영웅문4"); Thread.Sleep(1000); }
            SetWindowPos(wnd, IntPtr.Zero, 0, 0, 1000, 800, 0);
            clearscreen();
            click(new Point(57, 88));
            typeString("0345", true);
        }

        Bitmap ocr(Vector2 start, Vector2 size)
        {
            Vector2 screenSize = size;
            Bitmap screen = new Bitmap((int)size.X, (int)size.Y);
            using (Graphics graphic = Graphics.FromImage(screen))
            {
                graphic.CopyFromScreen(new Point((int)start.X, (int)start.Y), Point.Empty, screen.Size);
            }
            this.bitmap = screen;
            return screen;
        }
        private async void button1_Click(object sender, EventArgs e)
        {
            await Task.Run(() =>
            {
                clearscreen();
                click(buttonBounds[2]);
                Thread.Sleep(1000);
                Pix pix = PixConverter.ToPix(ocr(new Vector2(14, 565), new Vector2(700, 20)));

                var engine = new TesseractEngine(@"./tessdata", "kor+eng", EngineMode.Default);
                var result = engine.Process(pix);

                if (result.GetText().Contains("[999999]"))
                {
                    MessageBox.Show("장이 열리지 않은 날입니다.");
                    return;
                }
                SetSellOffer("039490", 1, 200000);
            });
            //await Task.Run(() => { main(); });

            //TesseractEngine t = new TesseractEngine("./tessdata", "kor", EngineMode.Default);
            //t.SetVariable("tessedit_char_whitelist", "-01234567890");
            //Console.WriteLine(t.Process(Pix.LoadFromFile("C:\\Users\\sw_303\\Desktop\\cap.png"), PageSegMode.SingleBlock).GetText());
            //button1.Enabled = false;
            //run();
            //timer1.Start();


            //if (GetWindowRect(wnd, out rect))
            //{
            //    int windowWidth = rect.Right - rect.Left;
            //    int windowHeight = rect.Bottom - rect.Top;

            //    Console.WriteLine($"Window Width: {windowWidth}, Height: {windowHeight}");
            //    double clientXPercent = 0.67;
            //    double clientYPercent = 0.79;

            //    int clientX = rect.Left + (int)(clientXPercent * (windowWidth));
            //    int clientY = rect.Top + (int)(clientYPercent * (windowHeight));

            //    Point point = new Point { X = clientX, Y = clientY };
            //    //ClientToScreen(wnd, ref point);
            //    SetCursorPos(point.X, point.Y);
            //    Console.WriteLine($"Window Width: {point.X}, Height: {point.Y}");
            //    //SetForegroundWindow(wnd);
            //    //SetFocus(wnd);

            //    //mouseclick();
            //    mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)point.X, (uint)point.Y, 0, 0);
            //    mouse_event(MOUSEEVENTF_LEFTUP, (uint)point.X, (uint)point.Y, 0, 0);
            //}
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            login(0, "690201");
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.DrawImage(bitmap, new Point(0, 0));
        }
    }
}
