using Interop.UIAutomationClient;
using System;
using System.Drawing;
using System.Numerics;
using System.Runtime.InteropServices;
using Tesseract;
using System.Windows.Forms.Automation;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.Input;
using System.Diagnostics;

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

        const byte AltKey = 18; const int KEYUP = 0x0002;

        internal struct INPUT
        {
            public UInt32 Type;
            public MOUSEKEYBDHARDWAREINPUT Data;
        }
        [StructLayout(LayoutKind.Explicit)]
        internal struct MOUSEKEYBDHARDWAREINPUT
        {
            [FieldOffset(0)]
            public MOUSEINPUT Mouse;
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


        public const uint MOUSEEVENTF_LEFTDOWN = 0x02;
        public const uint MOUSEEVENTF_LEFTUP = 0x04;

        IntPtr wnd = IntPtr.Zero;

        int imageIdx = 0;
        List<Bitmap> images = new List<Bitmap>();
        List<Bitmap> buttons = new List<Bitmap>();

        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            images.Add(new Bitmap(@"C:\Users\sw_303\Desktop\수강생107\김정우\ohmygod\images\kiwoom\a.png"));
            //images.Add(new Bitmap(@".\images\meritz\인증서(1).png"));
            //images.Add(new Bitmap(@".\images\meritz\로그인.png"));
            //images.Add(new Bitmap(@".\images\meritz\현재가.png"));
            //images.Add(new Bitmap(@".\images\meritz\매수.png"));
            //images.Add(new Bitmap(@".\images\meritz\주식매수주문.png"));
            if(FindWindow(null,"영웅문4 Login") == IntPtr.Zero && FindWindow(null,"영웅문4") == IntPtr.Zero)
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
            wnd = FindWindow(null, "영웅문4") == IntPtr.Zero ? wnd : FindWindow(null,"영웅문4");
            button1.Visible = true;
            ewfaklwje();
            //new Thread(a).Start();
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
            Graphics few = Graphics.FromImage(aa);
            buttons.Add(aa.Clone(new Rectangle(0, 0, 51, 40), aa.PixelFormat));
            buttons.Add(aa.Clone(new Rectangle(51, 0, 51, 40), aa.PixelFormat));
            for (int i = 2; i<28; i++)
            {
                Bitmap a = aa.Clone(new Rectangle(i * 44, 0, 44, 40), aa.PixelFormat);
                buttons.Add(a);
            }
        }

        Rectangle checkImage(Point start, Point end, Vector2 size, Bitmap bitmap)
        {
            //Console.WriteLine(end.Y - images[imageIdx].Height);
            //Console.WriteLine(imageIdx);
            Vector2 screenSize = size;
            using (Bitmap screen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
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
            while ((rectangle = checkImage(new Point(rect.Left, rect.Top), new Point(rect.Right, rect.Bottom), new Vector2(rect.Right - rect.Left, rect.Bottom - rect.Top),images[0])).X != -1)
            {
                Console.WriteLine(rectangle.X + " " + rectangle.Y);
                SetForegroundWindow(wnd);
                SetFocus(wnd);
                await click(new Point(rectangle.X + rectangle.Width / 2, rectangle.Y + rectangle.Height / 2));
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

            click(new Point(rectangle.X + rectangle.Width/2, rectangle.Y + rectangle.Height/2));
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            await Task.Run(() => { clearscreen(); });
            await Task.Run(() => { main(); });
            
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
    }
}
