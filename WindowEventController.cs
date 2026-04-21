using FlaUI.Core.Input;
using MySqlX.XDevAPI.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ohmygod
{
    internal class WindowEventController
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
        public static extern bool ClientToScreen(IntPtr hWnd, ref System.Drawing.Point lpPoint);

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

            public static implicit operator System.Drawing.Point(POINT point)
            {
                return new System.Drawing.Point(point.X, point.Y);
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

        internal const UInt32 MOUSEEVENTF_MOVE = 0x0001;
        internal const UInt32 MOUSEEVENTF_ABSOLUTE = 0x8000;

        public delegate bool EnumPWindowProc(IntPtr hWnd, IntPtr parameters);
        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr window, EnumPWindowProc callback, IntPtr i);

        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumPWindowProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll")]
        static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();


        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("kernel32.dll")]
        static extern int GetProcessId(IntPtr handle);

        const int WH_CBT = 5;
        const int HCBT_CREATEWND = 3;

        const int WM_COMMAND = 0x0111;
        const int IDOK = 1;

        private static IntPtr hook;
        private static HookProc proc;

        public static void InstallMessageBoxAutoCloser()
        {
            proc = HookCallback;
            hook = SetWindowsHookEx(WH_CBT, proc, IntPtr.Zero, GetCurrentThreadId());
        }

        public static void RemoveHook()
        {
            if (hook != IntPtr.Zero)
                UnhookWindowsHookEx(hook);
        }

        public static int GetPid(IntPtr handle) { return GetProcessId(handle); }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == HCBT_CREATEWND)
            {
                StringBuilder className = new StringBuilder(256);
                GetClassName(wParam, className, className.Capacity);

                if (className.ToString() == "#32770")
                {
                    SendMessage(wParam, WM_COMMAND, (IntPtr)IDOK, IntPtr.Zero);
                }
            }

            return CallNextHookEx(hook, nCode, wParam, lParam);
        }

        private static bool EnumMessageBox(IntPtr handle, IntPtr pointer)
        {
            StringBuilder className = new StringBuilder(256);
            GetClassName(handle, className, className.Capacity);

            if(className.ToString() == "#32770")
            {
                GCHandle gch = GCHandle.FromIntPtr(pointer);
                List<IntPtr> list = gch.Target as List<IntPtr>;

                list.Add(handle);
            }
            return true;
        }

        public Point GetMousePos()
        {
            POINT pos;
            GetCursorPos(out pos);
            return new Point(pos.X, pos.Y);
        }

        public void MoveMouse(int x, int y)
        {
            MoveMouseAbsolute(x, y);
        }

        public void MoveRelative(int x, int y)
        {
            MoveMouseRelative(x, y);
        }

        public static void MoveMouseRelative(int dx, int dy)
        {
            INPUT[] input = new INPUT[1];
            input[0].Type = INPUT_MOUSE;
            input[0].Data.Mouse.X = dx;
            input[0].Data.Mouse.Y = dy;
            input[0].Data.Mouse.Flags = MOUSEEVENTF_MOVE;
            input[0].Data.Mouse.MouseData = 0;
            input[0].Data.Mouse.Time = 0;
            input[0].Data.Mouse.ExtraInfo = IntPtr.Zero;

            SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
        }

        public static void MoveMouseAbsolute(int x, int y)
        {
            int screenWidth = GetSystemMetrics(0);
            int screenHeight = GetSystemMetrics(1);

            int normalizedX = (int)(x * 65535 / (screenWidth - 1));
            int normalizedY = (int)(y * 65535 / (screenHeight - 1));

            INPUT[] input = new INPUT[1];
            input[0].Type = INPUT_MOUSE;
            input[0].Data.Mouse.X = normalizedX;
            input[0].Data.Mouse.Y = normalizedY;
            input[0].Data.Mouse.Flags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE;
            input[0].Data.Mouse.MouseData = 0;
            input[0].Data.Mouse.Time = 0;
            input[0].Data.Mouse.ExtraInfo = IntPtr.Zero;

            SendInput(1, input, Marshal.SizeOf(typeof(INPUT)));
        }

        public static void CloseMessageBoxes(Process process)
        {
            EnumWindows((hWnd, lParam) =>
            {
                GetWindowThreadProcessId(hWnd, out uint pid);

                if (pid != process.Id)
                    return true;

                StringBuilder className = new StringBuilder(256);
                GetClassName(hWnd, className, className.Capacity);

                if (className.ToString() == "#32770")
                {
                    SendMessage(hWnd, WM_COMMAND, (IntPtr)IDOK, IntPtr.Zero);
                }

                return true;

            }, IntPtr.Zero);
        }

        public static List<IntPtr> GetMessageBoxes(IntPtr handle)
        {
            List<IntPtr> result = new();

            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                EnumWindows(EnumMessageBox, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                listHandle.Free();
            }

            return result;
        }

        private static List<IntPtr> GetChildWindows(IntPtr parent, bool addParent = true)
        {
            List<IntPtr> result = new();

            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                EnumPWindowProc childProc = new EnumPWindowProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated) listHandle.Free();
            }

            if (addParent) result.Add(parent);

            return result;
        }

        private static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(pointer);
            List<IntPtr> list = gch.Target as List<IntPtr>;
            if(list == null)
            {
                throw new InvalidCastException();
            }
            list.Add(handle);
            return true;
        }

        public static List<IntPtr> GetChildEnums(IntPtr parent)
        {
            return GetChildWindows(parent, false);
        }

        IntPtr wnd = IntPtr.Zero;

        private static ushort CharToVirtualKey(char ch)
        {
            if (ch >= 'A' && ch <= 'Z') return (ushort)(ch - 'A' + 0x41);
            if (ch >= 'a' && ch <= 'z') return (ushort)(ch - 'a' + 0x41);
            if (ch >= '0' && ch <= '9') return (ushort)(ch - '0' + 0x30);
            if (ch == ' ') return 0x20;
            return 0;
        }

        public void PressKey(ushort vk, bool extended = false)
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

        public void typeString(string str, bool pressenter = false)
        {
            if (str == "#VK_ESCAPE#") { PressKey(0x1B); return; }
            foreach (char a in str)
            {
                Console.WriteLine(a);
                PressKey(CharToVirtualKey(a));
                Thread.Sleep(50);
            }
            if (pressenter) PressKey(0x0D);
        }

        public void MoveWindow(int x, int y, int width, int height)
        {
            SetWindowPos(wnd, IntPtr.Zero, x, y, width, height, 0);
        }

        public void Click(Point point, Point offset=new Point())
        {
            point.Offset(offset);
            //SetCursorPos(point.X, point.Y);
            Mouse.LeftClick(point);
            Console.WriteLine(point);
            Thread.Sleep(500);
        }

        public void SetWindowPos(IntPtr wnd, int x, int y, int width=0, int height=0)
        {
            SetWindowPos(wnd, IntPtr.Zero, x, y, width, height, 0);
        }

        public void SetMousePos(Point point)
        {
            SetCursorPos(point.X, point.Y);
        }

        public void FocusWindow(IntPtr wnd)
        {
            SetForegroundWindow(wnd);
            SetFocus(wnd);
        }
    }
}
