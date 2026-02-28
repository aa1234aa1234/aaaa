using FlaUI.Core.Input;
using System;
using System.Collections.Generic;
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
            SetCursorPos(point.X, point.Y);
            Mouse.LeftClick(point);
            Console.WriteLine(point);
        }

        public void SetWindowPos(IntPtr wnd, int x, int y, int width=0, int height=0)
        {
            SetWindowPos(wnd, IntPtr.Zero, x, y, width, height, 0);
        }

        public void FocusWindow(IntPtr wnd)
        {
            SetForegroundWindow(wnd);
            SetFocus(wnd);
        }
    }
}
