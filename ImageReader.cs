using Mysqlx.Crud;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Org.BouncyCastle.Asn1.Crmf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using Tesseract;

namespace ohmygod
{
    internal class ImageReader
    {
        public static ImageReader instance;

        public ImageReader() { }

        public static ImageReader GetInstance()
        {
            if(instance == null) instance = new ImageReader();
            return instance;
        }
        
        public String ReadBitmap(TesseractEngine engine, Vector2 pos, Vector2 size)
        {
            //Bitmap bitmap = ScreenToBitmap(new Vector2(rect.X, rect.Y + 3), new Vector2(39, 16));
            Bitmap bitmap = ScreenToBitmap(pos, size);
            OpenCvSharp.Mat src = OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap), dst = new OpenCvSharp.Mat();
            OpenCvSharp.Cv2.Resize(src, dst, new OpenCvSharp.Size(size.X*1.75, size.Y*1.75));
            //bitmap = BitmapConverter.ToBitmap(dst);
            //dst = new Mat();
            //OpenCvSharp.Cv2.CopyMakeBorder(BitmapConverter.ToMat(bitmap), dst, 0, (int)(size.Y*1.75), 0, 0, BorderTypes.Transparent);
            Pix pix = PixConverter.ToPix(OpenCvSharp.Extensions.BitmapConverter.ToBitmap(dst));
            //Pix pix = Pix.LoadFromFile(@"C:\Users\sw_303\Pictures\Screenshots\test.png");
            var result = engine.Process(pix);
            string res = result.GetText();
            Console.WriteLine(result.GetText() + " fjewlakfjwlaejflawef");

            result.Dispose();
            pix.Dispose();
            OpenCvSharp.Extensions.BitmapConverter.ToBitmap(dst).Save(@"C:\Users\sw_303\Pictures\Screenshots\test.png");
            return res;
        }

        private Bitmap ScreenToBitmap(Vector2 start, Vector2 size)
        {
            Vector2 screenSize = size;
            Bitmap screen = new Bitmap((int)size.X, (int)size.Y);
            using (Graphics graphic = Graphics.FromImage(screen))
            {
                graphic.CopyFromScreen(new System.Drawing.Point((int)start.X, (int)start.Y), System.Drawing.Point.Empty, screen.Size);
            }
            return screen;
        }
    }
}
