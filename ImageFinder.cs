using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tesseract;

namespace ohmygod
{
    internal class ImageFinder
    {
        public static List<Rectangle>? FindImage(string filename, Bitmap screenBitmap)
        {
            Graphics.FromImage(screenBitmap).CopyFromScreen(new System.Drawing.Point(0, 0), System.Drawing.Point.Empty, new System.Drawing.Size(screenBitmap.Width, screenBitmap.Height));
            Mat temp = new Mat(filename, ImreadModes.Grayscale);
            Mat mat = BitmapConverter.ToMat(screenBitmap);
            Cv2.CvtColor(mat, mat, ColorConversionCodes.BGR2GRAY);
            List<Rectangle> list = new();
            using (Mat res = new Mat())
            {
                Cv2.MatchTemplate(mat, temp, res, TemplateMatchModes.CCoeffNormed);
                Cv2.Threshold(res, res, 0.8, 1.0, ThresholdTypes.Tozero);
                while (true)
                {
                    OpenCvSharp.Point minloc, maxloc;
                    double minval, maxval;
                    Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc);
                    var threshold = 0.5;
                    if (maxval >= threshold)
                    {
                        list.Add(new Rectangle(maxloc.X, maxloc.Y, temp.Width, temp.Height));
                        Cv2.Rectangle(res, new OpenCvSharp.Rect(maxloc.X, maxloc.Y, temp.Width, temp.Height), Scalar.All(0), -1);
                    }
                    else break;
                }
            }
            return (list.Count != 0 ? list : null);
        }

        public static Rectangle FindSingleImage(string filename, Bitmap screenBitmap)
        {
            Graphics.FromImage(screenBitmap).CopyFromScreen(new System.Drawing.Point(0, 0), System.Drawing.Point.Empty, new System.Drawing.Size(screenBitmap.Width, screenBitmap.Height));
            Mat temp = new Mat(filename, ImreadModes.Grayscale);
            Mat mat = BitmapConverter.ToMat(screenBitmap);

            Cv2.CvtColor(mat, mat, ColorConversionCodes.BGR2GRAY);
            Cv2.Threshold(mat, mat, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
            Cv2.Threshold(temp, temp, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
            using (Mat res = new Mat())
            {
                Cv2.MatchTemplate(mat, temp, res, TemplateMatchModes.CCoeffNormed);
                Cv2.Threshold(res, res, 0.8, 1.0, ThresholdTypes.Tozero);
                OpenCvSharp.Point minloc, maxloc;
                double minval, maxval;
                Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc);
                var threshold = 0.5;
                if (maxval >= threshold)
                {
                    return new Rectangle(maxloc.X, maxloc.Y, temp.Width, temp.Height);
                }
                else return Rectangle.Empty;
            }
        }

        private static bool checkRect(Bitmap screen, int x, int y, Bitmap bitmap)
        {
            BitmapData bmpData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
            BitmapData screenData = screen.LockBits(new Rectangle(0, 0, screen.Width, screen.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb);
            unsafe
            {
                byte* ptr = (byte*)bmpData.Scan0;
                byte *ptr2 = (byte*)screenData.Scan0;

                for (int i = 0; i <= bitmap.Width; i++)
                {
                    byte* screenRow = ptr2 + (y + i) * screenData.Stride * 4;
                    byte* bitmapRow = ptr + y * bmpData.Stride * 4;
                    for (int j = 0; j <= bitmap.Height; j++)
                    {
                        byte* t = bitmapRow + j * 4;
                        byte* t2 = screenRow + j * 4;
                        if (t[3] == 0) continue;
                        if (t[0] != t2[0] || t[1] != t2[1] || t[2] != t2[2] || t[3] != t2[3])
                        {
                            bitmap.UnlockBits(bmpData);
                            screen.UnlockBits(screenData);
                            return false;
                        }
                    }
                }
                
            }
            bitmap.UnlockBits(bmpData);
            screen.UnlockBits(screenData);
            return true;
        }

        public static Rectangle FindTextImage(string filename, Bitmap screenBitmap)
        {
            Graphics.FromImage(screenBitmap).CopyFromScreen(new System.Drawing.Point(0, 0), System.Drawing.Point.Empty, new System.Drawing.Size(screenBitmap.Width, screenBitmap.Height));
            Bitmap bitmap = new Bitmap(filename);
            for(int i = 0; i<screenBitmap.Width-bitmap.Width; i++)
            {
                for(int j = 0; j<screenBitmap.Height-bitmap.Height; j++)
                {
                    if (checkRect(screenBitmap, i, j, bitmap))
                    {
                        return new Rectangle(i, j, bitmap.Width, bitmap.Height);
                    }
                }
            }
            return Rectangle.Empty;
        }

        public static Rectangle FindTextImage(string targetText, Bitmap screenBitmap, TesseractEngine tesseractEngine)
        {
            Graphics.FromImage(screenBitmap).CopyFromScreen(new System.Drawing.Point(0, 0), System.Drawing.Point.Empty, new System.Drawing.Size(screenBitmap.Width, screenBitmap.Height));
            var img = PixConverter.ToPix(screenBitmap);
            var page = tesseractEngine.Process(img);
            var iter = page.GetIterator();
            //Tesseract.Rect rect;
            iter.Begin();
            do
            {
                string text = iter.GetText(PageIteratorLevel.TextLine);
                if (!string.IsNullOrEmpty(text) && text.Contains(targetText))
                {
                    if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out Tesseract.Rect rect))
                    {
                        Console.WriteLine($"Found at X:{rect.X1}, Y:{rect.Y1}, W:{rect.Width}, H:{rect.Height}");
                        return new Rectangle(rect.X1, rect.Y1, rect.Width, rect.Height);
                    }
                }
            } while (iter.Next(PageIteratorLevel.TextLine));
            return Rectangle.Empty;
        }
    }
}
