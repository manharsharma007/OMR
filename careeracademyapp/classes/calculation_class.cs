using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace omrapplication.classes
{
    class calculation_class
    {

        public Size calculateBlockSize(int rows, int column, int optionWidth, int optionHeight, int XSpacing, int YSpacing)
        {
            int width = 0, height = 0;
            for (int i = 0; i < column; i++)
            {
                if (i == 0 || i == column - 1)
                {
                    width += optionWidth;
                }
                else
                {
                    width += XSpacing;
                    width += optionWidth;
                }
            }

            for (int i = 0; i < rows; i++)
            {
                if (i == 0 || i == rows - 1)
                {
                    height += optionHeight;
                }
                else
                {
                    height += YSpacing;
                    height += optionHeight;
                }
            }
            return new Size(width, height);
        }

        public double percentageIncrease(int original_number, int new_number)
        {
            double increase = new_number - original_number;
            increase = ((double)increase / original_number) * 100;
            return increase;
        }
        public double percentageDecrease(int original_number, int new_number)
        {
            double decrease = original_number - new_number;
            decrease = ((double)decrease / original_number) * 100;
            return decrease;
        }

        public Size increaseScale(int width, int height, double scale)
        {
            double w, h;
            w = width + (width * (scale / 100));
            h = height + (height * (scale / 100));
            return new Size((int)w, (int)h);
        }
        public Size decreaseScale(int width, int height, double scale)
        {
            double w, h;
            w = width - (width * ((double)scale / 100));
            h = height - (height * ((double)scale / 100));
            return new Size((int)w, (int)h);
        }

    }
    public class gmseDeskew
    {
        // Representation of a line in the image.
        public class HougLine
        {
            // Count of points in the line.
            public int Count;
            // Index in Matrix.
            public int Index;
            // The line is represented as all x,y that solve y*cos(alpha)-x*sin(alpha)=d
            public double Alpha;
            public double d;
        }
        // The Bitmap
        Bitmap cBmp;
        // The range of angles to search for lines
        double cAlphaStart = -20;
        double cAlphaStep = 0.2;
        int cSteps = 40 * 5;
        // Precalculation of sin and cos.
        double[] cSinA;
        double[] cCosA;
        // Range of d
        double cDMin;
        double cDStep = 1;
        int cDCount;
        // Count of points that fit in a line.

        int[] cHMatrix;
        // Calculate the skew angle of the image cBmp.
        public double GetSkewAngle()
        {
            gmseDeskew.HougLine[] hl = null;
            int i = 0;
            double sum = 0;
            int count = 0;

            // Hough Transformation
            Calc();
            // Top 20 of the detected lines in the image.
            hl = GetTop(20);
            // Average angle of the lines
            for (i = 0; i <= 19; i++)
            {
                sum += hl[i].Alpha;
                count += 1;
            }
            return sum / count;
        }

        // Calculate the Count lines in the image with most points.
        private HougLine[] GetTop(int Count)
        {
            HougLine[] hl = null;
            int i = 0;
            int j = 0;
            HougLine tmp = null;
            int AlphaIndex = 0;
            int dIndex = 0;

            hl = new HougLine[Count + 1];
            for (i = 0; i <= Count - 1; i++)
            {
                hl[i] = new HougLine();
            }
            for (i = 0; i <= cHMatrix.Length - 1; i++)
            {
                if (cHMatrix[i] > hl[Count - 1].Count)
                {
                    hl[Count - 1].Count = cHMatrix[i];
                    hl[Count - 1].Index = i;
                    j = Count - 1;
                    while (j > 0 && hl[j].Count > hl[j - 1].Count)
                    {
                        tmp = hl[j];
                        hl[j] = hl[j - 1];
                        hl[j - 1] = tmp;
                        j -= 1;
                    }
                }
            }
            for (i = 0; i <= Count - 1; i++)
            {
                dIndex = hl[i].Index / cSteps;
                AlphaIndex = hl[i].Index - dIndex * cSteps;
                hl[i].Alpha = GetAlpha(AlphaIndex);
                hl[i].d = dIndex + cDMin;
            }
            return hl;
        }
        public gmseDeskew(Bitmap bmp)
        {
            cBmp = bmp;
        }
        // Hough Transforamtion:
        private void Calc()
        {
            int x = 0;
            int y = 0;
            int hMin = cBmp.Height / 4;
            int hMax = cBmp.Height * 3 / 4;

            Init();
            for (y = hMin; y <= hMax; y++)
            {
                for (x = 1; x <= cBmp.Width - 2; x++)
                {
                    // Only lower edges are considered.
                    if (IsBlack(x, y))
                    {
                        if (!IsBlack(x, y + 1))
                        {
                            Calc(x, y);
                        }
                    }
                }
            }
        }
        // Calculate all lines through the point (x,y).
        private void Calc(int x, int y)
        {
            int alpha = 0;
            double d = 0;
            int dIndex = 0;
            int Index = 0;

            for (alpha = 0; alpha <= cSteps - 1; alpha++)
            {
                d = y * cCosA[alpha] - x * cSinA[alpha];
                dIndex = Convert.ToInt32(CalcDIndex(d));
                Index = dIndex * cSteps + alpha;
                try
                {
                    cHMatrix[Index] += 1;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }
        private double CalcDIndex(double d)
        {
            return Convert.ToInt32(d - cDMin);
        }
        private bool IsBlack(int x, int y)
        {
            Color c = default(Color);
            double luminance = 0;

            c = cBmp.GetPixel(x, y);
            luminance = (c.R * 0.299) + (c.G * 0.587) + (c.B * 0.114);
            return luminance < 140;
        }
        private void Init()
        {
            int i = 0;
            double angle = 0;

            // Precalculation of sin and cos.
            cSinA = new double[cSteps];
            cCosA = new double[cSteps];
            for (i = 0; i <= cSteps - 1; i++)
            {
                angle = GetAlpha(i) * Math.PI / 180.0;
                cSinA[i] = Math.Sin(angle);
                cCosA[i] = Math.Cos(angle);
            }
            // Range of d:
            cDMin = -cBmp.Width;
            cDCount = Convert.ToInt32(2 * (cBmp.Width + cBmp.Height) / cDStep);
            cHMatrix = new int[cDCount * cSteps + 1];
        }

        public double GetAlpha(int Index)
        {
            return cAlphaStart + Index * cAlphaStep;
        }
        public bool NoiseRemoval(Bitmap IntensityImage)
        {

            /*It removes the pixel
             * that is stood alone any where in the 
             * vicinity.
             *It is found to be accurate 4 our System.
             * */

            Bitmap b2 = (Bitmap)IntensityImage.Clone();
            byte val;
            // GDI+ still lies to us - the return format is BGR, NOT RGIntensityImage.
            BitmapData bmData = IntensityImage.LockBits(new Rectangle(0, 0, IntensityImage.Width, IntensityImage.Height),
                ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            BitmapData bmData2 = b2.LockBits(new Rectangle(0, 0, IntensityImage.Width, IntensityImage.Height),
                ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

            int stride = bmData.Stride;
            System.IntPtr Scan0 = bmData.Scan0;
            System.IntPtr Scan02 = bmData2.Scan0;

            unsafe
            {
                byte* p = (byte*)(void*)Scan0;
                byte* p2 = (byte*)(void*)Scan02;

                int nOffset = stride - IntensityImage.Width * 3;
                int nWidth = IntensityImage.Width * 3;

                //int nPixel=0;

                p += stride;
                p2 += stride;
                //int val;
                for (int y = 1; y < IntensityImage.Height - 1; ++y)
                {
                    p += 3;
                    p2 += 3;

                    for (int x = 3; x < nWidth - 3; ++x)
                    {
                        val = p2[0];
                        if (val == 0)
                            if ((p2 + 3)[0] == 0 || (p2 - 3)[0] == 0 || (p2 + stride)[0] == 0 || (p2 - stride)[0] == 0 || (p2 + stride + 3)[0] == val || (p2 + stride - 3)[0] == 0 || (p2 - stride - 3)[0] == 0 || (p2 + stride + 3)[0] == 0)
                                p[0] = val;
                            else
                                p[0] = 255;

                        ++p;
                        ++p2;
                    }

                    p += nOffset + 3;
                    p2 += nOffset + 3;
                }
            }

            IntensityImage.UnlockBits(bmData);
            b2.UnlockBits(bmData2);
            return true;
        }
        public bool Binary(Bitmap b, bool flag)
        {
            // GDI+ still lies to us - the return format is BGR, NOT RGB. 
            BitmapData bmData = b.LockBits(new Rectangle(0, 0, b.Width, b.Height),
                ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            int stride = bmData.Stride; //the length of the line
            System.IntPtr Scan0 = bmData.Scan0;
            int Threshold = 220;
            unsafe
            {
                byte* p = (byte*)(void*)Scan0;

                int nOffset = stride - b.Width * 3;

                byte red, green, blue;
                byte binary;

                for (int y = 0; y < b.Height; ++y)
                {
                    for (int x = 0; x < b.Width; ++x)
                    {
                        blue = p[0];
                        green = p[1];
                        red = p[2];

                        binary = (byte)(.299 * red
                            + .587 * green
                            + .114 * blue);

                        if (binary < Threshold && flag)
                            p[0] = p[1] = p[2] = 0;
                        else
                            if (binary >= Threshold && flag)
                                p[0] = p[1] = p[2] = 255;
                            else
                                if (binary < Threshold && !flag)
                                    p[0] = p[1] = p[2] = 255;
                                else
                                    p[0] = p[1] = p[2] = 0;
                        p += 3;
                    }
                    p += nOffset;
                }

            }

            b.UnlockBits(bmData);
            return true;
        }
    }
}
