using System;
using System.Runtime.InteropServices;
using Windows.Graphics.Imaging;
using OpenCvSharp;

namespace OCRTools.Utils
{
    class ImageProcessor
    {
        private Mat matImage { get; set; }
        private BitmapPixelFormat bitmapPixelFormat { get; set; }
        private BitmapAlphaMode bitmapAlphaMode  { get; set; }

        public ImageProcessor(SoftwareBitmap softwareBitmap)
        {
            bitmapPixelFormat = softwareBitmap.BitmapPixelFormat;
            bitmapAlphaMode = softwareBitmap.BitmapAlphaMode;
            matImage = SoftwareBitmapToMat(softwareBitmap);
        }

        public SoftwareBitmap GetSoftwareBitmap()
        {
            SoftwareBitmap softwareBitmap = new SoftwareBitmap(bitmapPixelFormat, matImage.Width, matImage.Height, bitmapAlphaMode);
            MatToSoftwareBitmap(matImage, softwareBitmap);
            return softwareBitmap;
        }

        public void Resize(double scale)
        {
            if (scale < 1)
            {
                Cv2.Resize(matImage, matImage, new Size(matImage.Width * scale, matImage.Height * scale), interpolation: InterpolationFlags.Area);
            }
            else
            {
                Cv2.Resize(matImage, matImage, new Size(matImage.Width * scale, matImage.Height * scale), interpolation: InterpolationFlags.Linear);
            }
        }

        public void Crop(int x, int y, int width, int height)
        {
            Rect cropArea = new Rect(x, y, width, height);
            matImage = new Mat(matImage, cropArea);
        }

        public void Denoising()
        {
            Cv2.FastNlMeansDenoising(matImage, matImage);
        }

        public void GaussianBlur(int kernelSize)
        {
            Cv2.GaussianBlur(matImage, matImage, new Size(kernelSize, kernelSize), sigmaX: 0);
        }

        public void BilateralFilter(int kernelSize)
        {
            using (Mat gray = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC1))
            using (Mat bilateralFilter = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC1))
            {
                Cv2.CvtColor(matImage, gray, ColorConversionCodes.BGRA2GRAY);
                Cv2.BilateralFilter(gray, bilateralFilter, kernelSize, 80, 80);
                bilateralFilter.ConvertTo(matImage, MatType.CV_8UC4);
                Cv2.CvtColor(matImage, matImage, ColorConversionCodes.GRAY2BGRA);
            }
        }

        public void Sharpen(int type=1)
        {
            using (Mat sharpen = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            {
                //double[] sharpenArray = { -1, -1 -1,
                //                          -1, 9, -1,
                //                          -1, -1,-1 };
                //double[] sharpenArray = { -0.5, -0.5, -0.5,
                //                          -0.5, 5, -0.5,
                //                          -0.5, -0.5,-0.5 };
                double[] sharpenArray = { 1, 1, 1,
                                          1, 1, 1,
                                          1, 1, 1};
                if (type == 1)
                {
                    sharpenArray = new double[]{ 0, -1, 0,
                                                -1, 5, -1,
                                                0, -1,0 };
                }
                else if (type == 2)
                {
                    sharpenArray = new double[]{ 0, -1.5, 0,
                                                -1.5, 7, -1.5,
                                                0, -1.5,0 };

                }
                else if (type == 3)
                {
                    sharpenArray = new double[] { -0.5, -1, -0.5,
                                                  -1, 7, -1,
                                                  -0.5, -1,-0.5 };

                }

                else if (type == 4)
                {
                    sharpenArray = new double[]{ -0.5, -0.5, -0.5,
                                                 -0.5, 5, -0.5,
                                                 -0.5, -0.5,-0.5 };
                }

                else if (type == 5)
                {
                    sharpenArray = new double[]{ -1, -1 -1,
                                                 -1, 9, -1,
                                                 -1, -1,-1 };
                }

                var sharpeKernel = new Mat(rows: 3, cols: 3, type: MatType.CV_64FC1, data: sharpenArray);
                Cv2.Filter2D(matImage, sharpen, MatType.CV_64FC1, sharpeKernel);
                sharpen.ConvertTo(matImage, MatType.CV_8UC4);
            }
        }

        public void Invert()
        {
            using (Mat gray = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            using (Mat invert = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            {
                Cv2.CvtColor(matImage, gray, ColorConversionCodes.BGRA2GRAY);
                Cv2.Threshold(gray, invert, 127, 255, ThresholdTypes.BinaryInv);
                Cv2.CvtColor(invert, matImage, ColorConversionCodes.GRAY2BGRA);
            }
        }

        public void BitwiseNot()
        {
            Cv2.BitwiseNot(matImage, matImage);
        }

        public void AdaptiveThreshold(int blockSize=3)
        {
            using (Mat gray = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC1))
            using (Mat adaptiveThreshold = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC1))
            {
                Cv2.CvtColor(matImage, gray, ColorConversionCodes.BGRA2GRAY);
                gray.ConvertTo(gray, MatType.CV_8UC1);
                Cv2.AdaptiveThreshold(gray, adaptiveThreshold, 255, AdaptiveThresholdTypes.GaussianC, ThresholdTypes.Binary, blockSize, -2);
                adaptiveThreshold.ConvertTo(matImage, MatType.CV_8UC4);
                Cv2.CvtColor(matImage, matImage, ColorConversionCodes.GRAY2BGRA);
            }
        }

        public void Otsu()
        {
            using (Mat gray = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            using (Mat ostu = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            {
                Cv2.CvtColor(matImage, gray, ColorConversionCodes.BGRA2GRAY);
                Cv2.Threshold(gray, ostu, 127, 255, ThresholdTypes.Otsu);
                Cv2.CvtColor(ostu, matImage, ColorConversionCodes.GRAY2BGRA);
            }
        }

        public void GrayToZero()
        {
            using (Mat gray = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            using (Mat ostu = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            {
                Cv2.CvtColor(matImage, gray, ColorConversionCodes.BGRA2GRAY);
                Cv2.Threshold(gray, ostu, 127, 255, ThresholdTypes.Tozero);
                Cv2.CvtColor(ostu, matImage, ColorConversionCodes.GRAY2BGRA);
            }
        }

        public void GrayTrunc()
        {
            using (Mat gray = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            using (Mat ostu = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            {
                Cv2.CvtColor(matImage, gray, ColorConversionCodes.BGRA2GRAY);
                Cv2.Threshold(gray, ostu, 127, 255, ThresholdTypes.Trunc);
                Cv2.CvtColor(ostu, matImage, ColorConversionCodes.GRAY2BGRA);
            }

        }
        public void MorphologyExOpen(int kernelSize=3, int iterations=1)
        {
            var element = Cv2.GetStructuringElement(
                            MorphShapes.Rect,
                            new Size(kernelSize, kernelSize));
            Cv2.MorphologyEx(matImage, matImage, MorphTypes.Open, element,iterations: iterations);
        }

        public void MorphologyExClose(int kernelSize = 3, int iterations = 1)
        {
            var element = Cv2.GetStructuringElement(
                            MorphShapes.Rect,
                            new Size(kernelSize, kernelSize));
            Cv2.MorphologyEx(matImage, matImage, MorphTypes.Close, element, iterations: iterations);
        }

        public void MorphologyExErode(int kernelSize = 3, int iterations = 1)
        {
            var element = Cv2.GetStructuringElement(
                            MorphShapes.Ellipse,
                            new Size(kernelSize, kernelSize));
            Cv2.MorphologyEx(matImage, matImage, MorphTypes.Erode, element, iterations: iterations);
        }

        public void MorphologyExDilate(int kernelSize = 3, int iterations = 1)
        {
            var element = Cv2.GetStructuringElement(
                            MorphShapes.Ellipse,
                            new Size(kernelSize, kernelSize));
            Cv2.MorphologyEx(matImage, matImage, MorphTypes.Dilate, element, iterations: iterations);
        }

        public void Canny()
        {
            //using (Mat mOutput = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            using (Mat intermediate = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            {
                Cv2.Canny(matImage, intermediate,
                    threshold1: 150,
                    threshold2: 200,
                    apertureSize: 3);

                Cv2.CvtColor(intermediate, matImage, ColorConversionCodes.GRAY2BGRA);
            } 
        }

        public void DetectHorizontal()
        {
            var element = Cv2.GetStructuringElement(
                            MorphShapes.Rect,
                            new Size(matImage.Cols/70, 1));
            Cv2.Erode(matImage, matImage, element, new Point(-1,-1));
            Cv2.Dilate(matImage, matImage, element, new Point(-1, -1));        
        }

        public void DetectVertical()
        {
            var element = Cv2.GetStructuringElement(
                            MorphShapes.Rect,
                            new Size(1, matImage.Rows / 30));
            Cv2.Erode(matImage, matImage, element, new Point(-1, -1));
            Cv2.Dilate(matImage, matImage, element, new Point(-1, -1));  
        }

        public void RemoveVerticalLines()
        {
            using (Mat vertical = matImage.Clone())
            {
                var verticalElement = Cv2.GetStructuringElement(
                                MorphShapes.Rect,
                                new Size(1, 130));
                Cv2.Erode(vertical, vertical, verticalElement, new Point(-1, -1));
                Cv2.Dilate(vertical, vertical, verticalElement, new Point(-1, -1));
                var element = Cv2.GetStructuringElement(
                                    MorphShapes.Rect,
                                    new Size(3, 3));
                using (Mat gray = new Mat(vertical.Rows, vertical.Cols, MatType.CV_8UC4))
                using (Mat invert = new Mat(vertical.Rows, vertical.Cols, MatType.CV_8UC4))
                {
                    Cv2.CvtColor(vertical, gray, ColorConversionCodes.BGRA2GRAY);
                    Cv2.Threshold(gray, invert, 40, 255, ThresholdTypes.BinaryInv);
                    Cv2.CvtColor(invert, vertical, ColorConversionCodes.GRAY2BGRA);
                }
                Cv2.MorphologyEx(vertical, vertical, MorphTypes.Erode, element, iterations: 3);
                Cv2.BitwiseAnd(matImage, vertical, matImage);
            }
        }

        public void RemoveHorizontalLines()
        {
            using (Mat horizontal = matImage.Clone())
            {
                var horizontalElement = Cv2.GetStructuringElement(
                                MorphShapes.Rect,
                                new Size(130, 1));
                Cv2.Erode(horizontal, horizontal, horizontalElement, new Point(-1, -1));
                Cv2.Dilate(horizontal, horizontal, horizontalElement, new Point(-1, -1));
                var element = Cv2.GetStructuringElement(
                                    MorphShapes.Rect,
                                    new Size(3, 3));
                using (Mat gray = new Mat(horizontal.Rows, horizontal.Cols, MatType.CV_8UC4))
                using (Mat invert = new Mat(horizontal.Rows, horizontal.Cols, MatType.CV_8UC4))
                {
                    Cv2.CvtColor(horizontal, gray, ColorConversionCodes.BGRA2GRAY);
                    Cv2.Threshold(gray, invert, 40, 255, ThresholdTypes.BinaryInv);
                    Cv2.CvtColor(invert, horizontal, ColorConversionCodes.GRAY2BGRA);
                }
                Cv2.MorphologyEx(horizontal, horizontal, MorphTypes.Erode, element, iterations: 2);
                Cv2.BitwiseAnd(matImage, horizontal, matImage);
            }

        }

        [ComImport]
        [Guid("5B0D3235-4DBA-4D44-865E-8F1D0E4FD04D")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        unsafe interface IMemoryBufferByteAccess
        {
            void GetBuffer(out byte* buffer, out uint capacity);
        }

        public unsafe Mat SoftwareBitmapToMat(SoftwareBitmap softwareBitmap)
        {
            using (BitmapBuffer buffer = softwareBitmap.LockBuffer(BitmapBufferAccessMode.Write))
            {
                using (var reference = buffer.CreateReference())
                {
                    ((IMemoryBufferByteAccess)reference).GetBuffer(out var dataInBytes, out var capacity);

                    Mat outputMat = new Mat(softwareBitmap.PixelHeight, softwareBitmap.PixelWidth, MatType.CV_8UC4, (IntPtr)dataInBytes);
                    return outputMat;
                }
            }
        }

        public unsafe void MatToSoftwareBitmap(Mat input, SoftwareBitmap output)
        {
            using (BitmapBuffer buffer = output.LockBuffer(BitmapBufferAccessMode.ReadWrite))
            {
                using (var reference = buffer.CreateReference())
                {
                    ((IMemoryBufferByteAccess)reference).GetBuffer(out var dataInBytes, out var capacity);
                    BitmapPlaneDescription bufferLayout = buffer.GetPlaneDescription(0);

                    for (int i = 0; i < bufferLayout.Height; i++)
                    {
                        for (int j = 0; j < bufferLayout.Width; j++)
                        {
                            dataInBytes[bufferLayout.StartIndex + bufferLayout.Stride * i + 4 * j + 0] =
                                input.DataPointer[bufferLayout.StartIndex + bufferLayout.Stride * i + 4 * j + 0];
                            dataInBytes[bufferLayout.StartIndex + bufferLayout.Stride * i + 4 * j + 1] =
                                input.DataPointer[bufferLayout.StartIndex + bufferLayout.Stride * i + 4 * j + 1];
                            dataInBytes[bufferLayout.StartIndex + bufferLayout.Stride * i + 4 * j + 2] =
                                input.DataPointer[bufferLayout.StartIndex + bufferLayout.Stride * i + 4 * j + 2];
                            dataInBytes[bufferLayout.StartIndex + bufferLayout.Stride * i + 4 * j + 3] =
                                input.DataPointer[bufferLayout.StartIndex + bufferLayout.Stride * i + 4 * j + 3];
                        }
                    }
                }
            }
        }
    }
}
