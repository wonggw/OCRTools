using System;
using System.Runtime.InteropServices;
using Windows.Graphics.Imaging;
using OpenCvSharp;

namespace OCRTools.Utils
{
    class ImageProcessor
    {
        private int imageWidth { get; set; }
        private int imageHeight { get; set; }
        private BitmapPixelFormat bitmapPixelFormat { get; set; }
        private BitmapAlphaMode bitmapAlphaMode  { get; set; }
        private Mat matImage { get; set; }

        public ImageProcessor(SoftwareBitmap softwareBitmap)
        {
            imageWidth = softwareBitmap.PixelWidth;
            imageHeight = softwareBitmap.PixelHeight;
            bitmapPixelFormat = softwareBitmap.BitmapPixelFormat;
            bitmapAlphaMode = softwareBitmap.BitmapAlphaMode;
            matImage = SoftwareBitmapToMat(softwareBitmap);
        }

        public SoftwareBitmap GetSoftwareBitmap()
        {
            SoftwareBitmap softwareBitmap = new SoftwareBitmap(bitmapPixelFormat, imageWidth, imageHeight, bitmapAlphaMode);
            MatToSoftwareBitmap(matImage, softwareBitmap);
            return softwareBitmap;
        }

        public void Denoising()
        {
            Cv2.FastNlMeansDenoising(matImage, matImage);
        }

        public void GaussianBlur(int kernelSize)
        {
            Cv2.GaussianBlur(matImage, matImage, new Size(kernelSize, kernelSize), sigmaX: 0);
        }

        public void Sharpen()
        {
            using (Mat sharpen = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            {
                double[] sharpenArray = { -1, -1 -1,
                                          -1, 9, -1,
                                          -1, -1,-1 };
                var sharpeKernel = new Mat(rows: 3, cols: 3, type: MatType.CV_64FC1, data: sharpenArray);
                Cv2.Filter2D(matImage, sharpen, MatType.CV_64FC1, sharpeKernel);
                sharpen.ConvertTo(matImage, MatType.CV_8UC4);
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
        public void testing()
        {
            using (Mat test = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            using (Mat intermediate = new Mat(matImage.Rows, matImage.Cols, MatType.CV_8UC4))
            {
                Cv2.CvtColor(matImage, test, ColorConversionCodes.BGRA2GRAY);
                Cv2.Threshold(test, intermediate, 127, 255, ThresholdTypes.Otsu);
                int an = 2;
                //var element = Cv2.GetStructuringElement(
                //                MorphShapes.Cross,
                //                new Size(an * 2 + 1, an * 2 + 1),
                //                new Point(an, an));

                var element = Mat.Ones(an, an);
                //Cv2.Threshold(intermediate, test, 70, 255, ThresholdTypes.BinaryInv);
                //Cv2.MorphologyEx(test, intermediate, MorphTypes.Erode, element);
                //Cv2.MorphologyEx(test, intermediate, MorphTypes., element);
                //Cv2.Threshold(test, intermediate, 70, 255, ThresholdTypes.BinaryInv);
                Cv2.CvtColor(intermediate, matImage, ColorConversionCodes.GRAY2BGRA);
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
