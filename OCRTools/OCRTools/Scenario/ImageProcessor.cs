using System;
using System.Threading.Tasks;
using Tesseract;
using Windows.Storage.Streams;

namespace OCRTools
{
    class ImageProcessor
    {
        const string tessdata = @"./tessdata";
        public static string ImageToText(byte[] imageBytes)
        {
            string text = "";

            using (var engine = new TesseractEngine(tessdata, "eng", EngineMode.Default))
            {
                Pix imagePix = Pix.LoadFromMemory(imageBytes);
                var page = engine.Process(imagePix);
                text = page.GetText();
            }
            return text;
        }

        public static async Task<byte[]> imageStreamToBytes(IRandomAccessStream streamSource)
        {
            var dr = new DataReader(streamSource.GetInputStreamAt(0));
            var bytes = new byte[streamSource.Size];
            await dr.LoadAsync((uint)streamSource.Size);
            dr.ReadBytes(bytes);
            return bytes;
        }
    }
}
