using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using Windows.Storage.Streams;
using Tesseract;

namespace OCRTools.Utils
{
    class ImageProcessor
    {
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
