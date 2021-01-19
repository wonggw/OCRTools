using System;
using System.Threading.Tasks;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Storage.Streams;
using Windows.UI.Xaml.Media.Imaging;
using Windows.Graphics.Imaging;

namespace OCRTools.Utils
{
    class ImageParser
    {
        public static async Task<byte[]> ImageStreamToBytes(IRandomAccessStream streamSource)
        {
            var dataStream = new DataReader(streamSource.GetInputStreamAt(0));
            var bytes = new byte[streamSource.Size];
            await dataStream.LoadAsync((uint)streamSource.Size);
            dataStream.ReadBytes(bytes);
            return bytes;
        }

        public static async Task<SoftwareBitmap> ImageStreamToSoftwareBitmap(IRandomAccessStream streamSource)
        {
            BitmapDecoder decoder = await BitmapDecoder.CreateAsync(streamSource);
            return await decoder.GetSoftwareBitmapAsync();
        }

        public static async Task<WriteableBitmap> ImageStreamToWritableBitmap(IRandomAccessStream streamSource)
        {
            return await BitmapFactory.FromStream(streamSource);
        }

        public static async Task<byte[]> BitmapToByte(SoftwareBitmap softwareBitmap)
        {
            byte[] array = null;

            using (var ms = new InMemoryRandomAccessStream())
            {
                BitmapEncoder encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, ms);
                encoder.SetSoftwareBitmap(softwareBitmap);

                try
                {
                    await encoder.FlushAsync();
                }
                catch { }

                array = new byte[ms.Size];
                await ms.ReadAsync(array.AsBuffer(), (uint)ms.Size, InputStreamOptions.None);
            }

            return array;
        }

        public static async Task<SoftwareBitmapSource> SoftwareBitmapToSoftwareBitmapSource(SoftwareBitmap softwareBitmap)
        {
            SoftwareBitmapSource bitmapSource = new SoftwareBitmapSource();
            SoftwareBitmap displayableImage = SoftwareBitmap.Convert(softwareBitmap, BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
            await bitmapSource.SetBitmapAsync(displayableImage);
            return bitmapSource;
        }

        public static WriteableBitmap SoftwareBitmapToWriteableBitmap(SoftwareBitmap softwareBitmap)
        {
            WriteableBitmap writeable = new WriteableBitmap(softwareBitmap.PixelWidth, softwareBitmap.PixelHeight);
            softwareBitmap.CopyToBuffer(writeable.PixelBuffer);
            return writeable;
        }

        public static async Task<byte[]> BitmapToByte(WriteableBitmap writableBitmap)
        {
            SoftwareBitmap softwareBitmap = SoftwareBitmap.CreateCopyFromBuffer(writableBitmap.PixelBuffer, BitmapPixelFormat.Bgra8, writableBitmap.PixelWidth, writableBitmap.PixelHeight);
            byte[] array = null;

            using (var ms = new InMemoryRandomAccessStream())
            {
                BitmapEncoder encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, ms);
                encoder.SetSoftwareBitmap(softwareBitmap);

                try
                {
                    await encoder.FlushAsync();
                }
                catch { }

                array = new byte[ms.Size];
                await ms.ReadAsync(array.AsBuffer(), (uint)ms.Size, InputStreamOptions.None);
            }

            return array;
        }
    }
}
