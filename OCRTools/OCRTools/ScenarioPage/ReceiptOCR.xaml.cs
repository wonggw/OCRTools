using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Foundation;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.UI;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Media.Imaging;

namespace OCRTools.ScenarioPage
{
    public sealed partial class ReceiptOCR : Page
    {
        private MainPage rootPage = MainPage.Current;
        private PdfDocument pdfDocument;
        private Utils.TesseractTools tesseractTools;

        const int WrongPassword = unchecked((int)0x8007052b); // HRESULT_FROM_WIN32(ERROR_WRONG_PASSWORD)
        const int GenericFail = unchecked((int)0x80004005);   // E_FAIL

        public ReceiptOCR()
        {
            this.InitializeComponent();
            tesseractTools = new Utils.TesseractTools();
        }

        private async void LoadDocument(object sender, RoutedEventArgs args)
        {
            LoadButton.IsEnabled = false;

            pdfDocument = null;
            Output.Source = null;
            PageNumberBox.Text = "1";
            ImageText.Text = "";
            RenderingPanel.Visibility = Visibility.Collapsed;

            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".pdf");
            StorageFile file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                ProgressControl.Visibility = Visibility.Visible;
                try
                {
                    pdfDocument = await PdfDocument.LoadFromFileAsync(file, PasswordBox.Password);
                }
                catch (Exception ex)
                {
                    switch (ex.HResult)
                    {
                        case WrongPassword:
                            rootPage.NotifyUser("Document is password-protected and password is incorrect.", NotifyType.ErrorMessage);
                            break;

                        case GenericFail:
                            rootPage.NotifyUser("Document is not a valid PDF.", NotifyType.ErrorMessage);
                            break;

                        default:
                            // File I/O errors are reported as exceptions.
                            rootPage.NotifyUser(ex.Message, NotifyType.ErrorMessage);
                            break;
                    }
                }

                if (pdfDocument != null)
                {
                    RenderingPanel.Visibility = Visibility.Visible;
                    if (pdfDocument.IsPasswordProtected)
                    {
                        rootPage.NotifyUser("Document is password protected.", NotifyType.StatusMessage);
                    }
                    else
                    {
                        rootPage.NotifyUser("Document is not password protected.", NotifyType.StatusMessage);
                    }
                    PageCountText.Text = pdfDocument.PageCount.ToString();
                }
                ProgressControl.Visibility = Visibility.Collapsed;
            }
            LoadButton.IsEnabled = true;
        }

        private async void ViewPage(object sender, RoutedEventArgs args)
        {
            rootPage.NotifyUser("", NotifyType.StatusMessage);

            uint pageNumber;
            if (!uint.TryParse(PageNumberBox.Text, out pageNumber) || (pageNumber < 1) || (pageNumber > pdfDocument.PageCount))
            {
                rootPage.NotifyUser("Invalid page number.", NotifyType.ErrorMessage);
                return;
            }

            Output.Source = null;
            ProgressControl.Visibility = Visibility.Visible;

            // Convert from 1-based page number to 0-based page index.
            uint pageIndex = pageNumber - 1;
            using (PdfPage page = pdfDocument.GetPage(pageIndex))
            {
                var stream = new InMemoryRandomAccessStream();
                switch (Options.SelectedIndex)
                {
                    // View actual size.
                    case 0:
                        await page.RenderToStreamAsync(stream);
                        break;

                    // View half size on beige background.
                    case 1:
                        var options1 = new PdfPageRenderOptions();
                        options1.BackgroundColor = Windows.UI.Colors.Beige;
                        options1.DestinationHeight = (uint)(page.Size.Height / 2);
                        options1.DestinationWidth = (uint)(page.Size.Width / 2);
                        await page.RenderToStreamAsync(stream, options1);
                        break;

                    // Crop to center.
                    case 2:
                        var options2 = new PdfPageRenderOptions();
                        var rect = page.Dimensions.TrimBox;
                        options2.SourceRect = new Rect(rect.X + rect.Width / 4, rect.Y + rect.Height / 4, rect.Width / 2, rect.Height / 2);
                        await page.RenderToStreamAsync(stream, options2);
                        break;
                }

                System.Drawing.Rectangle descriptionBoundingBox = await GetDescriptionBoundingBox(stream);

                SoftwareBitmap inputBitmap = await Utils.ImageParser.ImageStreamToSoftwareBitmap(stream);
                var processImage = new Utils.ImageProcessor(inputBitmap);
                int croppedHeight = (inputBitmap.PixelHeight - descriptionBoundingBox.Bottom - 430);
                if (croppedHeight <= 0) croppedHeight = (int)((inputBitmap.PixelHeight - descriptionBoundingBox.Bottom) * 0.53);
                const int scale = 4;
                processImage.Crop(0, descriptionBoundingBox.Bottom, inputBitmap.PixelWidth, croppedHeight);
                processImage.Sharpen(1);
                processImage.BilateralFilter(3);
                processImage.Resize(scale);
                processImage.BitwiseNot();
                processImage.RemoveHorizontalLines();
                processImage.RemoveVerticalLines();
                processImage.GrayToZero();
                processImage.MorphologyExOpen(2, 1);
                processImage.MorphologyExDilate(2, 1);
                processImage.Resize(0.8);
                processImage.Invert();
                processImage.MorphologyExErode(2, 1);
                byte[] outputByte = await Utils.ImageParser.BitmapToByte(processImage.GetSoftwareBitmap());
                ImageText.Text = tesseractTools.ImageToText(outputByte);
                processImage.Resize(0.5);
                WriteableBitmap src = Utils.ImageParser.SoftwareBitmapToWriteableBitmap(processImage.GetSoftwareBitmap());
                Output.Source = src;

            }
            ProgressControl.Visibility = Visibility.Collapsed;
        }

        private async Task<System.Drawing.Rectangle> GetDescriptionBoundingBox(IRandomAccessStream streamImage)
        {
            SoftwareBitmap softwareBitmap = await Utils.ImageParser.ImageStreamToSoftwareBitmap(streamImage);

            const int scale = 2;
            var processImage = new Utils.ImageProcessor(softwareBitmap);
            processImage.Sharpen(1);
            processImage.Denoising();
            processImage.BilateralFilter(3);
            processImage.Resize(scale);
            processImage.Sharpen(3);
            processImage.MorphologyExErode(3);
            //WriteableBitmap src = Utils.ImageParser.SoftwareBitmapToWriteableBitmap(processImage.GetSoftwareBitmap());
            //src = src.Resize(src.PixelWidth / scale, src.PixelHeight / scale, WriteableBitmapExtensions.Interpolation.Bilinear);
            byte[] outputByte = await Utils.ImageParser.BitmapToByte(processImage.GetSoftwareBitmap());
            List<System.Drawing.Rectangle> boundBoxes = tesseractTools.GetTextBounds(outputByte, "description");
            if (boundBoxes.Count != 0)
            {
                //src.DrawRectangle(boundBoxes[0].X / scale, boundBoxes[0].Y / scale, boundBoxes[0].Right / scale, boundBoxes[0].Bottom / scale, Color.FromArgb(255, 255, 0, 0));
                //Output.Source = src;
                return new System.Drawing.Rectangle(boundBoxes[0].X / scale, boundBoxes[0].Y / scale, boundBoxes[0].Width / scale, boundBoxes[0].Height / scale);
            }
               
            else
            {
                //Output.Source = src;
                return new System.Drawing.Rectangle();
            }
        }

        private async void CreateDocument(object sender, RoutedEventArgs args)
        {
            FileSavePicker savePicker = new FileSavePicker();
            savePicker.SuggestedStartLocation = PickerLocationId.Desktop;
            savePicker.FileTypeChoices.Add("Excel Workbook", new List<string>() {".xlsx"});
            savePicker.SuggestedFileName = "Book1";
            StorageFile file = await savePicker.PickSaveFileAsync();

            if (file == null) return;
            ImageText.Text = file.Path;
            using (Stream stream = await file.OpenStreamForWriteAsync())
            {
                Utils.ExcelProcessor.CreateSpreadsheetWorkbook(stream);
            }                
        }
    }
}
