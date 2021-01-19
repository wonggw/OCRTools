using System;
using System.Collections.Generic;
using System.IO;
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

                //byte[] imageBytes = await Utils.ImageParser.ImageStreamToBytes(stream);
                SoftwareBitmap bitmap = await Utils.ImageParser.ImageStreamToSoftwareBitmap(stream);

                //SoftwareBitmap output = Utils.ImageProcessor.testing(bitmap);
                var processImage = new Utils.ImageProcessor(bitmap);
                processImage.Denoising();
                processImage.GaussianBlur(3);
                processImage.Sharpen();
                //processImage.Denoising();
                processImage.Otsu();
                //SoftwareBitmapSource src = await Utils.ImageParser.SoftwareBitmapToSoftwareBitmapSource(processImage.GetSoftwareBitmap());
                byte[] imageByte = await Utils.ImageParser.BitmapToByte(processImage.GetSoftwareBitmap());
                WriteableBitmap src = Utils.ImageParser.SoftwareBitmapToWriteableBitmap(processImage.GetSoftwareBitmap());
                List<System.Drawing.Rectangle> boundBoxes = tesseractTools.GetTextBounds(imageByte, "Description");
                if (boundBoxes.Count != 0)
                {
                    foreach (System.Drawing.Rectangle rect in boundBoxes)
                    {
                        src.DrawRectangle(rect.X, rect.Y, rect.Right, rect.Bottom, Color.FromArgb(255, 255, 0, 0));
                    }
                    Output.Source = src;
                }
                else
                //{
                //    src = src.Crop(0, 350, src.PixelWidth, 100);
                //    WriteableBitmap invertedSrc = src.Invert();
                //    byte[] imageByte = await Utils.ImageProcessor.EncodeJpeg(invertedSrc);
                //    ImageText.Text = tesseractTools.ImageToText(imageByte);
                //    List<System.Drawing.Rectangle> boundBox = tesseractTools.GetTextBounds(imageByte, "Description");
                //    foreach (System.Drawing.Rectangle rect in boundBox)
                //    {
                //        invertedSrc.DrawRectangle(rect.X, rect.Y, rect.Right, rect.Bottom, Color.FromArgb(255, 255, 0, 0));
                //    }
                //    Output.Source = invertedSrc;
                //}

                Output.Source = src;
                //ImageText.Text = tesseractTools.ImageToText(imageByte);

            }
            ProgressControl.Visibility = Visibility.Collapsed;
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
