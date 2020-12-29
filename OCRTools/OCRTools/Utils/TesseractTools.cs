using System.Collections.Generic;
using Tesseract;

namespace OCRTools.Utils
{
    class TesseractTools
    {
        const string tessdata = @".\tessdata";
        private TesseractEngine engine;

        public TesseractTools()
        {
            engine = new TesseractEngine(tessdata, "eng", EngineMode.Default);
        }

        public string ImageToText(byte[] imageBytes)
        {
            string text = "";
            Pix imagePix = Pix.LoadFromMemory(imageBytes);
            var page = engine.Process(imagePix);
            text = page.GetText();       
            return text;
        }

        public List<System.Drawing.Rectangle> GetTextBounds(byte[] imageBytes, string text)
        {
            List<System.Drawing.Rectangle> results = new List<System.Drawing.Rectangle>();

            Pix imagePix = Pix.LoadFromMemory(imageBytes);
            var page = engine.Process(imagePix);
            using (var iter = page.GetIterator())
            {
                results = retrieveResults(page, text);
            }
            return results;
        }


        private List<System.Drawing.Rectangle> retrieveResults(Page page, string text, float scale = 1.0f)
        {
            List<System.Drawing.Rectangle> results = new List<System.Drawing.Rectangle>();

            using (var iter = page.GetIterator())
            {
                iter.Begin();

                Rect r;

                while (iter.Next(PageIteratorLevel.TextLine))
                {
                    if (iter.TryGetBoundingBox(PageIteratorLevel.TextLine, out r))
                    {
                        string str = iter.GetText(PageIteratorLevel.TextLine);
                        if (str.ToUpper().Contains(text.ToUpper()))
                        {
                            System.Drawing.Rectangle rect =
                                new System.Drawing.Rectangle(
                                    new System.Drawing.Point((int)(r.X1 / scale), (int)(r.Y1 / scale)),
                                    new System.Drawing.Size((int)(r.Width / scale), (int)(r.Height / scale)));
                            results.Add(rect);
                        }

                    }
                }
            }

            return results;
        }
    }
}
