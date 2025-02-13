﻿using System.Collections.Generic;
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
            using (var page = engine.Process(imagePix)) 
            {
                text = page.GetText();
            }
            return text;
        }

        public List<System.Drawing.Rectangle> GetTextBounds(byte[] imageBytes, string text)
        {
            List<System.Drawing.Rectangle> results = new List<System.Drawing.Rectangle>();
            Pix imagePix = Pix.LoadFromMemory(imageBytes);
            using (var page = engine.Process(imagePix))
            {
                using (var iter = page.GetIterator())
                {
                    results = retrieveResults(page, text);
                }
            }
            return results;
        }

        private List<System.Drawing.Rectangle> retrieveResults(Page page, string text, float scale = 1.0f)
        {
            List<System.Drawing.Rectangle> results = new List<System.Drawing.Rectangle>();
            
            using (var pageIterator = page.GetIterator())
            {
                pageIterator.Begin();

                Rect tesseractRect;
                while (pageIterator.Next(PageIteratorLevel.TextLine))
                {
                    if (pageIterator.TryGetBoundingBox(PageIteratorLevel.TextLine, out tesseractRect))
                    {
                        string recognizedText = pageIterator.GetText(PageIteratorLevel.TextLine);
                        if (recognizedText.ToUpper().Contains(text.ToUpper()))
                        {
                            System.Drawing.Rectangle rect =
                                new System.Drawing.Rectangle(
                                    (int)(tesseractRect.X1 / scale), (int)(tesseractRect.Y1 / scale),
                                    (int)(tesseractRect.Width / scale), (int)(tesseractRect.Height / scale));
                            results.Add(rect);
                        }

                    }
                }
            }

            return results;
        }
    }
}
