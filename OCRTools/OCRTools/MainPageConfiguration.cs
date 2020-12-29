using System;
using System.Collections.Generic;
using Windows.UI.Xaml.Controls;

namespace OCRTools
{
    public partial class MainPage : Page
    {
        public const string FEATURE_NAME = "OCR";

        List<Scenario> scenarios = new List<Scenario>
        {
            new Scenario() { Title="Receipts", ClassType=typeof(ScenarioPage.ReceiptOCR)},
        };
    }

    public class Scenario
    {
        public string Title { get; set; }
        public Type ClassType { get; set; }
    }
}
