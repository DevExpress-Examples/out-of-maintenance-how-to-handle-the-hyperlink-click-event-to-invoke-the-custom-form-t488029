using DevExpress.Office.Utils;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using DXRichEditHyperlinkHandling.Forms;
using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using DevExpress.Office.Layout;
using DevExpress.Xpf.Core;
using System.Windows.Forms;
using DevExpress.Portable.Input;

namespace DXRichEditHyperlinkHandling
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : DevExpress.Xpf.Ribbon.DXRibbonWindow
    {
        #region #DataList
        static List<string> products = CreateProducts();
        static List<string> CreateProducts()
        {
            List<string> result = new List<string>();
            result.Add("XtraScheduler™ Suite");
            result.Add("XtraRichEdit™ Suite");
            result.Add("XtraSpellChecker™");
            result.Add("XtraReports™ Suite");
            result.Add("XtraGrid™ Suite");
            result.Add("XtraPivotGrid™ Suite");
            result.Add("XtraTreeList™ Suite");
            result.Add("XtraGauges™ Suite");
            result.Add("XtraWizard™ Control");
            result.Add("XtraVerticalGrid™ Suite");
            result.Add("XtraCharts™ Suite");
            result.Add("XtraLayoutControl™ Suite");
            result.Add("XtraNavBar™");
            result.Add("XtraEditors™ Library");
            result.Add("XtraPrinting™ Library");
            return result;
        }
        #endregion #DataList
        FloatingContainer activeWindow;

        public MainWindow()
        {
            InitializeComponent();
            richEditControl1.LoadDocument("HyperlinkClickHandling.rtf");
            richEditControl1.Options.Hyperlinks.ModifierKeys = PortableKeys.None;
        }
        #region #HyperlinkClickEvent
        private void richEditControl1_HyperlinkClick(object sender, HyperlinkClickEventArgs e)
        {
            if (e.ModifierKeys != this.richEditControl1.Options.Hyperlinks.ModifierKeys)
                return;

            //Initialize the custom form            
            SelectProductForm control = new SelectProductForm(products);

            //Subscribe it to the OnCommit event
            control.Commit += OnProductFormCommit;

            //Connect the form with the hyperlink range
            control.Range = e.Hyperlink.Range;

            //Associate the form with the FloatingContainer instance
            FloatingContainer container = FloatingContainerFactory.Create(FloatingMode.Window);

            control.OwnerWindow = container;
            container.Content = control;
            container.Owner = this.richEditControl1;
            ((ILogicalOwner)this.richEditControl1).AddChild(container);
            
            //Set the form's location and size
            container.SizeToContent = SizeToContent.WidthAndHeight;
            container.ContainerStartupLocation = WindowStartupLocation.Manual;
            container.FloatLocation = GetFormLocation();
            container.IsOpen = true;
            this.activeWindow = container;
            control.Focus();

            e.Handled = true;
        }
        #endregion #HyperlinkClickEvent

        #region #GetFormLocation
        System.Windows.Point GetFormLocation()
        {
            //Retrive the caret position
            DocumentPosition position = this.richEditControl1.Document.CaretPosition;
            Rectangle rect = this.richEditControl1.GetBoundsFromPosition(position);
            
            //Set the startup location relative to the retrieved position
            //within the application bounds
            Rectangle richViewBounds = GetRichEditViewBounds();
            System.Drawing.Point location = new System.Drawing.Point(rect.Right - richViewBounds.X, rect.Bottom - richViewBounds.Y);
            System.Drawing.Point localPoint = Units.DocumentsToPixels(location, this.richEditControl1.DpiX, this.richEditControl1.DpiY);
            return new System.Windows.Point(localPoint.X, localPoint.Y);
        }

        Rectangle GetRichEditViewBounds()
        {
            DocumentLayoutUnitConverter documentLayoutUnitConverter = new DocumentLayoutUnitDocumentConverter(this.richEditControl1.DpiX, this.richEditControl1.DpiY);
            return documentLayoutUnitConverter.LayoutUnitsToDocuments(richEditControl1.ViewBounds);
        }
        #endregion #GetFormLocation

        #region #OnProductFormCommit
        void OnProductFormCommit(object sender, EventArgs e)
        {
            SelectProductForm form = (SelectProductForm)sender;
            
            //Retrieve the selected item value
            string value = (string)form.EditValue;
            Document document = this.richEditControl1.Document;

            //Start the document modification
            document.BeginUpdate();
            
            //Replace the hyperlink range content
            //with the retireved value
            document.Replace(form.Range, value);

            //Finish the document update
            document.EndUpdate();
        }
        #endregion #OnProductFormCommit

    }
}
