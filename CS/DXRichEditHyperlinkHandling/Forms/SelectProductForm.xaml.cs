using System;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows;
using DevExpress.Xpf.Core;
using DevExpress.Utils;
using DevExpress.XtraRichEdit.API.Native;

namespace DXRichEditHyperlinkHandling.Forms
{
    /// <summary>
    /// Interaction logic for SelectProductForm.xaml
    /// </summary>
    public partial class SelectProductForm : UserControl
    {
        #region #Properties
        object fEditValue;
        DocumentRange fRange;
        public virtual object EditValue
        {
            get
            {
                return fEditValue;
            }
        }
        public DocumentRange Range
        {
            get
            {
                return fRange;
            }
            set
            {
                fRange = value;
            }
        }
        #endregion #Properties   

        #region #FloatingContainer
        FloatingContainer fOwnerWindow;
        public FloatingContainer OwnerWindow
        {
            get
            {
                return fOwnerWindow;
            }
            set
            {
                if (fOwnerWindow == value) return; fOwnerWindow = value; OnOwnerWindowChanged();
            }
        }
        #endregion #FloatingContainer

        #region #CommitEvent
        EventHandler onCommit;
        public event EventHandler Commit { add { onCommit += value; } remove { onCommit -= value; } }
        #endregion #CommitEvent

        public static readonly DependencyProperty ProductsProperty;
        static SelectProductForm()
        {
            ProductsProperty = DependencyProperty.Register("Products", typeof(List<string>), typeof(SelectProductForm), new PropertyMetadata(null));
        }
        protected void RaiseCommitEvent()
        {
            if (onCommit != null)
                onCommit(this, EventArgs.Empty);
        }
        public SelectProductForm()
        {
            this.KeyDown += PopupControlBase_KeyDown;
        }
        protected virtual void SetEditValueCore(object value)
        {
            this.fEditValue = value;
        }
        private void PopupControlBase_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
            else if (e.Key == Key.Enter)
                PerformCommit();
        }
        protected virtual void PerformCommit()
        {
            SetEditValue();
            RaiseCommitEvent();
            Close();
        }
        protected void Close()
        {
            if (OwnerWindow != null && OwnerWindow.IsOpen)
                OwnerWindow.IsOpen = false;
        }

        public SelectProductForm(List<string> list)
        {
            Guard.ArgumentNotNull(list, "list");
            Products = list;
            InitializeComponent();
            Dispatcher.BeginInvoke(new Action(() => this.listBox.Focus()));
        }

        public List<string> Products
        {
            get
            {
                return (List<string>)GetValue(ProductsProperty);
            }
            set
            {
                SetValue(ProductsProperty, value);
            }
        }

        protected void SetEditValue()
        {
            SetEditValueCore((string)this.listBox.SelectedItem);
        }

        #region #OnOwnerWindowChanged
        protected void OnOwnerWindowChanged()
        {
            if (OwnerWindow != null)
                OwnerWindow.Caption = "Select a product";
        }
        #endregion #OnOwnerWindowChanged
        private void listBox_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (this.listBox.SelectedIndex >= 0)
                PerformCommit();
        }
    }
}

