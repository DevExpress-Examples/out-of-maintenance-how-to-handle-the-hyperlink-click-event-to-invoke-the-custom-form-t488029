Imports DevExpress.Office.Utils
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports DXRichEditHyperlinkHandling.Forms
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows
Imports DevExpress.Office.Layout
Imports DevExpress.Xpf.Core
Imports System.Windows.Forms

Namespace DXRichEditHyperlinkHandling
    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Partial Public Class MainWindow
        Inherits DevExpress.Xpf.Ribbon.DXRibbonWindow

        #Region "#DataList"
        Private Shared products As List(Of String) = CreateProducts()
        Private Shared Function CreateProducts() As List(Of String)
            Dim result As New List(Of String)()
            result.Add("XtraScheduler™ Suite")
            result.Add("XtraRichEdit™ Suite")
            result.Add("XtraSpellChecker™")
            result.Add("XtraReports™ Suite")
            result.Add("XtraGrid™ Suite")
            result.Add("XtraPivotGrid™ Suite")
            result.Add("XtraTreeList™ Suite")
            result.Add("XtraGauges™ Suite")
            result.Add("XtraWizard™ Control")
            result.Add("XtraVerticalGrid™ Suite")
            result.Add("XtraCharts™ Suite")
            result.Add("XtraLayoutControl™ Suite")
            result.Add("XtraNavBar™")
            result.Add("XtraEditors™ Library")
            result.Add("XtraPrinting™ Library")
            Return result
        End Function
        #End Region ' #DataList
        Private activeWindow As FloatingContainer

        Public Sub New()
            InitializeComponent()
            richEditControl1.LoadDocument("HyperlinkClickHandling.rtf")
            richEditControl1.Options.Hyperlinks.ModifierKeys = Keys.None
        End Sub
        #Region "#HyperlinkClickEvent"
        Private Sub richEditControl1_HyperlinkClick(ByVal sender As Object, ByVal e As HyperlinkClickEventArgs)
            If e.ModifierKeys <> Me.richEditControl1.Options.Hyperlinks.ModifierKeys Then
                Return
            End If

            'Initialize the custom form            
            Dim control As New SelectProductForm(products)

            'Subscribe it to the OnCommit event
            AddHandler control.Commit, AddressOf OnProductFormCommit

            'Connect the form with the hyperlink range
            control.Range = e.Hyperlink.Range

            'Associate the form with the FloatingContainer instance
            Dim container As FloatingContainer = FloatingContainerFactory.Create(FloatingMode.Window)

            control.OwnerWindow = container
            container.Content = control
            container.Owner = Me.richEditControl1
            DirectCast(Me.richEditControl1, ILogicalOwner).AddChild(container)

            'Set the form's location and size
            container.SizeToContent = SizeToContent.WidthAndHeight
            container.ContainerStartupLocation = WindowStartupLocation.Manual
            container.FloatLocation = GetFormLocation()
            container.IsOpen = True
            Me.activeWindow = container
            control.Focus()

            e.Handled = True
        End Sub
        #End Region ' #HyperlinkClickEvent

        #Region "#GetFormLocation"
        Private Function GetFormLocation() As System.Windows.Point
            'Retrive the caret position
            Dim position As DocumentPosition = Me.richEditControl1.Document.CaretPosition
            Dim rect As Rectangle = Me.richEditControl1.GetBoundsFromPosition(position)

            'Set the startup location relative to the retrieved position
            'within the application bounds
            Dim richViewBounds As Rectangle = GetRichEditViewBounds()
            Dim location As New System.Drawing.Point(rect.Right - richViewBounds.X, rect.Bottom - richViewBounds.Y)
            Dim localPoint As System.Drawing.Point = Units.DocumentsToPixels(location, Me.richEditControl1.DpiX, Me.richEditControl1.DpiY)
            Return New System.Windows.Point(localPoint.X, localPoint.Y)
        End Function

        Private Function GetRichEditViewBounds() As Rectangle
            Dim documentLayoutUnitConverter As DocumentLayoutUnitConverter = New DocumentLayoutUnitDocumentConverter(Me.richEditControl1.DpiX, Me.richEditControl1.DpiY)
            Return documentLayoutUnitConverter.LayoutUnitsToDocuments(richEditControl1.ViewBounds)
        End Function
        #End Region ' #GetFormLocation

        #Region "#OnProductFormCommit"
        Private Sub OnProductFormCommit(ByVal sender As Object, ByVal e As EventArgs)
            Dim form As SelectProductForm = DirectCast(sender, SelectProductForm)

            'Retrieve the selected item value
            Dim value As String = DirectCast(form.EditValue, String)
            Dim document As Document = Me.richEditControl1.Document

            'Start the document modification
            document.BeginUpdate()

            'Replace the hyperlink range content
            'with the retireved value
            document.Replace(form.Range, value)

            'Finish the document update
            document.EndUpdate()
        End Sub
        #End Region ' #OnProductFormCommit

    End Class
End Namespace
