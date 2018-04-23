Imports System
Imports System.Collections.Generic
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Windows
Imports DevExpress.Xpf.Core
Imports DevExpress.Utils
Imports DevExpress.XtraRichEdit.API.Native

Namespace DXRichEditHyperlinkHandling.Forms
    ''' <summary>
    ''' Interaction logic for SelectProductForm.xaml
    ''' </summary>
    Partial Public Class SelectProductForm
        Inherits UserControl

        #Region "#Properties"
        Private fEditValue As Object
        Private fRange As DocumentRange
        Public Overridable ReadOnly Property EditValue() As Object
            Get
                Return fEditValue
            End Get
        End Property
        Public Property Range() As DocumentRange
            Get
                Return fRange
            End Get
            Set(ByVal value As DocumentRange)
                fRange = value
            End Set
        End Property
        #End Region ' #Properties   

        #Region "#FloatingContainer"
        Private fOwnerWindow As FloatingContainer
        Public Property OwnerWindow() As FloatingContainer
            Get
                Return fOwnerWindow
            End Get
            Set(ByVal value As FloatingContainer)
                If fOwnerWindow Is value Then
                    Return
                End If
                fOwnerWindow = value
                OnOwnerWindowChanged()
            End Set
        End Property
        #End Region ' #FloatingContainer

        #Region "#CommitEvent"
        Private onCommit As EventHandler
        Public Custom Event Commit As EventHandler
            AddHandler(ByVal value As EventHandler)
                onCommit = DirectCast(System.Delegate.Combine(onCommit, value), EventHandler)
            End AddHandler
            RemoveHandler(ByVal value As EventHandler)
                onCommit = DirectCast(System.Delegate.Remove(onCommit, value), EventHandler)
            End RemoveHandler
            RaiseEvent(ByVal sender As System.Object, ByVal e As System.EventArgs)
                If onCommit IsNot Nothing Then
                    For Each d As EventHandler In onCommit.GetInvocationList()
                        d.Invoke(sender, e)
                    Next d
                End If
            End RaiseEvent
        End Event
        #End Region ' #CommitEvent

        Public Shared ReadOnly ProductsProperty As DependencyProperty
        Shared Sub New()
            ProductsProperty = DependencyProperty.Register("Products", GetType(List(Of String)), GetType(SelectProductForm), New PropertyMetadata(Nothing))
        End Sub
        Protected Sub RaiseCommitEvent()
            If onCommit IsNot Nothing Then
                onCommit(Me, EventArgs.Empty)
            End If
        End Sub
        Public Sub New()
            AddHandler Me.KeyDown, AddressOf PopupControlBase_KeyDown
        End Sub
        Protected Overridable Sub SetEditValueCore(ByVal value As Object)
            Me.fEditValue = value
        End Sub
        Private Sub PopupControlBase_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            If e.Key = Key.Escape Then
                Close()
            ElseIf e.Key = Key.Enter Then
                PerformCommit()
            End If
        End Sub
        Protected Overridable Sub PerformCommit()
            SetEditValue()
            RaiseCommitEvent()
            Close()
        End Sub
        Protected Sub Close()
            If OwnerWindow IsNot Nothing AndAlso OwnerWindow.IsOpen Then
                OwnerWindow.IsOpen = False
            End If
        End Sub

        Public Sub New(ByVal list As List(Of String))
            Guard.ArgumentNotNull(list, "list")
            Products = list
            InitializeComponent()
            Dispatcher.BeginInvoke(New Action(Function() Me.listBox.Focus()))
        End Sub

        Public Property Products() As List(Of String)
            Get
                Return DirectCast(GetValue(ProductsProperty), List(Of String))
            End Get
            Set(ByVal value As List(Of String))
                SetValue(ProductsProperty, value)
            End Set
        End Property

        Protected Sub SetEditValue()
            SetEditValueCore(CStr(Me.listBox.SelectedItem))
        End Sub

        #Region "#OnOwnerWindowChanged"
        Protected Sub OnOwnerWindowChanged()
            If OwnerWindow IsNot Nothing Then
                OwnerWindow.Caption = "Select a product"
            End If
        End Sub
        #End Region ' #OnOwnerWindowChanged
        Private Sub listBox_MouseLeftButtonUp(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
            If Me.listBox.SelectedIndex >= 0 Then
                PerformCommit()
            End If
        End Sub
    End Class
End Namespace

