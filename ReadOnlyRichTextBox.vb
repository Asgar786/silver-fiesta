Imports System.Runtime.InteropServices
Imports System

Public Class ReadOnlyRichTextBox
    Inherits RichTextBox
    Public Sub New()
        Me.ReadOnly = True
        Me.SetStyle(ControlStyles.CacheText, True)
        'Me.SetStyle(ControlStyles.Opaque, True)
        ''Me.SetStyle(ControlStyles.ResizeRedraw, False)
        ''Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        ' ''Me.SetStyle(ControlStyles.AllPaintingInWmPaint, False)
    End Sub
    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            'cp.ExStyle = &H100000
            Return cp
        End Get
    End Property
    Private Declare Function HideCaret Lib "user32.dll" (<[In]()> ByVal hwnd As System.IntPtr) As Boolean
    Declare Auto Function SendMessage Lib "User32.dll" (<[In]()> ByVal hWnd As IntPtr, <[In]()> ByVal Msg As UInt32, <[In]()> ByVal wParam As IntPtr, <[In]()> ByVal lParam As IntPtr) As Int32
    Protected Overrides Function ProcessCmdKey(ByRef m As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
        If m.Msg = 256 Then
            Return True
        End If
        Return MyBase.ProcessCmdKey(m, keyData)
    End Function
    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        HideCaret(MyBase.Handle)
        MyBase.OnMouseDown(e)
    End Sub
    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
        HideCaret(MyBase.Handle)
        MyBase.OnTextChanged(e)
    End Sub
End Class
