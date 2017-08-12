'Option Strict On
''Option Explicit On
Imports Reader
Imports System.Reflection
Imports System
Imports Reader.CodeGen
Imports Reader.Collections
Imports Reader.CodeModel
Imports Reader.ILConvertor

Public Class ReflectorUIForm
    Private mBase As MethodBase
    Private Ofd As New OpenFileDialog()
    Private Loader As AssemblyLoader
    Private FullWidth As Integer
    Private RTBox1 As New ReadOnlyRichTextBox
    Private tv As TreeLoader
    Private dlllist As New DllPathSerializer
    Private Sub ReflectorUIForm_Load(ByVal Sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        RTBox1.ScrollBars = RichTextBoxScrollBars.ForcedBoth
        RTBox1.WordWrap = False
        RTBox1.BackColor = Color.White
        RTBox1.Dock = DockStyle.Fill
        RTBox1.Font = New Font("Courier New", 10.0, FontStyle.Regular, GraphicsUnit.Point, 0)
        Me.SplitContainer1.Panel2.Controls.Add(RTBox1)
        ' dlllist.Load()
        ' If dlllist.AssemblyPath.Count = 0 Then Return
        'For Each filename As String In dlllist.AssemblyPath
        'tv = New TreeLoader(filename)
        'tv.LoadTree(TreeView1)
        'Next

    End Sub
    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click

        Ofd.Filter = "PE File(*.dll;*.exe;*.mod;*.mdl)|*dll;*.exe;*.mod;*.mdl"
        Ofd.Title = "Open Assembly"

        If Ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            tv = New TreeLoader(Ofd.FileName)
            tv.LoadTree(TreeView1)
            dlllist.AssemblyPath.Add(Ofd.FileName)
        End If

    End Sub
    Private Sub TreeView1_NodeMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        FullWidth = Me.Width
        Me.Width = SplitContainer1.Panel1.Width
        SplitContainer1.Panel2Collapsed = True

    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect

        Dim count As Integer = 0
        Dim minfo As MethodInfo
        Dim cinfo As ConstructorInfo
        Dim StringToken As String = String.Empty
        If e.Node.Tag Is Nothing Then Return
        Me.Width = FullWidth
        SplitContainer1.Panel2Collapsed = False

        If TypeOf e.Node.Tag Is MethodInfo Then
            minfo = DirectCast(e.Node.Tag, MethodInfo)
            mBase = MethodBase.GetMethodFromHandle(minfo.MethodHandle, minfo.DeclaringType.TypeHandle)
            'Dim rtb As New RTBLoader(RTBox1, Ofd.FileName, mBase, minfo)
            'rtb.RTBLoadMethodDefinition()
            Dim generator As New MethodWriter(Ofd.FileName, mBase)
            RTBox1.Clear()
            For Each Str As String In generator.WriteLocals()
                RTBox1.AppendText(Str & vbNewLine)
            Next
            generator.MethodBuilder()
            Dim s As String
            For Each line As IOffset In generator.CodeLines
                If line.GetType.Name = "BinaryExpression" Then
                    Dim bx As BinaryExpression = TryCast(line, BinaryExpression)
                    s = bx.ToString1
                    RTBox1.AppendText(s & vbNewLine)
                ElseIf line.GetType.Name = "MethodExpression" Then
                    Dim mx As MethodExpression = TryCast(line, MethodExpression)
                    s = mx.ToString1
                    RTBox1.AppendText(s & vbNewLine)
                ElseIf line.GetType.Name = "ConvertExpression" Then
                    Dim cx As ConvertExpression = TryCast(line, ConvertExpression)
                    s = cx.ToString1
                    RTBox1.AppendText(s & vbNewLine)
                ElseIf line.GetType.Name = "CastExpression" Then
                    Dim castx As CastExpression = TryCast(line, CastExpression)
                    s = castx.ToString1
                    RTBox1.AppendText(s & vbNewLine)
                ElseIf line.GetType.Name = "CreateInstanceExpression" Then
                    Dim crix As CreateInstanceExpression = TryCast(line, CreateInstanceExpression)
                    s = crix.ToString1
                    RTBox1.AppendText(s & vbNewLine)
                End If

                Format(RTBox1, count, "Double")
                Format(RTBox1, count, "Boolean")
                Format(RTBox1, count, "Dim")
                Format(RTBox1, count, "Me")
                Format(RTBox1, count, "If")
                Format(RTBox1, count, "End")
                Format(RTBox1, count, "Then")
                Format(RTBox1, count, "String")
                Format(RTBox1, count, "New")
                Format(RTBox1, count, "Or")
                Format(RTBox1, count, "Do")
                Format(RTBox1, count, "Loop")
                Format(RTBox1, count, "While")
                FormatString(RTBox1, s, count)
                count += s.Length
            Next

        ElseIf TypeOf e.Node.Tag Is ConstructorInfo Then
            cinfo = DirectCast(e.Node.Tag, ConstructorInfo)
            mBase = MethodBase.GetMethodFromHandle(cinfo.MethodHandle, cinfo.DeclaringType.TypeHandle)
            ''Dim rtb As New RTBLoader(RTBox1, Ofd.FileName, mBase, Nothing)
            ''rtb.RTBLoadMethodDefinition()
        End If

    End Sub
    Private Sub TreeView1_AfterExpand(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterExpand

        If TypeOf e.Node.Tag Is Type Then
            e.Node.Nodes.Clear()
            tv.PopulateTypeNode(e.Node)
        End If

    End Sub
    Private Sub TreeView1_AfterCollapse(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterCollapse

    End Sub

    Private Sub FileToolStripMenuItem_MouseEnter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileToolStripMenuItem.MouseEnter
        FileToolStripMenuItem.ShowDropDown()
    End Sub
    Private Sub ReflectorUIForm_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        dlllist.Save()
        TreeView1.Nodes.Clear()
    End Sub
    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        TreeView1.Nodes.Clear()
        dlllist.AssemblyPath.Clear()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
    Private Sub FormatString(ByVal rtb As RichTextBox, ByVal text As String, ByVal start As Integer)
        Dim t() As String = text.Split(ChrW(34).ToString.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        Dim x As Integer = rtb.Find(ChrW(34).ToString, start, RichTextBoxFinds.None)
        If x >= 0 Then
            rtb.SelectionStart = x
            If t.Length > 1 Then
                rtb.SelectionLength = t(1).Length + 2
            End If
            rtb.SelectionColor = Color.Maroon
        End If
    End Sub
    Private Sub Format(ByVal rtb As RichTextBox, ByVal start As Integer, ByVal text As String)
        Dim x As Integer = rtb.Find(text, start, RichTextBoxFinds.MatchCase Or RichTextBoxFinds.WholeWord)
        If x >= 0 Then
            rtb.SelectionStart = x
            rtb.SelectionLength = text.Length
            rtb.SelectionColor = Color.Blue
        End If
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        RTBox1.Clear()
        RTBox1.BackColor = Color.Beige
        If TryCast(sender, ToolStripComboBox).SelectedIndex = 0 Then
            For Each ins As ILInstruction In New ILReader(mBase)
                RTBox1.AppendText(ins.ToString & vbNewLine)
            Next
        End If
    End Sub

    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub
End Class
