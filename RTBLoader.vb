Imports Reader
Imports System.Reflection
Imports System

Public Class RTBLoader
    Delegate Sub Newdelegate(ByVal p1 As Integer, ByVal p2 As Integer, ByVal p3 As Color, ByVal p4 As Font)
    Private SelectionDelagate As New Newdelegate(AddressOf SelectText)
    Private _filename As String
    Private _methodbase As MethodBase
    Private _minfo As MethodInfo
    Private rtb As ReadOnlyRichTextBox
    Private Lreader As LocalReader
    Public Sub New(ByVal richtextbox1 As RichTextBox, ByVal filename As String, ByVal m_base As MethodBase, ByVal m_info As MethodInfo)
        _filename = filename
        _methodbase = m_base
        _minfo = m_info
        rtb = richtextbox1
        Lreader = New LocalReader(_filename)
    End Sub
    Public Sub RTBLoadMethodDefinition()
        rtb.Clear()
        Dim md As New MethodDescriptor(_minfo)
        Dim stringToken As String = ".method" & ChrW(32)
        stringToken = String.Format("{0}{1}", stringToken, md.Attributes)
        rtb.AppendText(stringToken)
        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.MediumBlue, rtb.Font)
        rtb.AppendText(ChrW(32))
        If Not md.ReturnType.IsPrimitive Then
            If md.ReturnType.IsClass Then
                stringToken = "class "
            ElseIf md.ReturnType.IsValueType Then
                stringToken = "valuetype "
            End If
        End If
        rtb.AppendText(stringToken)
        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.MediumBlue, rtb.Font)
        stringToken = md.ReturnType.ToString
        If Not md.ReturnType.IsGenericType Then
            stringToken = ReflectedType.RType(md.ReturnType)
        End If
        If stringToken IsNot Nothing Then
            rtb.AppendText(stringToken)
            Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.FromArgb(0, 80, 0), rtb.Font)
        End If
        rtb.AppendText(ChrW(32))
        stringToken = md.Name
        rtb.AppendText(stringToken)
        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.Black, New Font(rtb.Font, FontStyle.Bold))
        rtb.AppendText("(")
        Me.SelectText(1, 1, Color.Black, rtb.Font)
        stringToken = md.Parameters
        rtb.AppendText(stringToken)
        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.FromArgb(0, 80, 0), rtb.Font)
        rtb.AppendText(")")
        Me.SelectText(1, 1, Color.Black, rtb.Font)
        stringToken = md.ImplAttributes
        rtb.AppendText(stringToken)
        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.MediumBlue, rtb.Font)
        rtb.AppendText(vbNewLine)
        RTBLoadStackSize()
        RTBLoadLocals()
        RTBLoadILInstructions()
    End Sub
    Public Sub RTBLoadStackSize()

        Dim Body As MethodBody = _methodbase.GetMethodBody()
        If Body Is Nothing Then Return
        Dim MaxStackSize As Integer = Body.MaxStackSize
        rtb.AppendText("{")
        rtb.AppendText(vbNewLine)
        rtb.AppendText(ChrW(32) & ChrW(32))
        Dim stringToken As String = String.Format(".maxstack {0}", MaxStackSize)
        rtb.AppendText(stringToken)
        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.MediumBlue, New Font("Tahoma", 11.5, FontStyle.Regular, GraphicsUnit.World))
        rtb.AppendText(vbNewLine)
    End Sub
    Public Sub RTBLoadLocals()

        Dim stringToken As String = ".locals init"
        Lreader.GetLocals(_methodbase)

        If Lreader.Locals.Count = 0 Then
            Exit Sub
        End If

        rtb.AppendText(ChrW(32) & ChrW(32))
        rtb.AppendText(stringToken)
        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.MediumBlue, New Font("Tahoma", 11.5, FontStyle.Regular, GraphicsUnit.World))
        rtb.AppendText(" (" & vbNewLine)

        For Each var As LocalVarInfo In Lreader.Locals

            Dim index As Integer = var.LocalIndex
            rtb.AppendText(ChrW(32) & ChrW(32) & ChrW(32) & ChrW(32) & ChrW(32) & ChrW(32) & ChrW(32) & ChrW(32) & ChrW(32))
            stringToken = index.ToString
            rtb.AppendText("[")
            rtb.AppendText(stringToken)
            SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.Maroon, rtb.Font)
            rtb.AppendText("]")
            Me.SelectText(1, 1, Color.Black, rtb.Font)
            rtb.AppendText(ChrW(32))
            stringToken = ""

            If Not var.LocalType.IsPrimitive Then

                Select Case True
                    Case var.LocalType.IsClass
                        stringToken = "class"
                    Case var.LocalType.IsValueType
                        stringToken = "valuetype"
                End Select

                rtb.AppendText(stringToken)
                SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.MediumBlue, rtb.Font)
                rtb.AppendText(ChrW(32))
                stringToken = var.LocalType.Assembly.FullName.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)(0)
                rtb.AppendText("[")
                Me.SelectionDelagate.Invoke(1, 1, Color.Black, rtb.Font)
                rtb.AppendText(stringToken)
                Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.FromArgb(0, 80, 0), rtb.Font)
                rtb.AppendText("]")
                Me.SelectionDelagate.Invoke(1, 1, Color.Black, rtb.Font)

            End If

            If Not var.LocalType.IsGenericType Then
                stringToken = ReflectedType.RType(var.LocalType)
            Else
                stringToken = var.LocalType.ToString
            End If

            If stringToken IsNot Nothing Then
                rtb.AppendText(stringToken)
                Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.FromArgb(0, 80, 0), rtb.Font)
            End If

            rtb.AppendText(ChrW(32))

            stringToken = var.Name
            rtb.AppendText(stringToken)
            Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.Black, New Font(rtb.Font, FontStyle.Bold))

            Dim i As Integer = Lreader.Locals.Count - 1

            If Lreader.Locals.IndexOf(var) = Lreader.Locals.Count - 1 Then
                rtb.AppendText(")")
            Else
                rtb.AppendText(",")
            End If

            Me.SelectionDelagate.Invoke(1, 1, Color.Black, rtb.Font)
            rtb.AppendText(vbNewLine)

        Next

    End Sub
    Public Sub RTBLoadILInstructions()

        Using [Enumerator] As IEnumerator(Of ILInstruction) = New ILReader(_methodbase).GetEnumerator

            While Enumerator.MoveNext

                Dim Instruction As ILInstruction = Enumerator.Current
                Dim stringToken As String
                rtb.AppendText(ChrW(32) & ChrW(32) & ChrW(32) & ChrW(32))
                rtb.AppendText(Instruction.IlOffset)
                rtb.AppendText(ChrW(32))
                rtb.AppendText(Instruction.Code.Name)
                Me.SelectionDelagate.Invoke(Instruction.Code.Name.Length, Instruction.Code.Name.Length, Color.FromArgb(0, 80, 0), rtb.Font)
                rtb.AppendText(vbTab)
                Dim OperandTypeName As String = Instruction.GetType().Name

                Select Case OperandTypeName
                    Case "Inline8Instruction"
                        Dim LongValue As Int64 = DirectCast(Instruction, Inline8Instruction).Value
                        rtb.AppendText(LongValue.ToString)
                        Me.SelectionDelagate.Invoke(LongValue.ToString.Length, LongValue.ToString.Length, Color.Maroon, rtb.Font)
                    Case "InlinelInstruction"
                        Dim intValue As Int32 = DirectCast(Instruction, InlinelInstruction).Value
                        Dim hexstring As String = String.Format("0x{0}", Hex(intValue))
                        rtb.AppendText(hexstring)
                        Me.SelectionDelagate.Invoke(hexstring.Length, hexstring.Length, Color.Maroon, rtb.Font)
                    Case "InlineRInstruction"
                        Dim DoubleValue As Double = DirectCast(Instruction, InlineRInstruction).Value
                        rtb.AppendText(DoubleValue.ToString)
                        Me.SelectionDelagate.Invoke(DoubleValue.ToString.Length, DoubleValue.ToString.Length, Color.Maroon, rtb.Font)
                    Case "InlineSigInstruction"

                    Case "InlineTokInstruction"
                        Dim tokenvalue As MemberInfo = DirectCast(Instruction, InlineTokInstruction).Value
                        rtb.AppendText(tokenvalue.Name)
                        Me.SelectionDelagate.Invoke(tokenvalue.Name.Length, tokenvalue.Name.Length, Color.Black, rtb.Font)
                    Case "InlineTypeInstruction"
                        Dim inlinetype As Type = DirectCast(Instruction, InlineTypeInstruction).GetType
                        rtb.AppendText(inlinetype.ToString)
                        Me.SelectText(inlinetype.ToString.Length, inlinetype.ToString.Length, Color.FromArgb(0, 80, 0), rtb.Font)
                    Case "InlineVarInstruction"
                        Dim index As UInt16 = DirectCast(Instruction, InlineVarInstruction).Value
                        Dim localvarname As String = Lreader.Locals(index).Name
                        rtb.AppendText(localvarname)
                        Me.SelectionDelagate.Invoke(localvarname.Length, localvarname.Length, Color.Black, New Font(rtb.Font, FontStyle.Bold))
                    Case "InlineStringInstruction"
                        stringToken = DirectCast(Instruction, InlineStringInstruction).Value
                        stringToken = ChrW(34) & stringToken & ChrW(34)
                        rtb.AppendText(stringToken)
                        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.Maroon, rtb.Font)
                    Case "InlineSwitchInstruction"
                        Dim deltas() As String = DirectCast(Instruction, InlineSwitchInstruction).Value.ToArray
                        rtb.AppendText("(")
                        For Each Label As String In deltas
                            rtb.AppendText(Label & ", ")
                            Me.SelectionDelagate.Invoke(Label.Length - 2, Label.Length + 2, Color.Black, rtb.Font)
                        Next
                    Case "InlineBrTargetInstruction"
                        stringToken = DirectCast(Instruction, InlineBrTargetInstruction).Value
                        rtb.AppendText(stringToken)
                        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.Black, rtb.Font)
                    Case "InlineFieldInstruction"
                        stringToken = DirectCast(Instruction, InlineFieldInstruction).Value.Name
                        rtb.AppendText(stringToken)
                        Me.SelectText(stringToken.Length, stringToken.Length, Color.FromArgb(0, 80, 0), rtb.Font)
                    Case "InlineMethodInstruction"
                        Dim base As MethodBase = DirectCast(Instruction, InlineMethodInstruction).Value
                        Dim Asm As Assembly = Assembly.GetAssembly(base.DeclaringType)
                        stringToken = String.Format("[{0}]{1}::{2}", Asm.FullName.Split(","c)(0), base.DeclaringType.FullName, base.Name)
                        rtb.AppendText(stringToken)
                        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.FromArgb(0, 80, 0), rtb.Font)
                    Case "ShortInlineBrTargetInstruction"
                        stringToken = DirectCast(Instruction, ShortInlineBrTargetInstruction).Value
                        rtb.AppendText(stringToken)
                        Me.SelectionDelagate.Invoke(stringToken.Length, stringToken.Length, Color.Black, rtb.Font)
                    Case "ShortInlinelInstruction"
                        Dim value As SByte = DirectCast(Instruction, ShortInlinelInstruction).Value
                        rtb.AppendText(value.ToString)
                        Me.SelectionDelagate.Invoke(value.ToString.Length, value.ToString.Length, Color.Maroon, rtb.Font)
                    Case "ShortInlineRInstruction"
                        Dim svalue As Single = DirectCast(Instruction, ShortInlineRInstruction).Value
                        rtb.AppendText(svalue.ToString)
                        Me.SelectionDelagate.Invoke(svalue.ToString.Length, svalue.ToString.Length, Color.Maroon, rtb.Font)
                    Case "ShortInlineVarInstruction"
                        Dim index As Byte = DirectCast(Instruction, ShortInlineVarInstruction).Value
                        Dim VarName As String
                        If Lreader.Locals.Count = 0 Then
                            VarName = _methodbase.GetParameters(index - 1).Name
                        Else
                            VarName = Lreader.Locals(index).Name
                        End If
                        rtb.AppendText(VarName)
                        Me.SelectionDelagate.Invoke(VarName.Length, VarName.Length, Color.Black, New Font(rtb.Font, FontStyle.Bold))

                End Select

                rtb.AppendText(vbNewLine)

            End While

        End Using

        rtb.AppendText("}")

    End Sub
    Private Sub SelectText(ByVal s_start As Integer, ByVal s_length As Integer, ByVal s_color As Color, ByVal s_font As Font)

        rtb.SelectionStart = rtb.TextLength - s_start
        rtb.SelectionLength = s_length
        rtb.SelectionColor = s_color
        rtb.SelectionFont = s_font

    End Sub
End Class
