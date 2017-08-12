
Imports System.Reflection
Imports System.Reflection.Emit
Public Class ILReader

    Implements IEnumerable(Of ILInstruction)

    Private IlInstructionList As New List(Of ILInstruction)

    Private m_byteArray() As [Byte]

    Private m_position As Int32

    Private m_enclosingMethod As MethodBase

    Private methodBody As MethodBody

    Public Shared s_OneByteOpCodes(255) As OpCode

    Public Shared s_TwoByteOpCodes(255) As OpCode

    Shared Sub New()

        For Each fi As FieldInfo In GetType(OpCodes).GetFields(BindingFlags.Public Or BindingFlags.[Static])

            Dim opCode As OpCode = DirectCast(fi.GetValue(Nothing), OpCode)

            Dim value As Short = opCode.Value

            If value < &H100 AndAlso value > -1 Then
                s_OneByteOpCodes(value) = opCode


            ElseIf (value And &HFF00) = &HFE00 Then

                s_TwoByteOpCodes(value And &HFF) = opCode
            End If

        Next
    End Sub

    Public Sub New(ByVal enclosingMethod As MethodBase)

        Me.m_enclosingMethod = enclosingMethod
        Me.methodBody = m_enclosingMethod.GetMethodBody()
        Me.m_byteArray = If((methodBody Is Nothing), New [Byte](-1) {}, methodBody.GetILAsByteArray())
        Me.m_position = -1

    End Sub
    <Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.Deny)> _
    Private Function NextInstruction() As ILInstruction

        Dim offset As Int32 = m_position
        Dim opCode As OpCode = OpCodes.Nop
        Dim token As Int32 = 0

        ' read first 1 or 2 bytes as opCode
        Dim value As [Byte] = ReadByte()

        If value <> &HFE Then
            opCode = s_OneByteOpCodes(value)
        Else
            value = ReadByte()
            opCode = s_TwoByteOpCodes(value)
        End If

        Select Case opCode.OperandType
            Case OperandType.InlineNone
                Return New InlineNoneInstruction(m_enclosingMethod, offset, opCode)

            Case OperandType.ShortInlineBrTarget
                Dim shortDelta As [SByte] = ReadSByte()
                Return New ShortInlineBrTargetInstruction(m_enclosingMethod, offset, opCode, shortDelta)

            Case OperandType.InlineBrTarget
                Dim delta As Int32 = ReadInt32()
                Return New InlineBrTargetInstruction(m_enclosingMethod, offset, opCode, delta)
            Case OperandType.ShortInlineI
                Dim int8 As [SByte] = ReadSByte()
                Return New ShortInlinelInstruction(m_enclosingMethod, offset, opCode, int8)
            Case OperandType.InlineI
                Dim int32 As Int32 = ReadInt32()
                Return New InlinelInstruction(m_enclosingMethod, offset, opCode, int32)
            Case OperandType.InlineI8
                Dim int64 As Int64 = ReadInt64()
                Return New Inline8Instruction(m_enclosingMethod, offset, opCode, int64)
            Case OperandType.ShortInlineR
                Dim float32 As [Single] = ReadSingle()
                Return New ShortInlineRInstruction(m_enclosingMethod, offset, opCode, float32)
            Case OperandType.InlineR
                Dim float64 As [Double] = ReadDouble()
                Return New InlineRInstruction(m_enclosingMethod, offset, opCode, float64)
            Case OperandType.ShortInlineVar
                Dim index8 As [Byte] = ReadByte()
                Return New ShortInlineVarInstruction(m_enclosingMethod, offset, opCode, index8)
            Case OperandType.InlineVar
                Dim index16 As UInt16 = ReadUInt16()
                Return New InlineVarInstruction(m_enclosingMethod, offset, opCode, index16)
            Case OperandType.InlineString
                token = ReadInt32()
                Return New InlineStringInstruction(m_enclosingMethod, offset, opCode, token)
            Case OperandType.InlineSig
                token = ReadInt32()
                Return New InlineSigInstruction(m_enclosingMethod, offset, opCode, token)
            Case OperandType.InlineField
                token = ReadInt32()
                Return New InlineFieldInstruction(m_enclosingMethod, offset, opCode, token)
            Case OperandType.InlineType
                token = ReadInt32()
                Return New InlineTypeInstruction(m_enclosingMethod, offset, opCode, token)
            Case OperandType.InlineTok
                token = ReadInt32()
                Return New InlineTokInstruction(m_enclosingMethod, offset, opCode, token)
            Case OperandType.InlineMethod
                token = ReadInt32()
                Return New InlineMethodInstruction(m_enclosingMethod, offset, opCode, token)

            Case OperandType.InlineSwitch
                Dim cases As Int32 = ReadInt32()
                Dim deltas As Int32() = New Int32(cases - 1) {}
                For i As Int32 = 0 To cases - 1
                    deltas(i) = ReadInt32()
                Next
                Return New InlineSwitchInstruction(m_enclosingMethod, offset, opCode, deltas)
            Case Else
                Throw New BadImageFormatException("unexpected OperandType " + opCode.OperandType)
        End Select
    End Function
    Private Function ReadByte() As [Byte]
        Return DirectCast(m_byteArray(System.Math.Max(System.Threading.Interlocked.Increment(m_position), m_position - 1)), [Byte])
    End Function
    Private Function ReadSByte() As [SByte]
        Dim int8 As SByte
        Dim b As Byte = ReadByte()
        int8 = Convert.ToSByte(Hex(b), 16)
        Return int8
    End Function
    Private Function ReadUInt16() As UInt16
        m_position += 2
        Return BitConverter.ToUInt16(m_byteArray, m_position - 1)
    End Function
    Private Function ReadUInt32() As UInt32
        m_position += 4
        Return BitConverter.ToUInt32(m_byteArray, m_position - 3)
    End Function
    Private Function ReadUInt64() As UInt64
        m_position += 8
        Return BitConverter.ToUInt64(m_byteArray, m_position - 7)
    End Function
    Private Function ReadInt32() As Int32
        m_position += 4
        Return BitConverter.ToInt32(m_byteArray, m_position - 3)
    End Function
    Private Function ReadInt64() As Int64
        m_position += 8
        Return BitConverter.ToInt64(m_byteArray, m_position - 7)
    End Function
    Private Function ReadSingle() As [Single]
        m_position += 4
        Return BitConverter.ToSingle(m_byteArray, m_position - 3)
    End Function
    Private Function ReadDouble() As [Double]
        m_position += 8
        Return BitConverter.ToDouble(m_byteArray, m_position - 7)

    End Function

    Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of ILInstruction) Implements System.Collections.Generic.IEnumerable(Of ILInstruction).GetEnumerator
        While m_position < m_byteArray.Length - 1
            IlInstructionList.Add(NextInstruction())
        End While
        Return IlInstructionList.GetEnumerator
    End Function

    Public Function GetEnumerator1() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function
End Class
