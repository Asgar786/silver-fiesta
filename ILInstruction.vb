Imports System.Reflection
Imports System.Reflection.Emit
Public MustInherit Class ILInstruction
    Protected op_code As OpCode
    Protected _offset As Int32
    Protected _enclosingMethod As MethodBase
    Protected m As [Module]
    Protected GenericTypeArgs() As System.Type
    Protected GenericTypes() As System.Type
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode)
        _enclosingMethod = m_enclosingMethod
        _offset = offset + 1
        op_code = opCode
        m = [Assembly].GetAssembly(_enclosingMethod.DeclaringType).ManifestModule
        If _enclosingMethod.ContainsGenericParameters Then
            GenericTypeArgs = _enclosingMethod.GetGenericArguments
        End If
        If _enclosingMethod.DeclaringType.IsGenericType Then
            GenericTypes = _enclosingMethod.DeclaringType.GetGenericArguments
        End If

    End Sub
    Public ReadOnly Property IlOffset() As Integer
        Get
            Return _offset
        End Get
    End Property
    Public ReadOnly Property Code() As OpCode
        Get
            Return op_code
        End Get
    End Property
    Public MustOverride ReadOnly Property Value() As Object
    Public Overridable Shadows Function ToString() As String
        Dim xString As String = Hex(_offset).ToLower.PadLeft(4, "0"c)
        Dim InstructionName As String = Me.op_code.Name
        Return String.Format("IL_{0}:{1}{2}{3}{4}", xString, vbTab, InstructionName, vbTab, vbTab)
    End Function
End Class
Public Class InlineNoneInstruction
    Inherits ILInstruction
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode)
        MyBase.New(m_enclosingMethod, offset, opCode)
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return Nothing
        End Get
    End Property
    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function
End Class
Public Class ShortInlineBrTargetInstruction
    Inherits ILInstruction
    Private short_delta As [SByte]
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal shortdelta As SByte)
        MyBase.New(m_enclosingMethod, offset, opCode)
        short_delta = shortdelta
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _offset + short_delta + 2
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString

        Dim Label As String = "IL_" & Hex(_offset + short_delta + 2).ToLower.PadLeft(4, "0"c)
        Return String.Format("{0}{1}", baseString, Label)
    End Function
End Class
Public Class InlineBrTargetInstruction
    Inherits ILInstruction
    Private _delta As Int32
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal delta As [Int32])
        MyBase.New(m_enclosingMethod, offset, opCode)
        _delta = delta
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _offset + _delta + 5
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString

        Return String.Format("{0}{1}", baseString, "")
    End Function
End Class
Public Class ShortInlinelInstruction
    Inherits ILInstruction
    Private _int8 As [SByte]
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal int8 As [SByte])
        MyBase.New(m_enclosingMethod, offset, opCode)
        _int8 = int8
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _int8
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString()
        Return String.Format("{0}{1}", baseString, Value)
    End Function
End Class
Public Class InlinelInstruction
    Inherits ILInstruction
    Private _int32 As Int32
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal int32 As Int32)
        MyBase.New(m_enclosingMethod, offset, opCode)
        _int32 = int32
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _int32
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString()
        Return String.Format("{0}{1}", baseString, Value)
    End Function
End Class
Public Class Inline8Instruction
    Inherits ILInstruction
    Private _int64 As Int64
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal int64 As Int64)
        MyBase.New(m_enclosingMethod, offset, opCode)
        _int64 = int64
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _int64
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString()
        Return String.Format("{0}{1}", baseString, Value)
    End Function
End Class
Public Class ShortInlineRInstruction
    Inherits ILInstruction
    Private _float32 As [Single]
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal float32 As [Single])
        MyBase.New(m_enclosingMethod, offset, opCode)
        _float32 = float32
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _float32
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString()
        Return String.Format("{0}{1}", baseString, Value)
    End Function
End Class
Public Class InlineRInstruction
    Inherits ILInstruction
    Private _float64 As [Double]
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal float64 As [Double])
        MyBase.New(m_enclosingMethod, offset, opCode)
        _float64 = float64
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _float64
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString()
        Return String.Format("{0}{1}", baseString, Value)
    End Function
End Class
Public Class ShortInlineVarInstruction
    Inherits ILInstruction
    Private _index8 As [Byte]
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal index8 As [Byte])
        MyBase.New(m_enclosingMethod, offset, opCode)
        _index8 = index8
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _index8
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim t As Type = MyBase._enclosingMethod.DeclaringType
        Dim baseString As String = MyBase.ToString

        Return String.Format("{0}{1}", baseString, "")

    End Function
End Class
Public Class InlineVarInstruction
    Inherits ILInstruction
    Private _index16 As UInt16
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal index16 As UInt16)
        MyBase.New(m_enclosingMethod, offset, opCode)
        _index16 = index16
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return _index16
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString()
        Return String.Format("{0}{1}", baseString, Value)
    End Function
End Class
Public Class InlineStringInstruction
    Inherits ILInstruction
    Private _token As Int32
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal token As Int32)
        MyBase.New(m_enclosingMethod, offset, opCode)
        _token = token
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return m.ResolveString(_token)
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim t As Type = MyBase._enclosingMethod.DeclaringType
        Dim baseString As String = MyBase.ToString
        Return String.Format("{0}""{1}""", baseString, Me.Value)

    End Function
End Class
Public Class InlineSigInstruction
    Inherits ILInstruction
    Private _token As Int32
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal token As Int32)
        MyBase.New(m_enclosingMethod, offset, opCode)
        _token = token
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return m.ResolveSignature(_token)
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString()
        Return String.Format("{0}{1}", baseString, Value)
    End Function
End Class
Public Class InlineFieldInstruction
    Inherits ILInstruction
    Private _token As Int32
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal token As Int32)
        MyBase.New(m_enclosingMethod, offset, opCode)
        _token = token
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return m.ResolveField(_token, GenericTypes, GenericTypeArgs)
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString()
        Dim type As System.Type = Me.Value.DeclaringType
        Dim assem As Assembly = Assembly.GetAssembly(type)
        Return String.Format("{0}[{1}]{2}::{3}", baseString, assem.FullName.Split(",")(0), type, Me.Value.Name)
    End Function
End Class
Public Class InlineTypeInstruction
    Inherits ILInstruction
    Private _token As Int32
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal token As Int32)
        MyBase.New(m_enclosingMethod, offset, opCode)
        _token = token
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return m.ResolveType(_token, GenericTypes, GenericTypeArgs)
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString
        Return String.Format("{0}{1}", baseString, Value.ToString)
    End Function
End Class
Public Class InlineTokInstruction
    Inherits ILInstruction
    Private _token As Int32
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal token As Int32)
        MyBase.New(m_enclosingMethod, offset, opCode)
        _token = token
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get

            Return m.ResolveMember(_token, GenericTypes, GenericTypeArgs)
        End Get
    End Property

End Class
Public Class InlineMethodInstruction
    Inherits ILInstruction
    Private _token As Int32

    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal token As Int32)
        MyBase.New(m_enclosingMethod, offset, opCode)
        _token = token
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Return m.ResolveMethod(_token, GenericTypes, GenericTypeArgs)
        End Get
    End Property
    Public Overrides Function ToString() As String

        Dim t As Type = Value.DeclaringType
        Dim Assem As Assembly = Assembly.GetAssembly(t)
        Dim baseString As String = MyBase.ToString
        Dim AssemblyName As String = Assem.FullName.Split(",")(0)
        Dim paramString As String = String.Empty

        For Each p As ParameterInfo In Value.GetParameters
            paramString &= p.ParameterType.ToString & ","
        Next
        Return String.Format("{0}[{1}]{2}::{3}({4})", baseString, AssemblyName, t.FullName, Value.Name, paramString.TrimEnd(New Char() {Chr(32), ","c}))
    End Function
End Class
Public Class InlineSwitchInstruction
    Inherits ILInstruction
    Private _deltas As Int32()
    Public Sub New(ByVal m_enclosingMethod As MethodBase, ByVal offset As Int32, ByVal opCode As OpCode, ByVal deltas As Int32())
        MyBase.New(m_enclosingMethod, offset, opCode)
        _deltas = deltas
    End Sub
    Public Overrides ReadOnly Property Value() As Object
        Get
            Dim _offsets As New List(Of Int32)
            For Each deltas As Int32 In _deltas
                _offsets.Add(deltas + 82)
            Next
            Return _offsets
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim baseString As String = MyBase.ToString
        Dim Label As String = "( "
        For i As Integer = 0 To _deltas.Count - 1

            Label &= "IL_" & Hex(_deltas(i) + 82).ToLower.PadLeft(4, "0"c) & ","
        Next
        Label &= " )"
        Return String.Format("{0}{1}", baseString, Label)
    End Function
End Class