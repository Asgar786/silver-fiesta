Imports System.Diagnostics.SymbolStore
Imports System.Reflection
Imports stdole
Public Class LocalReader
    Inherits AssemblyLoader
    Private localDictionary As New Dictionary(Of Int32, String)
    Private localList As New List(Of LocalVarInfo)
    Public Sub New(ByVal m_filename As String)
        MyBase.New(m_filename)

    End Sub
    Public Sub GetLocals(ByVal m_base As MethodBase)
        If symReader Is Nothing Then
            Return
        End If
        For Each var As ISymbolVariable In symReader.GetLocalVariables(m_base.MetadataToken)
            If Not localDictionary.ContainsKey(var.AddressField1) Then
                localDictionary.Add(var.AddressField1, var.Name)
            End If
        Next
        AddVariableName(m_base)
    End Sub
    Public ReadOnly Property Locals() As List(Of LocalVarInfo)
        Get
            Return localList
        End Get
    End Property
    Public Sub AddVariableName(ByVal m_base As MethodBase)
        Dim varList As New List(Of LocalVariableInfo)
        Dim index As Integer = 1
        Try
            varList = m_base.GetMethodBody.LocalVariables.ToList
        Catch
        End Try
        For Each lvi As LocalVariableInfo In varList
            Dim LocalName As String = String.Empty
            If localDictionary.ContainsKey(lvi.LocalIndex) Then
                LocalName = localDictionary.Item(lvi.LocalIndex)
            Else
                LocalName = "var" & index.ToString
                index += 1
            End If
            localList.Add(New LocalVarInfo(lvi.IsPinned, lvi.LocalIndex, lvi.LocalType, LocalName))
        Next
    End Sub
End Class
<System.Runtime.InteropServices.ComVisibleAttribute(True)> Public NotInheritable Class LocalVarInfo
    Private _IsPinned As Boolean
    Private _LocalIndex As Integer
    Private _LocalType As System.Type
    Private _Name As String
    Public Sub New(ByVal Is_Pinned As Boolean, ByVal Local_Index As Integer, ByVal Local_Type As Type, ByVal LocalName As String)
        _IsPinned = Is_Pinned
        _LocalIndex = Local_Index
        _LocalType = Local_Type
        _Name = LocalName
    End Sub

    Public ReadOnly Property IsPinned() As Boolean
        Get
            Return _IsPinned
        End Get
    End Property
    Public ReadOnly Property LocalIndex() As Integer
        Get
            Return _LocalIndex
        End Get
    End Property
    Public ReadOnly Property LocalType() As System.Type
        Get
            Return _LocalType
        End Get
    End Property
    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim formatter() As String = New String() {"[", LocalIndex, "]", ReflectedType.RType(LocalType), Name}
        Return String.Format("{0}{1}{2} {3} {4}", formatter)
    End Function
End Class
Public Class ReflectedType
    Shared TypeDictionary As New Dictionary(Of Type, String)
    Shared Sub New()
        TypeDictionary.Add(GetType(Byte), "uint8")
        TypeDictionary.Add(GetType(SByte), "int8")
        TypeDictionary.Add(GetType(UShort), "uint16")
        TypeDictionary.Add(GetType(Short), "int16")
        TypeDictionary.Add(GetType(UInt32), "uint32")
        TypeDictionary.Add(GetType(Int32), "int32")
        TypeDictionary.Add(GetType(UInt64), "uint64")
        TypeDictionary.Add(GetType(Int64), "int64")
        TypeDictionary.Add(GetType(Single), "float32")
        TypeDictionary.Add(GetType(Double), "float64")
        TypeDictionary.Add(GetType(Boolean), "bool")
        TypeDictionary.Add(GetType(String), "string")
        TypeDictionary.Add(GetType(Void), "void")
        TypeDictionary.Add(GetType(Byte()), "uint8[]")
        TypeDictionary.Add(GetType(SByte()), "int8[]")
        TypeDictionary.Add(GetType(UShort()), "uint16[]")
        TypeDictionary.Add(GetType(Short()), "int16[]")
        TypeDictionary.Add(GetType(UInt32()), "uint32[]")
        TypeDictionary.Add(GetType(Int32()), "int32[]")
        TypeDictionary.Add(GetType(UInt64()), "uint64[]")
        TypeDictionary.Add(GetType(Int64()), "int64[]")
        TypeDictionary.Add(GetType(Single()), "float32[]")
        TypeDictionary.Add(GetType(Double()), "float64[]")
    End Sub
    Public Shared ReadOnly Property RType(ByVal type As System.Type) As String
        Get
            Dim typestring As String = String.Empty
            If TypeDictionary.ContainsKey(type) Then
                typestring = TypeDictionary.Item(type)
            ElseIf type.IsGenericType Then
                typestring = "T"
            Else
                typestring = type.FullName
            End If
            Return typestring
        End Get
    End Property
End Class