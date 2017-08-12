Imports System.Diagnostics.SymbolStore
Imports System.Reflection
Imports System.Text
Public Class MethodDescriptor
    Private m_info As MethodInfo
    Public Sub New(ByVal m_methodinfo As MethodInfo)
        m_info = m_methodinfo
    End Sub
    Public ReadOnly Property ReturnType() As Type
        Get
            Return m_info.ReturnType
        End Get
    End Property
    Public ReadOnly Property Attributes() As String
        Get
            Dim str As String = String.Empty
            Dim attr As MethodAttributes = m_info.Attributes

            If attr And MethodAttributes.Static <> MethodAttributes.Static Then
                str = "instance"
            End If
            Return (attr.ToString.Replace(","c, ChrW(32)).ToLower & str)
        End Get
    End Property
    Public ReadOnly Property Name() As String
        Get
            Return m_info.Name
        End Get
    End Property
    Public ReadOnly Property Parameters() As String
        Get
            Dim sb As New StringBuilder()
            Dim params() As ParameterInfo = m_info.GetParameters()
            For Each pinfo As ParameterInfo In params

                sb.Append(pinfo.ParameterType.ToString & ChrW(32) & pinfo.Name & ",")
            Next
            If params.Count > 0 Then sb.Remove(sb.Length - 1, 1)
            Return sb.ToString
        End Get
    End Property
    Public ReadOnly Property ImplAttributes() As String
        Get
            Return m_info.GetMethodImplementationFlags.ToString.ToLower
        End Get
    End Property
End Class