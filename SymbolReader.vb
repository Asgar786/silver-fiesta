Imports System.Diagnostics
Imports System.Diagnostics.SymbolStore
Imports Microsoft.Samples.Debugging.CorSymbolStore
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Security.Permissions
Namespace SymbolReader
    <ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), CLSCompliant(True), Guid("809c652e-7396-11d2-9771-00a0c9b4d50c")> _
     Public Interface IMetaDataDispenser
        Sub DefineScope_Placeholder()
        Sub OpenScope(<[In](), MarshalAs(UnmanagedType.LPWStr)> ByVal szScope As String, <[In]()> ByVal dwOpenFlags As Integer, <[In]()> ByRef riid As Guid, <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef punk As Object)
    End Interface

    <ComVisible(True), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), CLSCompliant(True), Guid("7DAC8207-D3AE-4c75-9B67-92801A497D44")> _
    Public Interface IMetadataImport
        Sub Placeholder()
    End Interface
    Class SymUtil
        ' Methods
        Public Shared Function GetSymbolReaderForFile(ByVal pathModule As String, ByVal searchPath As String) As ISymbolReader
            Return SymUtil.GetSymbolReaderForFile(New SymbolBinder, pathModule, searchPath)
        End Function
        <SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.UnmanagedCode), Security.Permissions.ReflectionPermission(SecurityAction.Demand)> _
        Public Shared Function GetSymbolReaderForFile(ByVal binder As SymbolBinder, ByVal pathModule As String, ByVal searchPath As String) As ISymbolReader
            Dim objDispenser As New Object
            Dim objImporter As New Object
            Dim reader As ISymbolReader = Nothing
            Dim dispenserClassID As New Guid(&HE5CB7A31, &H7512, &H11D2, &H89, &HCE, 0, &H80, &HC7, &H92, &HE5, &HD8)
            Dim dispenserIID As New Guid(&H809C652E, &H7396, &H11D2, &H97, &H71, 0, 160, &HC9, 180, &HD5, 12)
            Dim importerIID As New Guid(&H7DAC8207, &HD3AE, &H4C75, &H9B, &H67, &H92, &H80, &H1A, &H49, &H7D, &H44)
            NativeMethods.CoCreateInstance((dispenserClassID), Nothing, 1, (dispenserIID), objDispenser)
            DirectCast(objDispenser, IMetaDataDispenser).OpenScope(pathModule, 0, (importerIID), objImporter)
            Dim importerPtr As IntPtr = IntPtr.Zero
            Try
                importerPtr = Marshal.GetComInterfaceForObject(objImporter, GetType(IMetadataImport))
                reader = binder.GetReader(importerPtr, pathModule, searchPath)
            Catch ex As COMException

            Finally
                If (importerPtr <> IntPtr.Zero) Then
                    Marshal.Release(importerPtr)
                End If
            End Try
            Return reader
        End Function
        ' Nested Types
        Private Class NativeMethods
            ' Methods
            <DllImport("ole32.dll")> _
            Public Shared Function CoCreateInstance(<[In]()> ByRef rclsid As Guid, <[In](), MarshalAs(UnmanagedType.IUnknown)> ByVal pUnkOuter As Object, <[In]()> ByVal dwClsContext As UInt32, <[In]()> ByRef riid As Guid, <Out(), MarshalAs(UnmanagedType.Interface)> ByRef ppv As Object) As Integer
            End Function
        End Class
    End Class
    Public Class SymbolReader
        Private Reader As ISymbolReader

        Public Sub New(ByVal m_filename As String)

            Reader = SymUtil.GetSymbolReaderForFile(m_filename, m_filename)
        End Sub
        Public Function GetMethod(ByVal methodToken As Int32) As ISymbolMethod
            Dim method As ISymbolMethod = Nothing
            Try
                method = Reader.GetMethod(New SymbolToken(methodToken))
            Catch
            End Try
            Return method
        End Function
        Public Function GetLocalVariables(ByVal methodtoken As Int32) As ISymbolVariable()

            Dim LocalVariableList As New List(Of ISymbolVariable)
            Dim Method As ISymbolMethod = GetMethod(methodtoken)
            If Method IsNot Nothing Then
                LocalHelper(Method.RootScope, LocalVariableList)
            End If
            Return LocalVariableList.ToArray

        End Function
        Private Sub LocalHelper(ByVal scope As ISymbolScope, ByVal VList As List(Of ISymbolVariable))

            VList.AddRange(scope.GetLocals)

            For Each childscope As ISymbolScope In scope.GetChildren
                LocalHelper(childscope, VList)
            Next

        End Sub
        Public Function GetNamespaces() As ISymbolNamespace()
            Dim n As New List(Of ISymbolNamespace)
            Try
                n = Reader.GetNamespaces().ToList
            Catch ex As Exception

            End Try
            Return n.ToArray
        End Function
    End Class
End Namespace

