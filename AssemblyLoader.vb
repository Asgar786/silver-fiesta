Imports System.Text
Imports System.Reflection
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Diagnostics.SymbolStore
Imports System.Collections.Generic
Public Class AssemblyLoader
    Private m_filename As String
    Private l_assembly As [Assembly]
    Private reader As SymbolReader.SymbolReader
    Private t_comparer As New Comparison(Of Type)(AddressOf TypeComparer)
    Private r_comparer As New Comparison(Of AssemblyName)(AddressOf AssemblyNameComparer)
    Public ReadOnly Property [Assembly]() As Assembly
        Get
            Return l_assembly
        End Get
    End Property
    Public ReadOnly Property SymReader() As SymbolReader.SymbolReader
        Get
            reader = New SymbolReader.SymbolReader(m_filename)
            Return reader
        End Get
    End Property
    Public Sub New(ByVal filename As String)
        m_filename = filename
    End Sub
    Public Sub LoadAssembly()
        Try
            l_assembly = Assembly.LoadFrom(m_filename)
        Catch ex As Exception

        End Try
    End Sub
    Public Function GetModulesFromAssembly() As List(Of [Module])
        Dim moduleList As New List(Of [Module])

        For Each modl As [Module] In l_assembly.GetModules()
            moduleList.Add(modl)
        Next
        moduleList.Sort()
        Return moduleList
    End Function
    Public Function GetReferencedAssembliesfromAssembly() As List(Of AssemblyName)
        Dim RefAsmNames As New List(Of AssemblyName)
        For Each name As AssemblyName In l_assembly.GetReferencedAssemblies
            RefAsmNames.Add(name)

        Next
        RefAsmNames.Sort(r_comparer)
        Return RefAsmNames
    End Function
    Public Function GetResourcesFromAssembly() As List(Of String)
        Dim ResourceList As New List(Of String)
        For Each r As String In l_assembly.GetManifestResourceNames
            ResourceList.Add(r)
        Next
        ResourceList.Sort()
        Return ResourceList
    End Function
    Public Function GetTypesFromAssembly() As List(Of Type)

        Dim TypeList As New List(Of System.Type)

        For Each t As Type In l_assembly.GetTypes()
            If Not t.IsNested Then
                TypeList.Add(t)
            End If
        Next

        TypeList.Sort(t_comparer)
        Return TypeList

    End Function
    Public Function GetNameSpaces() As List(Of String)

        Dim NamespaceList As New List(Of String)
        Dim tempList As New List(Of String)

        Using Enumerator As IEnumerator(Of Type) = GetTypesFromAssembly.GetEnumerator

            While Enumerator.MoveNext

                Dim s As String = String.Empty
                tempList = Enumerator.Current.FullName.Split(New Char() {Type.Delimiter}, StringSplitOptions.RemoveEmptyEntries).ToList
                tempList.Remove(tempList.Last)

                Using Enumerator2 As IEnumerator(Of String) = tempList.GetEnumerator
                    While Enumerator2.MoveNext
                        s &= Enumerator2.Current & Type.Delimiter
                    End While
                End Using

                If s.Contains(Type.Delimiter) Then
                    s = s.Remove(s.LastIndexOf(Type.Delimiter), 1)
                End If

                If Not NamespaceList.Contains(s) Then
                    NamespaceList.Add(s)
                End If

            End While

        End Using

        NamespaceList.Sort()
        Return NamespaceList

    End Function
    Private Shared Function TypeComparer(ByVal t1 As Type, ByVal t2 As Type)
        Return String.Compare(t1.Name, t2.Name)
    End Function
    Private Shared Function AssemblyNameComparer(ByVal a1 As AssemblyName, ByVal a2 As AssemblyName)
        Return String.Compare(a1.Name, a2.Name)
    End Function
End Class
Public Class TypeMembers
    Shared f_comparer As New Comparison(Of FieldInfo)(AddressOf FieldComparer)
    Shared e_comparer As New Comparison(Of EventInfo)(AddressOf EventComparer)
    Shared p_comparer As New Comparison(Of PropertyInfo)(AddressOf PropertyComparer)
    Shared m_comparer As New Comparison(Of MethodInfo)(AddressOf MethodComparer)
    Public Shared Function GetFields(ByVal type As Type) As List(Of FieldInfo)
        Dim FieldInfoList As List(Of FieldInfo) = type.GetFields(BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.IgnoreCase).ToList
        FieldInfoList.Sort(f_comparer)
        Return FieldInfoList
    End Function
    Public Shared Function GetEvents(ByVal type As Type) As List(Of EventInfo)
        Dim EventInfoList As List(Of EventInfo) = type.GetEvents(BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.IgnoreCase).ToList
        EventInfoList.Sort(e_comparer)
        Return EventInfoList
    End Function
    Public Shared Function GetProperties(ByVal type As Type) As List(Of PropertyInfo)
        Dim PropertyInfoList As List(Of PropertyInfo) = type.GetProperties(BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.IgnoreCase).ToList
        PropertyInfoList.Sort(p_comparer)
        Return PropertyInfoList
    End Function
    Public Shared Function GetMethods(ByVal type As Type) As List(Of MethodInfo)
        Dim MethodInfoList As List(Of MethodInfo) = type.GetMethods(BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.IgnoreCase).ToList
        MethodInfoList.Sort(m_comparer)
        Return MethodInfoList
    End Function
    Public Shared Function GetConstructors(ByVal type As Type) As List(Of ConstructorInfo)
        Dim ConstructorInfoList As List(Of ConstructorInfo) = type.GetConstructors(BindingFlags.DeclaredOnly Or BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.IgnoreCase).ToList
        Return ConstructorInfoList
    End Function
    Private Shared Function FieldComparer(ByVal f1 As FieldInfo, ByVal f2 As FieldInfo)
        Return String.Compare(f1.Name, f2.Name)
    End Function
    Private Shared Function EventComparer(ByVal e1 As EventInfo, ByVal e2 As EventInfo)
        Return String.Compare(e1.Name, e2.Name)
    End Function
    Private Shared Function PropertyComparer(ByVal p1 As PropertyInfo, ByVal p2 As PropertyInfo)
        Return String.Compare(p1.Name, p2.Name)
    End Function
    Private Shared Function MethodComparer(ByVal m1 As MethodInfo, ByVal m2 As MethodInfo)
        Return String.Compare(m1.Name, m2.Name)
    End Function
End Class
Public Class Assembly_Info
    Dim assembly As [Assembly]
    Public Sub New(ByVal assem As Assembly)
        assembly = assem
    End Sub
    Public Shadows Function ToString() As String
        Dim sb As New StringBuilder()
        sb.Append(assembly.FullName)
        sb.AppendLine()
        sb.AppendFormat(assembly.ImageRuntimeVersion)
        sb.AppendLine()
        Return sb.ToString()
    End Function
End Class
