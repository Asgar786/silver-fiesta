Imports Reader
Imports Reader.Formatter
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports Reader.Blocks
Imports Reader.CodeModel

Public Class TreeLoader

    Private filename As String
    Private Loader As AssemblyLoader
    Public Sub New(ByVal m_filename As String)

        filename = m_filename
        Loader = New AssemblyLoader(m_filename)

    End Sub
    Public Sub LoadTree(ByVal TreeView1 As TreeView)

        Dim AssemblyName As String = Path.GetFileNameWithoutExtension(filename)

        Loader.LoadAssembly()
        TreeView1.Nodes.Add(AssemblyName, AssemblyName, "VSObject_Assembly.bmp")
        TreeView1.Nodes(AssemblyName).SelectedImageKey = "VSObject_Assembly.bmp"

        Using [Enumerator] As IEnumerator(Of [Module]) = Loader.GetModulesFromAssembly.GetEnumerator
            While Enumerator.MoveNext
                Dim m As [Module] = Enumerator.Current
                TreeView1.Nodes(AssemblyName).Nodes.Add(m.ScopeName, m.ScopeName, "VSObject_Module.bmp")
                TreeView1.Nodes(AssemblyName).Nodes(m.ScopeName).SelectedImageKey = "VSObject_Module.bmp"
                TreeView1.Nodes(AssemblyName).Nodes(m.ScopeName).Nodes.Add("References", "References", "")
                TreeView1.Nodes(AssemblyName).Nodes(m.ScopeName).Nodes("References").SelectedImageKey = ""
                PopulateReferencesNode(TreeView1.Nodes(AssemblyName).Nodes(m.ScopeName).Nodes("References"), Loader.Assembly)
                PopulateModuleNodes(TreeView1.Nodes(AssemblyName).Nodes(m.ScopeName))
            End While
        End Using

        TreeView1.Nodes(AssemblyName).Nodes.Add("Resources", "Resources", "VSFolder_closed.bmp")
        TreeView1.Nodes(AssemblyName).Nodes("Resources").SelectedImageKey = "VSFolder_closed.bmp"

        Using [Enumerator] As IEnumerator(Of String) = Loader.GetResourcesFromAssembly.GetEnumerator
            While Enumerator.MoveNext
                TreeView1.Nodes(AssemblyName).Nodes("Resources").Nodes.Add(Enumerator.Current, Enumerator.Current, "VSProject_resourcefile.bmp")
                TreeView1.Nodes(AssemblyName).Nodes("Resources").Nodes(Enumerator.Current).SelectedImageKey = "VSProject_resourcefile.bmp"
                TreeView1.Nodes(AssemblyName).Nodes("Resources").Nodes(Enumerator.Current).Tag = Enumerator.Current
            End While
        End Using

    End Sub
    Public Sub PopulateModuleNodes(ByVal node As TreeNode)

        Using Enumerator As IEnumerator(Of String) = Loader.GetNameSpaces.GetEnumerator
            While Enumerator.MoveNext
                Dim n As String = Enumerator.Current
                If n = "" Then n = "-"
                node.Nodes.Add(n, n, "VSObject_Namespace.bmp")
                node.Nodes(n).SelectedImageKey = "VSObject_Namespace.bmp"
            End While
        End Using

        PopulateNamespaceNodes(node)

    End Sub
    Public Sub PopulateReferencesNode(ByVal node As TreeNode, ByVal _assembly As Assembly)

        Dim Alist As List(Of AssemblyName) = _assembly.GetReferencedAssemblies().ToList

        Using Enumerator As IEnumerator(Of AssemblyName) = Alist.GetEnumerator
            While Enumerator.MoveNext
                node.Nodes.Add(Enumerator.Current.Name, Enumerator.Current.Name, "VSObject_Assembly")
            End While
        End Using

    End Sub
    Public Sub PopulateNamespaceNodes(ByVal node As TreeNode)

        Using Enumerator As IEnumerator(Of Type) = Loader.GetTypesFromAssembly.GetEnumerator

            While Enumerator.MoveNext

                Dim t As Type = Enumerator.Current
                Dim str As String = t.Name
                If t.IsGenericType Then str = FormatTypeString(t)
                Dim path As List(Of String) = Enumerator.Current.FullName.Split(New Char() {Type.Delimiter}, StringSplitOptions.RemoveEmptyEntries).ToList
                path.Remove(path.Last)
                Dim nodetext As String = String.Join(Type.Delimiter, path.ToArray)

                If nodetext = "" Then nodetext = "-"
                node.Nodes(nodetext).Nodes.Add(Enumerator.Current.FullName, str, TypeDescriptor.TypeImageKey(Enumerator.Current))
                node.Nodes(nodetext).Nodes(Enumerator.Current.FullName).SelectedImageKey = TypeDescriptor.TypeImageKey(Enumerator.Current)
                node.Nodes(nodetext).Nodes(Enumerator.Current.FullName).Nodes.Add("PlaceHolder")
                node.Nodes(nodetext).Nodes(Enumerator.Current.FullName).Tag = Enumerator.Current

            End While

        End Using

    End Sub
    Public Sub PopulateTypeNode(ByVal node As TreeNode)

        Dim t As System.Type = DirectCast(node.Tag, System.Type)

        If t Is Nothing Then Return

        Using [Enumerator] As IEnumerator(Of Type) = t.GetNestedTypes(CType(62, BindingFlags)).ToList.GetEnumerator
            While Enumerator.MoveNext

                Dim type As Type = Enumerator.Current
                Dim TypeString As String = New TypeFormatter(Enumerator.Current).ToString(IFormatter.FormatOption.OptionShort)

                node.Nodes.Add(Enumerator.Current.FullName, TypeString, TypeDescriptor.TypeImageKey(Enumerator.Current))
                node.Nodes(Enumerator.Current.FullName).SelectedImageKey = TypeDescriptor.TypeImageKey(Enumerator.Current)
                node.Nodes(Enumerator.Current.FullName).Nodes.Add("PlaceHolder")
                node.Nodes(Enumerator.Current.FullName).Tag = Enumerator.Current
                PopulateTypeNode(node.Nodes(Enumerator.Current.FullName))
            End While
        End Using

        Using [Enumerator] As IEnumerator(Of ConstructorInfo) = TypeMembers.GetConstructors(t).GetEnumerator
            While Enumerator.MoveNext
                Dim ConstructorString As String = New ConstructorFormatter(Enumerator.Current).ToString(IFormatter.FormatOption.OptionShort)
                node.Nodes.Add(ConstructorString, ConstructorString, "VSObject_Method.bmp")
                node.Nodes(ConstructorString).SelectedImageKey = "VSObject_Method.bmp"
                node.Nodes(ConstructorString).Tag = Enumerator.Current
            End While
        End Using

        Using [Enumerator] As IEnumerator(Of MethodInfo) = TypeMembers.GetMethods(t).GetEnumerator

            While Enumerator.MoveNext

                Dim m As MethodInfo = Enumerator.Current
                Dim str As String = New MethodFormatter(Enumerator.Current).ToString(IFormatter.FormatOption.OptionShort)

                If Not (m.IsSpecialName AndAlso (m.Name.StartsWith("get_") OrElse m.Name.StartsWith("set_"))) Then
                    node.Nodes.Add(str, str, TypeDescriptor.MethodImageKey(Enumerator.Current))
                    node.Nodes(str).SelectedImageKey = TypeDescriptor.MethodImageKey(Enumerator.Current)
                    node.Nodes(str).Tag = Enumerator.Current
                End If

            End While

        End Using

        Using [Enumerator] As IEnumerator(Of PropertyInfo) = TypeMembers.GetProperties(t).GetEnumerator

            While Enumerator.MoveNext
                Dim pinfo As PropertyInfo = Enumerator.Current
                Dim str As String = New PropertyFormatter(pinfo).ToString(IFormatter.FormatOption.OptionShort)
                Dim Key As String = "P" & Enumerator.Current.Name
                node.Nodes.Add(Key, str, TypeDescriptor.PropertyImageKey(Enumerator.Current))
                node.Nodes(Key).SelectedImageKey = TypeDescriptor.PropertyImageKey(Enumerator.Current)

                Dim Getinfo As MethodInfo = Enumerator.Current.GetGetMethod(True)
                Dim SetInfo As MethodInfo = Enumerator.Current.GetSetMethod(True)

                If Getinfo IsNot Nothing Then
                    Dim GetInfoString As String = New MethodFormatter(Getinfo).ToString(IFormatter.FormatOption.OptionShort)
                    node.Nodes(Key).Nodes.Add(Getinfo.Name, GetInfoString, TypeDescriptor.MethodImageKey(Getinfo))
                    node.Nodes(Key).Nodes(Getinfo.Name).SelectedImageKey = TypeDescriptor.MethodImageKey(Getinfo)
                    node.Nodes(Key).Nodes(Getinfo.Name).Tag = Getinfo
                End If

                If SetInfo IsNot Nothing Then
                    Dim SetinfoString As String = New MethodFormatter(SetInfo).ToString(IFormatter.FormatOption.OptionShort)
                    node.Nodes(Key).Nodes.Add(SetInfo.Name, SetinfoString, TypeDescriptor.MethodImageKey(SetInfo))
                    node.Nodes(Key).Nodes(SetInfo.Name).SelectedImageKey = TypeDescriptor.MethodImageKey(SetInfo)
                    node.Nodes(Key).Nodes(SetInfo.Name).Tag = SetInfo
                End If

            End While

        End Using

        Using [Enumerator] As IEnumerator(Of EventInfo) = TypeMembers.GetEvents(t).GetEnumerator
            While Enumerator.MoveNext
                node.Nodes.Add(Enumerator.Current.Name, Enumerator.Current.Name, TypeDescriptor.EventImageKey(Enumerator.Current))
                node.Nodes(Enumerator.Current.Name).SelectedImageKey = TypeDescriptor.EventImageKey(Enumerator.Current)
            End While
        End Using

        Using [Enumerator] As IEnumerator(Of FieldInfo) = TypeMembers.GetFields(t).GetEnumerator
            While Enumerator.MoveNext
                Dim FieldString As String = Enumerator.Current.Name
                If Enumerator.Current.FieldType.Name <> Enumerator.Current.DeclaringType.Name Then
                    FieldString = String.Format("{0}{1}:{2}{3}", Enumerator.Current.Name, ChrW(32), ChrW(32), FormatTypeString(Enumerator.Current.FieldType))
                End If
                node.Nodes.Add(FieldString, FieldString, TypeDescriptor.FieldImageKey(Enumerator.Current))
                node.Nodes(FieldString).SelectedImageKey = TypeDescriptor.FieldImageKey(Enumerator.Current)
            End While
        End Using

    End Sub
    Public Function FormatTypeString(ByVal type As Type) As String

        Dim sb As New StringBuilder(type.Name)

        If type.IsGenericType Then

            sb.Remove(sb.Length - 2, 2)
            sb.Append(ChrW(60))

            Using [Enumerator] As IEnumerator(Of Type) = type.GetGenericArguments.ToList.GetEnumerator

                While Enumerator.MoveNext

                    Dim t As Type = Enumerator.Current
                    sb.Append(t.Name)

                    If t.IsGenericType Then
                        sb.Remove(sb.Length - t.Name.Length, t.Name.Length)
                        sb.Append(FormatTypeString(t))
                    End If

                    sb.Append(ChrW(44))

                End While

            End Using

            sb.Remove(sb.Length - 1, 1)
            sb.Append(ChrW(62))

        End If

        Return sb.ToString

    End Function
    Public Function FormatMethodString(ByVal minfo As MethodInfo) As String

        Dim sb As New StringBuilder()
        sb.Append(minfo.Name)
        sb.Append(ChrW(40))

        Using [Enumerator] As IEnumerator(Of ParameterInfo) = minfo.GetParameters.ToList.GetEnumerator
            While Enumerator.MoveNext
                Dim paramType As String = FormatTypeString(Enumerator.Current.ParameterType)
                sb.AppendFormat("{0}{1}{2}", paramType, ChrW(44), ChrW(32))
            End While
        End Using

        If minfo.GetParameters.Count > 0 Then sb.Remove(sb.Length - 2, 2)
        sb.AppendFormat("{0}{1}{2}{3}", ChrW(41), ChrW(32), ChrW(58), ChrW(32))

        Dim rtype As String = FormatTypeString(minfo.ReturnType)
        sb.Append(rtype)

        Return sb.ToString

    End Function
    Public Function FormatMethodString(ByVal cinfo As ConstructorInfo) As String

        Dim sb As New StringBuilder()
        sb.Append(cinfo.Name)
        sb.Append(ChrW(40))

        Using [Enumerator] As IEnumerator(Of ParameterInfo) = cinfo.GetParameters.ToList.GetEnumerator
            While Enumerator.MoveNext
                Dim paramType As String = FormatTypeString(Enumerator.Current.ParameterType)
                sb.AppendFormat("{0}{1}{2}", paramType, ChrW(44), ChrW(32))
            End While
        End Using

        If cinfo.GetParameters.Count > 0 Then sb.Remove(sb.Length - 2, 2)
        sb.Append(ChrW(41))

        Return sb.ToString

    End Function
    Public Function CurrentToRootNode(ByVal node As TreeNode)
        Static cTorList As New List(Of TreeNode)
        cTorList.Add(node)
        If node.Parent IsNot Nothing Then
            CurrentToRootNode(node.Parent)
        End If
        Return cTorList
    End Function
    ''Public Sub CollapseNodes(ByVal node As TreeNode, ByVal treeView1 As TreeView)

    ''    Dim nodeList As New List(Of String)
    ''    CurrentToRootNode(node, nodeList)
    ''    Dim items = (From n As TreeNode In treeView1.Nodes(0).Nodes Where Not nodeList.Contains(n) Select n).ToList

    ''    If Not node.Parent Is Nothing Then
    ''        Dim nodes() = (From n As TreeNode In node.Parent.Nodes Where Not nodeList.Contains(n) Select n).ToArray
    ''        items.AddRange(nodes)
    ''    End If

    ''    For Each n As TreeNode In items
    ''        n.Collapse(False)
    ''    Next

    ''End Sub

End Class