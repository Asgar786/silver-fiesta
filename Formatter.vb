Imports System.Text
Imports System.Reflection
Imports Reader.CodeModel
Namespace Formatter
#Region "Formatters"
    Public Class TypeFormatter
        Implements IFormatter
        Private t As Type
        Public Sub New(ByVal type As Type)
            t = type
        End Sub
        Public Shadows Function ToString(ByVal FormatOption As IFormatter.FormatOption) As String Implements IFormatter.ToString

            Dim sb As New StringBuilder

            Dim typename As String = t.Name

            If FormatOption = IFormatter.FormatOption.OptionLong Then

                typename = t.ToString
            End If

            sb.Append(typename.Split("`".ToCharArray, StringSplitOptions.RemoveEmptyEntries)(0))

            If t.IsGenericType Then

                sb.Append(ChrW(60))

                Using [Enumerator] As IEnumerator(Of Type) = t.GetGenericArguments.ToList.GetEnumerator

                    While Enumerator.MoveNext

                        Dim type As Type = Enumerator.Current
                        typename = type.Name

                        If FormatOption = IFormatter.FormatOption.OptionLong Then
                            If type.FullName IsNot Nothing Then
                                typename = type.FullName.Split("`".ToCharArray, StringSplitOptions.RemoveEmptyEntries)(0)
                            End If
                        End If

                        sb.Append(typename)
                        If type.IsGenericType Then
                            sb.Remove(sb.Length - typename.Length, typename.Length)
                            sb.Append(New TypeFormatter(type).ToString(FormatOption))
                        End If

                        sb.Append(ChrW(44))

                    End While

                End Using

                sb.Remove(sb.Length - 1, 1)
                sb.Append(ChrW(62))

            End If

            Return sb.ToString

        End Function
    End Class
    Public Class ConstructorFormatter
        Implements IFormatter
        Private c_info As ConstructorInfo
        Public Sub New(ByVal info As ConstructorInfo)
            c_info = info
        End Sub
        Public Shadows Function ToString(ByVal FormatOption As IFormatter.FormatOption) As String Implements IFormatter.ToString

            Dim sb As New StringBuilder()

            If FormatOption = IFormatter.FormatOption.OptionLong Then
                sb.Append("void" & ChrW(32))
            End If

            sb.Append(c_info.Name)
            sb.Append(ChrW(40))

            Using [Enumerator] As IEnumerator(Of ParameterInfo) = c_info.GetParameters.ToList.GetEnumerator

                While Enumerator.MoveNext

                    Dim ParamInfo As ParameterInfo = Enumerator.Current
                    Dim typeformatter As New TypeFormatter(ParamInfo.ParameterType)

                    If FormatOption = IFormatter.FormatOption.OptionShort Then
                        Dim t As Type = ParamInfo.ParameterType
                        sb.AppendFormat("{0}{1}{2}", typeformatter.ToString(IFormatter.FormatOption.OptionShort), ChrW(44), ChrW(32))
                    ElseIf FormatOption = IFormatter.FormatOption.OptionLong Then
                        sb.AppendFormat("{0}{1}{2}{3}{4}", typeformatter.ToString(IFormatter.FormatOption.OptionLong), ChrW(32), ParamInfo.Name, ChrW(44), ChrW(32))
                    End If

                End While

            End Using

            If c_info.GetParameters.Count > 0 Then sb.Remove(sb.Length - 2, 2)
            sb.AppendFormat("{0}{1}", ChrW(41), ChrW(32))

            Return sb.ToString

        End Function
    End Class
    Public Class MethodFormatter
        Implements IFormatter

        Private m_info As MethodInfo
        Public Sub New(ByVal info As MethodInfo)
            m_info = info
        End Sub
        Public Shadows Function ToString(ByVal FormatOption As IFormatter.FormatOption) As String Implements IFormatter.ToString

            Dim sb As New StringBuilder()
            Dim rtype As String = New TypeFormatter(m_info.ReturnType).ToString(FormatOption)

            If FormatOption = IFormatter.FormatOption.OptionLong Then
                sb.Append(rtype & ChrW(32))
            End If

            sb.Append(m_info.Name)
            sb.Append(ChrW(40))

            Using [Enumerator] As IEnumerator(Of ParameterInfo) = m_info.GetParameters.ToList.GetEnumerator

                While Enumerator.MoveNext

                    Dim ParamInfo As ParameterInfo = Enumerator.Current
                    Dim typeformatter As New TypeFormatter(ParamInfo.ParameterType)

                    If FormatOption = IFormatter.FormatOption.OptionShort Then
                        sb.AppendFormat("{0}{1}{2}", typeformatter.ToString(IFormatter.FormatOption.OptionShort), ChrW(44), ChrW(32))
                    ElseIf FormatOption = IFormatter.FormatOption.OptionLong Then
                        sb.AppendFormat("{0}{1}{2}{3}{4}", typeformatter.ToString(IFormatter.FormatOption.OptionLong), ChrW(32), ParamInfo.Name, ChrW(44), ChrW(32))
                    End If

                End While

            End Using

            If m_info.GetParameters.Count > 0 Then sb.Remove(sb.Length - 2, 2)
            sb.AppendFormat("{0}{1}", ChrW(41), ChrW(32))

            If FormatOption = IFormatter.FormatOption.OptionShort Then
                sb.AppendFormat("{0}{1}", ChrW(58), ChrW(32))
                sb.Append(rtype)
            End If

            Return sb.ToString

        End Function
    End Class
    Public Class PropertyFormatter
        Implements IFormatter
        Private p_info As PropertyInfo
        Public Sub New(ByVal info As PropertyInfo)
            p_info = info
        End Sub

        Public Shadows Function ToString(ByVal FormatOption As IFormatter.FormatOption) As String Implements IFormatter.ToString

            Dim sb As New StringBuilder()
            Dim rtype As String = New TypeFormatter(p_info.PropertyType).ToString(FormatOption)
            Dim ParamInfoList As List(Of ParameterInfo) = p_info.GetIndexParameters.ToList

            If FormatOption = IFormatter.FormatOption.OptionLong Then
                sb.Append(rtype & ChrW(32))
            End If

            sb.Append(p_info.Name)


            If ParamInfoList.Count > 0 Then
                sb.Append(ChrW(40))
                Using Enumerator As IEnumerator(Of ParameterInfo) = ParamInfoList.GetEnumerator

                    While Enumerator.MoveNext

                        Dim ParamInfo As ParameterInfo = Enumerator.Current
                        Dim typeformatter As New TypeFormatter(ParamInfo.ParameterType)

                        If FormatOption = IFormatter.FormatOption.OptionShort Then
                            sb.AppendFormat("{0}{1}{2}", typeformatter.ToString(IFormatter.FormatOption.OptionShort), ChrW(44), ChrW(32))
                        ElseIf FormatOption = IFormatter.FormatOption.OptionLong Then
                            sb.AppendFormat("{0}{1}{2}{3}{4}", typeformatter.ToString(IFormatter.FormatOption.OptionLong), ChrW(32), ParamInfo.Name, ChrW(44), ChrW(32))
                        End If

                    End While

                End Using

                sb.Remove(sb.Length - 2, 2)
                sb.AppendFormat("{0}{1}", ChrW(41), ChrW(32))

            End If

            If FormatOption = IFormatter.FormatOption.OptionShort Then
                sb.AppendFormat("{0}{1}", ChrW(58), ChrW(32))
                sb.Append(rtype)
            End If

            Return sb.ToString

        End Function
    End Class
#End Region
End Namespace