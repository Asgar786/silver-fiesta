Imports System.Reflection
Public Class TypeDescriptor
    Public Shared Function TypeImageKey(ByVal t As Type) As String
        Dim ImageKeyString As String = String.Empty
        Select Case True
            Case t.IsClass
                Dim tp As Type = t.BaseType
                If t.BaseType Is GetType(System.MulticastDelegate) Then
                    ImageKeyString = "VSObject_Delegate.bmp"
                Else
                    If t.IsVisible Then
                        ImageKeyString = "VSObject_Class.bmp"
                    ElseIf t.IsSealed Then
                        ImageKeyString = "VSObject_Class_Sealed.bmp"

                    Else
                        ImageKeyString = "VSObject_Class.bmp"
                    End If
                End If
            Case t.IsEnum
                ImageKeyString = "VSObject_Enum.bmp"
            Case t.IsValueType
                ImageKeyString = "VSObject_Structure.bmp"
            Case t.IsInterface
                ImageKeyString = "VSObject_Interface.bmp"
        End Select
        Return ImageKeyString
    End Function
    Public Shared Function FieldImageKey(ByVal info As FieldInfo) As String
        Dim ImageKeyString As String = "VSObject_Field.bmp"
        If info.IsPrivate Then
            ImageKeyString = "VSObject_Field_Private.bmp"
        ElseIf info.IsFamily Then
            ImageKeyString = "VSObject_Field_Protected.bmp"
        ElseIf info.IsFamilyOrAssembly Then
            ImageKeyString = "VSObject_Field_Friend.bmp"
        ElseIf Not info.DeclaringType.IsEnum And info.IsLiteral Then
            ImageKeyString = "VSObject_Constant.bmp"
        ElseIf info.DeclaringType.IsEnum Then
            If Not info.Name = "value__" Then
                ImageKeyString = "VSObject_EnumItem.bmp"
            End If
            End If
            Return ImageKeyString
    End Function
    Public Shared Function MethodImageKey(ByVal info As MethodInfo) As String
        Dim imagekey As String = "VSObject_Method.bmp"
        Select Case True
            Case info.IsPrivate
                imagekey = "VSObject_Method_Private.bmp"
            Case info.IsFamily
                imagekey = "VSObject_Method_Protected.bmp"
            Case info.IsFamilyOrAssembly
                imagekey = "VSObject_Method_Friend.bmp"
        End Select
        Return imagekey
    End Function
    Public Shared Function PropertyImageKey(ByVal info As PropertyInfo) As String
        Dim imagekey As String = "VSObject_Properties.bmp"
        Return imagekey
    End Function
    Public Shared Function EventImageKey(ByVal info As EventInfo) As String
        Dim imagekey As String = "VSObject_Event.bmp"
        Return imagekey
    End Function
End Class
