Imports System.ComponentModel
''' <summary>
''' 枚举辅助类
''' </summary>
Public NotInheritable Class EnumHelper

    ''' <summary>
    ''' 获取枚举注释
    ''' </summary>
    Public Shared Function GetEnumDescription(Of T)() As List(Of String)

        If Not GetType(T).IsEnum Then
            Throw New Exception("0x0006: 非枚举类型")
        End If

        Dim tmpList As New List(Of String)

        Dim fieldItems = GetType(T).GetFields
        For Each item In fieldItems
            If Not item.FieldType.IsEnum Then
                Continue For
            End If

            Dim tmpDescriptionItems() As DescriptionAttribute = item.GetCustomAttributes(GetType(DescriptionAttribute), False)
            If tmpDescriptionItems.Count = 0 Then
                Throw New Exception("0x0007: 未找到 DescriptionAttribute")
            End If

            tmpList.Add(tmpDescriptionItems(0).Description)

        Next

        Return tmpList

    End Function

End Class
