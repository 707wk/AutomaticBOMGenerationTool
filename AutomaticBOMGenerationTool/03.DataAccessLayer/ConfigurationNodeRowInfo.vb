﻿''' <summary>
''' 配置项选中项的位置信息
''' </summary>
Public Class ConfigurationNodeRowInfo
    Public ID As String

    ''' <summary>
    ''' 所属配置项ID
    ''' </summary>
    Public ConfigurationNodeID As String
    ''' <summary>
    ''' 所属配置项名称
    ''' </summary>
    Public ConfigurationNodeName As String

    ''' <summary>
    ''' 物料行号
    ''' </summary>
    Public MaterialRowID As Integer

    ''' <summary>
    ''' 是否是物料节点
    ''' </summary>
    Public IsMaterial As Boolean
    ''' <summary>
    ''' 模板中替换位置列表
    ''' </summary>
    Public MaterialRowIDList As List(Of Integer)
    ''' <summary>
    ''' 替换物料信息
    ''' </summary>
    Public MaterialValue As MaterialInfo

    ''' <summary>
    ''' 选项值ID
    ''' </summary>
    Public SelectedValueID As String
    ''' <summary>
    ''' 选项值
    ''' </summary>
    Public SelectedValue As String

End Class
