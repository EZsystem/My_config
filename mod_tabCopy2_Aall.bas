Attribute VB_Name = "mod_tabCopy2_Aall"
Option Compare Database
Option Explicit

'-----------------------------------------------------
' モジュール名: mod_tabCopy2_Aall
' 目的　　　: clsTableTransferSetting を使って Icube_ から関係テーブルへ一括転写
' 使用クラス: clsTableTransferSetting, clsTableTransferExecutor
'-----------------------------------------------------

Public Sub mod_tabCopy2_Aall()
    Dim transferList As Collection
    Set transferList = New Collection

    Dim setting As clsTableTransferSetting

    ' A-1: 基本工事_完工
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_"
    setting.TargetTable = "kt_基本工事_完工"
    setting.KeyField = "基本工事コード"
    transferList.Add setting

    ' A-2: 基本工事_作業所
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_"
    setting.TargetTable = "kt_基本工事_作業所"
    setting.KeyField = "基本工事コード"
    transferList.Add setting

    ' A-3: 基本工事_受注
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_"
    setting.TargetTable = "kt_基本工事_受注"
    setting.KeyField = "基本工事コード"
    transferList.Add setting

    ' A-4: 工事コード情報
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_"
    setting.TargetTable = "kt_工事コード情報"
    setting.KeyField = "工事コード"
    transferList.Add setting

    ' A-5: 枝番工事
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_"
    setting.TargetTable = "kt_枝番工事"
    setting.KeyField = "枝番工事コード"
    transferList.Add setting

    ' 実行クラスを使って一括転写
    Dim executor As New clsTableTransferExecutor
    executor.RunAll transferList
End Sub