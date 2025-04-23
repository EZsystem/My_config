Attribute VB_Name = "mod_tabCopy2Ball"
Option Compare Database
Option Explicit

'-----------------------------------------------------
' モジュール名: mod_tabCopy2Ball
' 目的　　　: clsTableTransferSetting を使って Icube_累計 から関係テーブルへ一括転写
'　　　　　   転写先に重複データがあれば事前に削除してから転写する
' 使用クラス: clsTableTransferSetting, clsTableTransferExecutor
'-----------------------------------------------------

Public Sub mod_tabCopy2Ball()
    Dim transferList As Collection
    Set transferList = New Collection

    Dim setting As clsTableTransferSetting

    ' B-1: 基本工事_完工
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.TargetTable = "kt_基本工事_完工"
    setting.KeyField = "基本工事コード"
    transferList.Add setting

    ' B-2: 基本工事_作業所
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.TargetTable = "kt_基本工事_作業所"
    setting.KeyField = "基本工事コード"
    transferList.Add setting

    ' B-3: 基本工事_受注
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.TargetTable = "kt_基本工事_受注"
    setting.KeyField = "基本工事コード"
    transferList.Add setting

    ' B-4: 工事コード情報
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.TargetTable = "kt_工事コード情報"
    setting.KeyField = "工事コード"
    transferList.Add setting

    ' B-5: 枝番工事
    Set setting = New clsTableTransferSetting
    setting.SourceTable = "Icube_累計"
    setting.TargetTable = "kt_枝番工事"
    setting.KeyField = "枝番工事コード"
    transferList.Add setting

    ' 実行クラスを使って一括転写（重複削除を含む版を使用）
    Dim settingItem As clsTableTransferSetting
    Dim db As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim sqlDelete As String
    Set db = CurrentDb

    For Each settingItem In transferList
        Set rsSource = db.OpenRecordset("SELECT DISTINCT [" & settingItem.KeyField & "] FROM [" & settingItem.SourceTable & "]", dbOpenSnapshot)
        Do While Not rsSource.EOF
            sqlDelete = "DELETE FROM [" & settingItem.TargetTable & "] WHERE [" & settingItem.KeyField & "] = '" & Replace(rsSource(settingItem.KeyField), "'", "''") & "'"
            db.Execute sqlDelete, dbFailOnError
            rsSource.MoveNext
        Loop
        rsSource.Close
        Set rsSource = Nothing

        TransferTable settingItem.SourceTable, settingItem.TargetTable, settingItem.KeyField
    Next

    Set db = Nothing
End Sub