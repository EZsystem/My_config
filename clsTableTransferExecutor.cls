VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableTransferExecutor"
Option Compare Database
Option Explicit

'-----------------------------------------------------
' クラス名: clsTableTransferExecutor
' 用途　　: clsTableTransferSetting の一覧を受け取り、
'　　　　   TransferTable 関数を使って一括で転写処理を実行する
'-----------------------------------------------------

Public Sub RunAll(settings As Collection)
    Dim setting As clsTableTransferSetting

    For Each setting In settings
        Debug.Print "■ 転写開始: [" & setting.SourceTable & "] → [" & setting.TargetTable & "]"
        TransferTable setting.SourceTable, setting.TargetTable, setting.KeyField
    Next
End Sub