
Attribute VB_Name = "mod_tabCopy2_edaban1"
Option Compare Database
Option Explicit

Public Sub mod_tabCopy2_edaban1()
    On Error GoTo ErrHandler

    ' --- 初期化 ---
    Dim cleaner As New clsTableCleaner
    Dim logger As New clsErrorLogger
    Dim delCond As New clsDeleteCondition
    Dim db As DAO.Database
    Set db = CurrentDb

    ' --- ロガー初期化 ---
    logger.Init

    ' --- 一時テーブル初期化 ---
    cleaner.Init
    cleaner.AddTable "Icube_Temp"
    cleaner.Clear

    ' --- 転写元から一時テーブルへデータ挿入 ---
    db.Execute "INSERT INTO Icube_Temp SELECT * FROM Icube_;", dbFailOnError

    ' --- 削除条件設定（枝番工事コードが一致するもの） ---
    delCond.targetField = "枝番工事コード"
    delCond.OperatorStr = "IN"
    delCond.TargetValue = "(SELECT 枝番工事コード FROM Icube_Temp)"

    ' --- 転写先（Icube_累計）の削除処理 ---
    Dim deleteSQL As String
    deleteSQL = "DELETE FROM Icube_累計 " & vbCrLf & _
                "WHERE " & delCond.targetField & " " & delCond.OperatorStr & " " & delCond.TargetValue

    db.Execute deleteSQL, dbFailOnError
    logger.LogError "削除実行", "Icube_累計から重複レコードを削除にゃ：" & vbCrLf & deleteSQL

    ' --- 一時テーブルから転写先へ追加 ---
    db.Execute "INSERT INTO Icube_累計 SELECT * FROM Icube_Temp;", dbFailOnError

    ' --- 完了ログ ---
    logger.LogError "処理完了", "データ転写が完了したにゃ"

    Exit Sub

ErrHandler:
    logger.LogError "エラー発生", "内容: " & Err.Description
    logger.ShowAllErrors True
End Sub
