VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableTransferSetting"
Option Compare Database
Option Explicit

'-----------------------------------------
' クラス名: clsTableTransferSetting
' 用途　　: テーブル転写のための設定情報を保持する
' 使用例　: 転写元テーブル、転写先テーブル、照合に使うキーを指定
'-----------------------------------------

' 転写元テーブル名（例："Icube_"）
Public SourceTable As String

' 転写先テーブル名（例："kt_基本工事_完工"）
Public TargetTable As String

' 転写条件フィールド（例："基本工事コード"）
Public KeyField As String