VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Excel処理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
#Const 本番 = 0
#Const デバグ = 1

Option Explicit

Implements 訪

Public the仕様 As 収集抽出転記
Private 別エクセル As New 別エクセル
Private シート処理 As New 収集抽出転記

Public Sub ルートフォルダ内のフックを収集抽出転記する(仕様 As 収集抽出転記)
  On Error GoTo OnError
  Set the仕様 = 仕様
  巡.ルートフォルダ内の全ファイルを処理する Me
OnError:
  Set シート処理 = Nothing
  Set 別エクセル = Nothing
  Set the仕様 = Nothing
  With Err()
    If .Number = 0 Then Exit Sub
    .Source = "巡庫.Excel処理.ルートフォルダ内のフックを収集抽出転記する"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Sub

Private Property Get 訪_ルートフォルダパス() As String
  On Error GoTo OnError
  訪_ルートフォルダパス = the仕様.ルートフォルダパス
Exit Property
OnError:
  With Err()
    .Source = "巡庫.Excel処理.訪_ルートフォルダパス"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Property

Private Static Property Get 訪_探索対象(フォルダ As Scripting.IFolder) As 対象
  On Error GoTo OnError
  訪_探索対象 = the仕様.探索対象(フォルダ)
Exit Property
OnError:
  With Err()
    .Source = "巡庫.Excel処理.訪_探索対象"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Property

Private Static Sub 訪_処理する(ファイル As Scripting.IFile)
  On Error GoTo OnError
  If 巡.アプリ(ファイル) <> エクセル Then Exit Sub
  If Not the仕様.処理対象ファイル(ファイル) Then Exit Sub
  Dim シート As Worksheet
  For Each シート In 別エクセル.シート集(ファイル.Path)
    If the仕様.処理対象シート(シート) Then シート処理.処理する シート
  Next シート
OnError:
  With Err()
    .Source = "巡庫.Excel処理.訪_処理する"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Sub
