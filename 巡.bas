Attribute VB_Name = "巡"
#Const 本番 = 0
#Const デバグ = 1

Option Explicit

Public Enum アプリ
  エクセル = 1
  ワード = 2
  パワポ = 3
  その他 = 9
End Enum

Public the指示 As 訪

Public Sub ルートフォルダ内の全ファイルを処理する(指示 As 訪)
  On Error GoTo OnError
  Set the指示 = 指示
  Dim ルートフォルダ As Folder
  With New FileSystemObject
    Set ルートフォルダ = .GetFolder(the指示.ルートフォルダパス)
    With New 訪
      .導く ルートフォルダ
    End With
  End With
OnError:
  Set the指示 = Nothing
  With Err()
    If .Number = 0 Then Exit Sub
    .Source = "巡庫.巡.ルートフォルダ内の全ファイルを処理する"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Sub

Public Static Function 処理する(フォルダ As Folder) As Folders
  On Error GoTo OnError
  With the指示
    If (.探索対象(フォルダ) Imp 対象.Fileを処理) = True Then
      Dim ファイル As File
      For Each ファイル In フォルダ.Files
        .処理する ファイル
      Next ファイル
    End If

    If (.探索対象(フォルダ) Imp 対象.Subfolderを処理) = True Then
      Set 処理する = フォルダ.SubFolders
    Else
      Set 処理する = Nothing
    End If
  End With
Exit Function
OnError:
  With Err()
    .Source = "巡庫.巡.処理する"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Function

Public Static Function アプリ(ファイル As File) As アプリ
  On Error GoTo OnError
  Dim タイプ As String
  タイプ = ファイル.Type
  Select Case True
    Case InStr(タイプ, "Excel") + InStr(タイプ, "クエリ") > 0
      アプリ = エクセル
    Case InStr(タイプ, "csv") > 0
      アプリ = エクセル
    Case InStr(タイプ, "Word") > 0
      アプリ = ワード
    Case InStr(タイプ, "Power") > 0
      アプリ = パワポ
    Case Else
      アプリ = その他
  End Select
Exit Function
OnError:
  With Err()
    .Source = "巡庫.巡.アプリ"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Function
