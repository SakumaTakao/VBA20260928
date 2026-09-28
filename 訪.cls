VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "訪"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
#Const 本番 = 0
#Const デバグ = 1

Option Explicit

Public Enum 対象
  Fileを処理 = Not 2 ^ 0
  Subfolderを処理 = Not 2 ^ 1
End Enum

Public Property Get ルートフォルダパス() As String

End Property

Public Property Get 探索対象(フォルダ As Folder) As 対象

End Property

Public Sub 処理する(ファイル As File)

End Sub

Friend Static Sub 導く(ルートフォルダ As Folder)
  Static 次訪 As New 訪
  Static サブフォルダ集 As Folders
  Set サブフォルダ集 = 巡.処理する(ルートフォルダ)
  If サブフォルダ集 Is Nothing Then Exit Sub
  Dim サブフォルダ As Folder
  For Each サブフォルダ In サブフォルダ集
    次訪.導く サブフォルダ
  Next サブフォルダ
Exit Sub
OnError:
  With Err()
    .Source = "巡庫.訪.導く"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Sub

#If MODE = デバグ Then
Private Sub Class_Terminate()
  Debug.Print Now(), "「巡.訪.Class_Teminate」を実行しました。"
End Sub
#End If
