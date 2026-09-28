VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "転記先"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const 本番 = 0
#Const デバグ = 1

Option Explicit

Private 実 As 収集抽出転記

Private Sub Class_Initialize()
  On Error GoTo OnError
  
  Dim 収集抽出転記AsWorksheet As Worksheet
  Set 収集抽出転記AsWorksheet = Excel処理.the仕様
  収集抽出転記AsWorksheet.Copy
  With 収集抽出転記AsWorksheet.Application
    Set 実 = .ActiveSheet
  End With
  
  実.転記先Initialize
Exit Sub
OnError:
  With Err()
    .Source = "巡庫.転記先.Class_Initialize"
    Debug.Print Now, .Number, .Source, .Description
    .Raise 440
  End With
End Sub

Public Static Sub 転記する(データ As Variant)
  On Error GoTo OnError
  実.転記する データ
Exit Sub
OnError:
  With Err()
    .Source = "巡庫.転記先.転記する"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Sub

Private Sub Class_Terminate()
  On Error GoTo OnError
  実.転記先Terminate
  Set 実 = Nothing
Exit Sub
OnError:
  With Err()
    .Source = "巡庫.転記先.Class_Terminate"
    Debug.Print Now, .Number, .Source, .Description
    .Raise 440
  End With
End Sub
