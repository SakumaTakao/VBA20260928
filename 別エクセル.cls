VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "別エクセル"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#Const 本番 = 0
#Const デバグ = 1

Option Explicit

Private 別エクセル As New Application

Public Static Function シート集(パス As String) As Worksheets
  On Error GoTo OnError
  With 別エクセル.Workbooks
    .Close
    With .Open(パス, ReadOnly:=True)
      Set シート集 = .Worksheets
    End With
  End With
Exit Function
OnError:
  With Err()
    .Source = "巡庫.別エクセル.シート集"
    Debug.Print Now(), .Number, .Source, .Description
    .Raise -1
  End With
End Function

Private Sub Class_Terminate()
  On Error GoTo OnError
  With 別エクセル
    .Workbooks.Close
    .Quit
  End With
  Set 別エクセル = Nothing
Exit Sub
OnError:
  With Err()
    .Source = "巡庫.別エクセル.Class_Terminate"
    Debug.Print Now(), .Number, .Source, .Description
    Resume Next
  End With
End Sub
