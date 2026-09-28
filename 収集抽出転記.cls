VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "収集抽出転記"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
#Const 本番 = 0
#Const デバグ = 1

Option Explicit

Private Const Excel最大行 = 1048576

Dim the見出アドレス As String
Dim the最終行判定列 As Long
Dim フィルタ As AutoFilter
Dim 転記先 As New 転記先

Private Sub Class_Initialize()
  On Error GoTo OnError
  With Excel処理.the仕様
    the見出アドレス = .見出アドレス
    the最終行判定列 = .最終行判定列
  End With
  Dim 仕様AsWorksheet As Worksheet
  Set 仕様AsWorksheet = Excel処理.the仕様
  With 仕様AsWorksheet
    If .AutoFilterMode Then Set フィルタ = .AutoFilter
  End With
Exit Sub
OnError:
  With Err()
    .Source = "巡庫.収集抽出転記.Class_Initialize"
    Debug.Print Now, .Number, .Source, .Description
    .Raise -1
  End With
End Sub

Public Property Get ルートフォルダパス() As String

End Property

Public Property Get 探索対象(フォルダ As Folder) As 対象

End Property

Public Property Get 処理対象ファイル(ファイル As File) As Boolean

End Property

Public Property Get 処理対象シート(シート As Worksheet) As Boolean

End Property

Public Property Get 見出アドレス() As String

End Property

Public Property Get 最終行判定列() As Long

End Property

Friend Sub 処理する(シート As Worksheet)
  On Error GoTo OnError
  With シート
    .AutoFilterMode = False
    Dim 最終行 As Long
    最終行 = .Cells(Excel最大行, 最終行判定列).End(xlUp).Row
    
    Dim 表領域 As Range
    Set 表領域 = .Range(見出アドレス)
    Set 表領域 = .Range(見出アドレス).Resize(最終行 - 表領域.Cells(1).Row() + 1&)
  End With
  
  If Not フィルタ Is Nothing Then
    With フィルタ.Filters
      Dim i As Long
On Error Resume Next
      For i = 1 To .Count
        With .Item(i)
          If .On Then
            表領域.AutoFilter i, .Criteria1
            表領域.AutoFilter i, .Criteria1, .Operator
            表領域.AutoFilter i, .Criteria1, .Operator, .Criteria2
          End If
        End With
      Next
On Error GoTo OnError
    End With
  End If

  With 表領域
    If .Rows.Count > 1& Then
      転記先.転記する .Offset(1).Resize(.Rows.Count - 1&).Value
    End If
  End With
Exit Sub
OnError:
  With Err()
    If .Number = 0 Then Exit Sub
    Debug.Print Now, .Number, "巡庫.収集抽出転記.処理する", .Description
    .Raise .Number
  End With
End Sub

Public Sub 転記先Initialize()

End Sub

Public Sub 転記する(データ As Variant)

End Sub

Public Sub 転記先Terminate()

End Sub

Private Sub Class_Terminate()
  On Error GoTo OnError
  Set 転記先 = Nothing
  Set フィルタ = Nothing
Exit Sub
OnError:
  With Err()
    Debug.Print Now, .Number, "巡庫.収集抽出転記.Class_Terminate", .Description
  End With
  Resume Next
End Sub
