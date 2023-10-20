Attribute VB_Name = "Module1"
Option Explicit
Sub リンク切れ確認()
    Dim ファイル場所 As String
    Dim 作業ブック As Workbook
    Dim シート As Worksheet
    Dim 定義名探査文 As String, 入力規則探査文 As String
    Dim 範囲 As Variant
    'ダイアログボックスから対象ファイル選択
    ファイル場所 = Application.GetOpenFilename("Excel ブック,*.xls?")
    If ファイル場所 = "False" Then Exit Sub
    'リンク更新を行わない指定でブックを開く
    Set 作業ブック = Workbooks.Open(ファイル場所, UpdateLinks:=0)
    
    '名前定義箇所の探査
    定義名探査文 = 定義名リンク切れ探査(作業ブック)
    Select Case 定義名探査文
        Case "": 定義名探査文 = "名前の定義のリンク切れ箇所候補：無し"
        Case Else: 定義名探査文 = "名前の定義のリンク切れ箇所候補" & 定義名探査文
    End Select
    
    '入力規則設定箇所の探査
    For Each シート In 作業ブック.Worksheets
        On Error Resume Next
        'シート内の入力規則が設定されている全セルを変数格納→存在しない場合エラー
        Set 範囲 = シート.Cells.SpecialCells(xlCellTypeAllValidation)
        If Not (範囲 Is Nothing) Then
            入力規則探査文 = 入力規則探査文 & シート内入力規則リンク切れ探査(シート)
        End If
        On Error GoTo 0
    Next
    Select Case 入力規則探査文
        Case "": 入力規則探査文 = "リスト型入力規則のリンク切れ箇所候補：無し"
        Case Else: 入力規則探査文 = "リスト型入力規則のリンク切れ箇所候補" & 入力規則探査文
    End Select
    
    MsgBox 定義名探査文 & vbCrLf & vbCrLf & 入力規則探査文
    作業ブック.Close SaveChanges:=False
    Set 作業ブック = Nothing
End Sub
Function 定義名リンク切れ探査(作業ブック As Workbook)
    Dim 定義名 As Variant
    For Each 定義名 In 作業ブック.Names
        If InStr(定義名.RefersTo, "#REF") > 0 Or InStr(定義名.RefersTo, ".xl") > 0 Then
            定義名リンク切れ探査 = 定義名リンク切れ探査 & vbCrLf & 定義名.Name & " : " & 定義名.RefersTo
        End If
    Next
End Function
Function シート内入力規則リンク切れ探査(シート As Worksheet) As String
    Dim 範囲 As Variant
    For Each 範囲 In シート.Cells.SpecialCells(xlCellTypeAllValidation)
        '入力規則の種類が「リスト」形式の場合
        If 範囲.Validation.Type = xlValidateList Then
            If InStr(範囲.Validation.Formula1, "#REF") > 0 Or InStr(範囲.Validation.Formula1, ".xl") > 0 Then
                シート内入力規則リンク切れ探査 = シート内入力規則リンク切れ探査 & vbCrLf & シート.Name & " : " & 範囲.Address(False, False) & 範囲.Validation.Formula1
            End If
        End If
    Next
End Function
