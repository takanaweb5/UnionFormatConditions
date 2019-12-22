Attribute VB_Name = "UnionFCondition"
Option Explicit
'Option Private Module

Private Type TSaveInfo
    NewAppliesTo As Range   '再設定させるセル範囲
    Delete       As Boolean 'True:削除対象
End Type

'*****************************************************************************
'[概要] Debug用のセル関数
'[引数] objCell:条件付き書式の設定されたセル、lngNum:FormatConditionsの何番目？
'[戻値] 例：Type:1 Operator:4 TextOperator:# Text:# Formula1:=0 Formula2:#  Formula1:=0 Formula2:# AppliesTo:A1:A20
'*****************************************************************************
Public Function GetFConditionStr(objCell As Range, lngNum As Long) As String
    Dim objFCondition As Object
    Set objFCondition = objCell.FormatConditions(lngNum)
        
    Dim s(1 To 11)
    Dim i As Long
    For i = 1 To UBound(s)
        s(i) = "#" 'エラーの時
    Next
    
    On Error Resume Next
    With objFCondition
        s(1) = .Type
        s(2) = TypeName(objFCondition)
        s(3) = .Operator
        s(4) = .TextOperator
        s(5) = .Text
        s(6) = .Formula1
        s(7) = .Formula2
        s(8) = Application.ConvertFormula(.Formula1, xlA1, xlR1C1, , GetTopLeftCell(.AppliesTo))
        s(9) = Application.ConvertFormula(.Formula2, xlA1, xlR1C1, , GetTopLeftCell(.AppliesTo))
        s(10) = .AppliesTo.AddressLocal(False, False)
        s(11) = GetTopLeftCell(.AppliesTo).AddressLocal(False, False)
    End With
    On Error GoTo 0
    
    Dim strMsg As String
    strMsg = "Type:{1} TypeName:{2} Operator:{3} TextOperator:{4} Text:{5} Formula1:{6} Formula2:{7}  Formula1:{8} Formula2:{9} AppliesTo:{10} TopLeftCell:{11}"
    For i = 1 To UBound(s)
        strMsg = Replace(strMsg, "{" & i & "}", s(i))
    Next
    GetFConditionStr = strMsg
End Function

Public Sub FormatConditions()
    Call UnionFormatConditions(ActiveSheet)
End Sub

'*****************************************************************************
'[概要] ワークシート内の条件付き書式を統合する
'[引数] 対象のワークシート
'[戻値] なし
'*****************************************************************************
Private Sub UnionFormatConditions(ByRef objWorksheet As Worksheet)
    Dim FConditions As FormatConditions
    Set FConditions = objWorksheet.Cells.FormatConditions
    If FConditions.Count = 0 Then
        Exit Sub
    End If
    
    ReDim SaveArray(1 To FConditions.Count) As TSaveInfo
    Dim i As Long
    Dim j As Long
    
    '条件付き書式を後方からLOOPし、統合出来るかどうかの情報をSaveArrayに設定
    For i = FConditions.Count To 1 Step -1
        For j = 1 To i - 1
            If IsSameFormatCondition(FConditions(i), FConditions(j)) Then
                '(i)と(j)が等しければ、後方の(i)を削除して、前方の(j)に統合
                If SaveArray(j).NewAppliesTo Is Nothing Then
                    Set SaveArray(j).NewAppliesTo = Application.Union(FConditions(i).AppliesTo, FConditions(j).AppliesTo)
                Else
                    Set SaveArray(j).NewAppliesTo = Application.Union(FConditions(i).AppliesTo, SaveArray(j).NewAppliesTo)
                End If
                SaveArray(i).Delete = True
            End If
        Next
    Next
    
    '条件付き書式を後方から削除し、前方の条件付き書式に統合
    For i = FConditions.Count To 1 Step -1
        If SaveArray(i).Delete = True Then
            Call FConditions(i).Delete
        Else
            If Not (SaveArray(i).NewAppliesTo Is Nothing) Then
                '念のため、A1,A2,A3 → A1:A3 とするおまじない(不要かも？)
                Dim objWk As Range
                Set objWk = SaveArray(i).NewAppliesTo
                Set objWk = Application.Intersect(objWk, objWk)

                '条件付き書式の統合
                Call FConditions(i).ModifyAppliesToRange(objWk)
            End If
        End If
    Next
End Sub

'*****************************************************************************
'[概要] 一番左上のセルを取得する
'[引数] 条件付き書式の適用範囲
'[戻値] 一番左上のセル
'*****************************************************************************
Private Function GetTopLeftCell(ByRef objRange As Range) As Range
    Dim objArea As Range
    Dim lngRow As Long
    Dim lngCol As Long
    lngRow = Rows.Count
    lngCol = Columns.Count
    
    For Each objArea In objRange.Areas
        With objArea.Cells(1)
            lngRow = WorksheetFunction.Min(.Row, lngRow)
            lngCol = WorksheetFunction.Min(.Column, lngCol)
        End With
    Next
    Set GetTopLeftCell = Cells(lngRow, lngCol)
End Function

'*****************************************************************************
'[概要] 条件および書式が一致するか判定
'[引数] 比較対象のFormatConditionオブジェクト
'[戻値] True:一致
'*****************************************************************************
Private Function IsSameFormatCondition(ByRef F1 As Object, ByRef F2 As Object) As Boolean
    IsSameFormatCondition = False
    If Not (TypeOf F1 Is FormatCondition) Then
        Exit Function
    End If
    If Not (TypeOf F2 Is FormatCondition) Then
        Exit Function
    End If

    Dim FCondition1 As FormatCondition
    Dim FCondition2 As FormatCondition
    Set FCondition1 = F1
    Set FCondition2 = F2
    
    If FCondition1.Type <> FCondition2.Type Then
        Exit Function
    End If
    
'    Select Case FCondition1.Type
'        'セルの値、数式、文字列、期間 のみ判定対象とする
'        Case xlCellValue, xlExpression, xlTextString, xlTimePeriod
'        Case Else
'            Exit Function
'    End Select
    
    '条件が一致するか判定
    Dim Operator(1 To 2)      As String '次の値に等しい、次の値の間etc
    Dim TextOperator(1 To 2)  As String 'Type=xlTextStringの時、次の値を含む、次の値で始まるetc
    Dim Text(1 To 2)          As String 'Type=xlTextStringの時の文字列
    Dim Formula1_R1C1(1 To 2) As String '数式をR1C1タイプで設定
    Dim Formula2_R1C1(1 To 2) As String '数式をR1C1タイプで設定
    
    'タイプによっては直接判定すると例外となる項目があるため例外を抑制し変数に設定
    On Error Resume Next
    With FCondition1
        Operator(1) = .Operator
        TextOperator(1) = .TextOperator
        Text(1) = .Text
        Formula1_R1C1(1) = Application.ConvertFormula(.Formula1, xlA1, xlR1C1, , GetTopLeftCell(.AppliesTo))
        Formula2_R1C1(1) = Application.ConvertFormula(.Formula2, xlA1, xlR1C1, , GetTopLeftCell(.AppliesTo))
    End With
    With FCondition2
        Operator(2) = .Operator
        TextOperator(2) = .TextOperator
        Text(2) = .Text
        Formula1_R1C1(2) = Application.ConvertFormula(.Formula1, xlA1, xlR1C1, , GetTopLeftCell(.AppliesTo))
        Formula2_R1C1(2) = Application.ConvertFormula(.Formula2, xlA1, xlR1C1, , GetTopLeftCell(.AppliesTo))
    End With
    On Error GoTo 0
    
    If Operator(1) <> Operator(2) Then
        Exit Function
    End If
    If TextOperator(1) <> TextOperator(2) Then
        Exit Function
    End If
    If Text(1) <> Text(2) Then
        Exit Function
    End If
    If Formula1_R1C1(1) <> Formula1_R1C1(2) Then
        Exit Function
    End If
    If Formula2_R1C1(1) <> Formula2_R1C1(2) Then
        Exit Function
    End If
    
    
    '書式が一致するか判定
    Dim FontBold(1 To 2)      As String 'フォント太字
    Dim FontColor(1 To 2)     As String 'フォント色
    Dim InteriorColor(1 To 2) As String '塗りつぶし色
    Dim NumberFormat(1 To 2)  As String '値の表示形式 例：#,##0
    
    '場合によっては直接判定すると例外となる項目があることを考慮して例外を抑制し変数に設定
    On Error Resume Next
    With FCondition1
        FontBold(1) = .Font.Bold
        FontColor(1) = .Font.Color
        InteriorColor(1) = .Interior.Color
        NumberFormat(1) = .NumberFormat
    End With
    With FCondition2
        FontBold(2) = .Font.Bold
        FontColor(2) = .Font.Color
        InteriorColor(2) = .Interior.Color
        NumberFormat(2) = .NumberFormat
    End With
    On Error GoTo 0
    
    If FontBold(1) <> FontBold(2) Then
        Exit Function
    End If
    If FontColor(1) <> FontColor(2) Then
        Exit Function
    End If
    If InteriorColor(1) <> InteriorColor(2) Then
        Exit Function
    End If
    If NumberFormat(1) <> NumberFormat(2) Then
        Exit Function
    End If
    
    IsSameFormatCondition = True
End Function

