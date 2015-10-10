' --------------------------
' * Excel to Json Exporter *
' --------------------------
'  - 만든이: Koo <koo@kormail.net>
'  - 라이센스: MIT
'  - 주의사항: 이 코드는 어떠한 정상적인 동작도 보증하지 않습니다. 오동작 가능성이 있으며 테스트를 충분히 수행하지 아니 하였습니다.
'  - 사용법
' ----------------------------------------------------------------------------------------------------------------------------------
'    1. 전체 코드를 복사해서 Excel 파일의 현재_통합_문서의 VBAProject에 붙여 넣습니다.(일반), (선언) 이부분에 넣으면 됩니다.
'    2. Config() 프로시저에서 OUTPUT_PATH 부분에 파일을 저장하려는 경로를 넣습니다. 현재 엑셀 파일이 있는 경로를 사용하려면
'       OUTPUT_PATH = Application.ActiveWorkbook.Path를 입력하면 되고, C:\data에 저장하려면 OUTPUT_PATH="C:\data"라고 적으면 됩니다.
'    3. IGNORE_EMPTY를 설정합니다.
'       IGNORE_EMPTY가 True인 경우, 값이 0 혹은 "" 인 것은 저장되지 않고 생략 됩니다.
'    4. ONE_LINE을 설정합니다.
'       ONE_LINE이 True인 경우에는 한 행 당 한 줄로 데이터가 저장되며, False인 경우에는 여러 줄로 저장 됩니다.
'    5. KV_TYPE을 설정합니다.
'       KV_TYPE이 True이면 엑셀 시트에서 첫 번째 열의 값을 key로 하는 dictionary 형태로 저장되며
'       KV_TYPE이 False이면 키가 없이 list 형태로 저장 됩니다.
'    6. ENC_TYPE을 설정합니다.
'       "utf-8", "euc-kr" 등 원하는 저장 인코딩 방식을 선택 합니다.
'    7. 엑셀을 다른으름으로 저장하여 Excel 매크로 사용 통합 문서 형식으로 저장 합니다.
'    8. 엑셀 문서를 저장하면 자동으로 OUTPUT_PATH에 <시트이름>.json으로 각각 저장 됩니다.
' ----------------------------------------------------------------------------------------------------------------------------------

Public OUTPUT_PATH As String
Public IGNORE_EMPTY As Boolean
Public ONE_LINE As Boolean
Public KV_TYPE As Boolean
Public ENC_TYPE As String
Public IGNORE_EXCEPT As Variant

Public TABS As String
Public ENDL As String
Public QUOT As String
Public SPACE As String
Public COMMA As String

Private Sub Config()
    'OUTPUT_PATH = "C:\data"
    OUTPUT_PATH = Application.ActiveWorkbook.Path
    IGNORE_EMPTY = True
    ONE_LINE = True
    KV_TYPE = True
    ENC_TYPE = "utf-8" ' or "euc-kr"
    IGNORE_EXCEPT = Array("CharExp:exp") ' 이 배열에 "시트명:컬럼명" 형태로 추가하면 해당 컬럼은 값이 0이더라도 생략하지 않고 표시해준다.
End Sub

Function GetLine(ByRef sheet As Worksheet, ByVal iRow As Integer)
    Dim iCol, iColMax, iStart As Integer
    Dim res As String
    Dim key As String
    Dim val As String
    Dim bIgnore As Boolean
    
    iColMax = sheet.Cells(1, sheet.Columns.Count).End(xlToLeft).Column
    res = "{"
    
    If KV_TYPE Then
        iStart = 2
    Else
        iStart = 1
    End If
    
    For iCol = iStart To iColMax
        bIgnore = False
        
        If Left(sheet.Cells(1, iCol), 1) = "_" Then
            bIgnore = True
        End If
        
        If Not bIgnore Then
            key = QUOT + sheet.Cells(1, iCol) + QUOT
            
            If IsNumeric(sheet.Cells(iRow, iCol)) Then
                If (IGNORE_EMPTY And UBound(Filter(IGNORE_EXCEPT, sheet.Name + ":" + sheet.Cells(1, iCol))) >= 0) Then
                    val = Trim(Str(sheet.Cells(iRow, iCol)))
                ElseIf IGNORE_EMPTY And sheet.Cells(iRow, iCol) = 0 Then
                    bIgnore = True
                Else
                    val = Trim(Str(sheet.Cells(iRow, iCol)))
                End If
            Else
                val = Trim(sheet.Cells(iRow, iCol))
                If Len(val) >= 2 And Left(val, 1) = "[" And Right(val, 1) = "]" Then
                    ' [로 시작해서 ]로 끝나는것은 리스트로 간주하여 스트링으로 저장하지 않고 리스트로 저장한다.
                    val = Trim(sheet.Cells(iRow, iCol))
                ElseIf IGNORE_EMPTY And Trim(sheet.Cells(iRow, iCol)) = "" Then
                    bIgnore = True
                Else
                    val = QUOT + Trim(sheet.Cells(iRow, iCol)) + QUOT
                End If
            End If
            
            If Not bIgnore Then
                If iCol > iStart Then
                    res = res + COMMA
                End If
                If ONE_LINE Then
                    res = res + SPACE
                Else
                    res = res + ENDL
                End If
                If Not ONE_LINE Then
                    res = res + TABS + TABS
                End If
                res = res + key + ": " + val
            End If
        End If
    Next iCol
    If Not ONE_LINE And Right(res, 1) <> "{" Then
        res = res + ENDL
    End If
    If ONE_LINE Then
        GetLine = res + " }"
    Else
        GetLine = res + TABS + "}"
    End If
End Function

Private Sub SaveToJson(ByRef sheet As Worksheet)
    Dim iRow, iRowMax As Integer
    Dim key As String
    Dim val, vals As String
    Dim filepath As String
    Dim fs
    
    filepath = OUTPUT_PATH + "\" + sheet.Name
   
    If KV_TYPE Then
        vals = "{" + ENDL
    Else
        vals = "[" + ENDL
    End If
    
    iRowMax = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
    For iRow = 2 To iRowMax
        key = QUOT + Trim(Str(sheet.Cells(iRow, 1))) + QUOT
        val = GetLine(sheet, iRow)
        If iRow < iRowMax Then
            If KV_TYPE Then
                vals = vals + TABS + key + ": " + val + COMMA + ENDL
            Else
                vals = vals + TABS + val + COMMA + ENDL
            End If
        Else
            If KV_TYPE Then
                vals = vals + TABS + key + ": " + val + ENDL
            Else
                vals = vals + TABS + val + ENDL
            End If
        End If
    Next iRow
    If KV_TYPE Then
        vals = vals + "}" + ENDL
    Else
        vals = vals + "]" + ENDL
    End If
    
    
    Set fs = CreateObject("ADODB.Stream")
    With fs
        .Charset = ENC_TYPE
        .Open
        .WriteText vals
        .SaveToFile filepath + ".json", 2
    End With
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim i As Integer
    
    TABS = "  "
    ENDL = vbCrLf
    QUOT = """"
    SPACE = " "
    COMMA = ","
    
    Config
    
    For i = 1 To ActiveWorkbook.Worksheets.Count
        SaveToJson ActiveWorkbook.Worksheets(i)
    Next i
End Sub
