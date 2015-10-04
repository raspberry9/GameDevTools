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
'    6. 엑셀을 다른으름으로 저장하여 Excel 매크로 사용 통합 문서 형식으로 저장 합니다.
' ----------------------------------------------------------------------------------------------------------------------------------

Public OUTPUT_PATH As String
Public IGNORE_EMPTY As Boolean
Public ONE_LINE As Boolean
Public KV_TYPE As Boolean

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
    KV_TYPE = False
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
                If IGNORE_EMPTY And sheet.Cells(iRow, iCol) = 0 Then
                    bIgnore = True
                Else
                    val = Trim(Str(sheet.Cells(iRow, iCol)))
                End If
            Else
                If IGNORE_EMPTY And Trim(sheet.Cells(iRow, iCol)) = "" Then
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
        GetLine = res + "}"
    Else
        GetLine = res + TABS + "}"
    End If
End Function

Private Sub SaveToJson(ByRef sheet As Worksheet)
    Dim iRow, iRowMax As Integer
    Dim key As String
    Dim val As String
    Dim filepath As String
    
    filepath = OUTPUT_PATH + "\" + sheet.Name
   
    Open filepath + ".json" For Output As #1
    If KV_TYPE Then
        Print #1, "{"
    Else
        Print #1, "["
    End If
    Close #1
    
    Open filepath + ".json" For Append As #1
    iRowMax = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
    For iRow = 2 To iRowMax
        key = QUOT + Trim(Str(sheet.Cells(iRow, 1))) + QUOT
        val = GetLine(sheet, iRow)
        If iRow < iRowMax Then
            If KV_TYPE Then
                Print #1, TABS + key + ": " + val + COMMA
            Else
                Print #1, TABS + val + COMMA
            End If
        Else
            If KV_TYPE Then
                Print #1, TABS + key + ": " + val
            Else
                Print #1, TABS + val
            End If
        End If
    Next iRow
    If KV_TYPE Then
        Print #1, "}"
    Else
        Print #1, "]"
    End If
    Close #1
    
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




