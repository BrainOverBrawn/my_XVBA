Attribute VB_Name = "CSV_to_Excel"
Sub getCSV_utf8()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)

    Dim strPath As String
    strPath = "C:\DEV_v02\my_XVBA\csv_files\mysql.sample_table.csv"

    Dim i As Long, j As Long
    Dim strLine As String
    Dim arrLine As Variant

    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")

    i = 1
    With adoSt
        .Charset = "SJIS"  ' Changed To UTF-8 As per Function name
        .Open
        .LoadFromFile strPath

        Do Until .EOS
            strLine = .ReadText(adReadLine)
            Debug.Print strLine
            If strLine = "" Then Exit Do  ' Exit If empty line is encountered

                '            arrLine = Split(Replace(replaceColon(strLine), """", ""), ":")
                arrLine = Split(strLine, ",")

                For j = 0 To UBound(arrLine)
                    Debug.Print IIf(arrLine(j) = "", "NULL", arrLine(j))

                    '                ws.Cells(i, j + 1).Value = arrLine(j)  ' Uncommented this line
                Next j
                i = i + 1
            Loop

            .Close
        End With

        Debug.Print "CSV import completed. " & (i - 1) & " rows processed.", vbInformation
End Sub

Function replaceColon(Byval str As String) As String
    Dim strTemp As String
    Dim quotCount As Long

    Dim l As Long
    For l = 1 To Len(str)
        strTemp = Mid(str, l, 1)
        If strTemp = """" Then
            quotCount = quotCount + 1
        Elseif strTemp = "," Then
            If quotCount Mod 2 = 0 Then
                str = Left(str, l - 1) & ":" & Right(str, Len(str) - l)
            End If
        End If
    Next l

    replaceColon = str
End Function
