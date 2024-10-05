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

                arrLine = Split(Replace(strLine, """", ""), ",")

                For j = 0 To UBound(arrLine)
                    Debug.Print IIf(arrLine(j) = "", ChrW(171) & " NULL " & ChrW(187), arrLine(j))
                Next j
                i = i + 1
            Loop

            .Close
        End With

        Debug.Print "CSV import completed. " & (i - 1) & " rows processed.", vbInformation
End Sub
