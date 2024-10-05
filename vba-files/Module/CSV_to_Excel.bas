Attribute VB_Name = "CSV_to_Excel"
Sub main()
    Dim strPath As String
    Dim encoding As String
    getCSV_utf8 "C:\DEV_v02\my_XVBA\csv_files\mysql.sample_table.csv", "SJIS"
End Sub

Function getCSV_utf8(strPath As String, encoding As String)

    Dim row As Long, col As Long
    Dim strLine As String
    Dim arrLine As Variant
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    Dim maxCols As Long

    Dim combinedArray() As Variant
    ReDim combinedArray(1 To 1000, 1 To 1)

    row = 0
    With adoSt
        .Charset = encoding
        .Open
        .LoadFromFile strPath

        Do Until .EOS
            strLine = .ReadText(adReadLine)
            Debug.Print strLine
            If strLine = "" Then Exit Do  ' Exit If empty line is encountered
                row = row + 1

                arrLine = Split(Replace(strLine, """", ""), ",")

                For col = 0 To UBound(arrLine)
                    If UBound(arrLine) + 1 > maxCols Then
                        maxCols = UBound(arrLine) + 1
                        ReDim Preserve combinedArray(1 To 1000, 1 To maxCols)
                    End If

                    combinedArray(row, col + 1) = IIf(arrLine(col) = "", ChrW(171) & " NULL " & ChrW(187), arrLine(col))

                Next col

            Loop

            .Close
        End With

        Debug.Print "CSV import completed. " & row & " rows processed.", vbInformation
End Function

