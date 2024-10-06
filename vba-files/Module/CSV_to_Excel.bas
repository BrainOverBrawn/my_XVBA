Attribute VB_Name = "CSV_to_Excel"
Option Explicit

Sub main()
    Dim strPath As String
    Dim encoding As String
    getCSV_utf8 "C:\DEV_v02\my_XVBA\csv_files\mysql.sample_table.csv", "SJIS"
End Sub

Function getCSV_utf8(strPath As String, encoding As String)

    Dim row As Long, col As Long, row_pos As Long
    Dim strLine As String
    Dim arrLine As Variant
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    Dim maxCols As Long
    Dim two_spaces As Long
    two_spaces = 2


    Dim combinedArray() As Variant
    ReDim combinedArray(1 To 1000, 1 To 1)
    ReDim timestampPos(1 To 1000, 1 To 2)

    row = 0
    With adoSt
        .Charset = encoding
        .Open
        .LoadFromFile strPath
        Do Until .EOS
            strLine = .ReadText(adReadLine)
            Debug.Print strLine
            If strLine = "" Then Exit Do
                row = row + 1

                arrLine = Split(Replace(strLine, """", ""), ",")
                If UBound(arrLine) + two_spaces > maxCols Then
                    maxCols = UBound(arrLine) + two_spaces
                    ReDim Preserve combinedArray(1 To 1000, 1 To maxCols)
                End If


                For col = 1 To UBound(arrLine)

                    combinedArray(row, col + two_spaces) = IIf(arrLine(col) = "", ChrW(171) & " NULL " & ChrW(187), arrLine(col))

                    If IsDateTimeFormat(combinedArray(row, col + two_spaces)) Then
                        row_pos = row_pos + 1
                        timestampPos(row_pos, 1) = row
                        timestampPos(row_pos, 2) = col + two_spaces
                    End If

                Next col

            Loop
            .Close
        End With

        Debug.Print "CSV import completed. " & row & " rows processed.", vbInformation
End Function

Function IsDateTimeFormat(Byval strValue As String) As Boolean
    Dim regEx As Object
    Dim strPattern As String

    ' Create RegEx object
    Set regEx = CreateObject("VBScript.RegExp")

    ' Define the pattern For yyyy/mm/dd hh:mm:ss.000
    strPattern = "^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}\.\d{3}$"

    ' Set RegEx properties
    With regEx
        .Global = False
        .MultiLine = False
        .IgnoreCase = False
        .Pattern = strPattern
    End With

    ' Test If the string matches the pattern
    IsDateTimeFormat = regEx.test(strValue)

    ' Clean up
    Set regEx = Nothing
End Function