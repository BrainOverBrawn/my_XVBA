Attribute VB_Name = "CSV_to_Excel"
Option Explicit

Sub main()
    Dim strPath As String
    Dim encoding As String
    Dim resultArray As Variant
    resultArray = getCSV_utf8("C:\DEV_v02\my_XVBA\csv_files", "SJIS")
End Sub

Function getCSV_utf8(folderPath As String, encoding As String) As Variant

    Dim fso As New FileSystemObject
    Dim folder As folder
    Dim file As file

    Dim row As Long, col As Long, row_pos As Long
    Dim strLine As String
    Dim arrLine As Variant
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    Dim maxCols As Long, minCols As Long
    Dim two_spaces As Long, lineCount As Long
    two_spaces = 2

    Dim combinedArray() As Variant
    ReDim combinedArray(1 To 1000, 1 To 1)
    ReDim timestampPos(1 To 1000, 1 To 2)

    row = 0
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files


        lineCount = 0
        With adoSt
            .Charset = encoding
            .Open
            .LoadFromFile file.Path

            Do Until .EOS
                lineCount = lineCount + 1
                strLine = .ReadText(adReadLine)
                Debug.Print strLine
                If strLine = "" Then Exit Do


                    row = row + 1
                    arrLine = Split(Replace(strLine, """", ""), ",")
                    minCols = UBound(arrLine) + 1 + two_spaces
                    If minCols > maxCols Then
                        maxCols = minCols
                        ReDim Preserve combinedArray(1 To 1000, 1 To maxCols)
                    End If

                    If lineCount = 1 Then
                        combinedArray(row, 1 + two_spaces) = file.Name
                        row = row + 1
                    End If


                    For col = 1 + two_spaces To minCols

                        combinedArray(row, col) = IIf(arrLine(col - 1 - two_spaces) = "", ChrW(171) & " NULL " & ChrW(187), arrLine(col - 1 - two_spaces))

                        If IsDateTimeFormat(combinedArray(row, col)) Then
                            row_pos = row_pos + 1
                            timestampPos(row_pos, 1) = row
                            timestampPos(row_pos, 2) = col
                        End If

                    Next col

                Loop
                .Close
            End With
            Debug.Print "CSV import completed. " & row & " rows processed.", vbInformation
            row = row + 2
        Next file
        getCSV_utf8 = combinedArray
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