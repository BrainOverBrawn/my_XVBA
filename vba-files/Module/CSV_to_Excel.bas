Attribute VB_Name = "CSV_to_Excel"
Option Explicit

Sub main()
    Dim strPath As String
    Dim encoding As String
    Dim rsltArray As Variant
    Dim row As Long, col As Long

    encoding = GetEncoding

    rsltArray = GetRsltArray("C:\DEV_v02\my_XVBA\csv_files", encoding)

    ActiveSheet.Range("A1").Resize(UBound(rsltArray(1, 1), 1), UBound(rsltArray(1, 1), 2)).Value = rsltArray(1, 1)

    For row = 1 To UBound(rsltArray(2, 1), 1)
        If Not (rsltArray(2, 1)(row, 2) = 0) Then
            Cells(rsltArray(2, 1)(row, 1), rsltArray(2, 1)(row, 2)).NumberFormat = "yyyy/mm/dd hh:mm:ss.000"
        End If
    Next row

End Sub

Function GetRsltArray(folderPath As String, encoding As String) As Variant
    Dim rsltArray() As Variant
    ReDim rsltArray(1 To 2, 1 To 1)

    Dim fso As New FileSystemObject
    Dim folder As folder
    Dim file As file

    Dim row As Long, col As Long, row_pos As Long
    Dim strLine As String
    Dim arrLine As Variant
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    Dim maxCols As Long
    Dim two_spaces As Long, lineCount As Long
    two_spaces = 2

    Dim combinedArray() As Variant
    ReDim combinedArray(1 To 1000, 1 To 1)
    ReDim tmstpPosArray(1 To 1000, 1 To 2)

    row = 0
    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.Files

        lineCount = 0
        row = row + 3
        With adoSt
            .Charset = encoding
            .Open
            .LoadFromFile file.Path

            Do Until .EOS
                strLine = .ReadText(adReadLine)
                Debug.Print strLine
                If strLine = "" Then Exit Do
                    lineCount = lineCount + 1
                    row = row + 1
                    arrLine = Split(Replace(strLine, """", ""), ",")
                    ExpandColumns combinedArray, arrLine, two_spaces, maxCols

                    For col = 1 + two_spaces To UBound(arrLine) + 1 + two_spaces
                        combinedArray(row, col) = IIf(arrLine(col - 1 - two_spaces) = "", ChrW(171) & " NULL " & ChrW(187), arrLine(col - 1 - two_spaces))

                        If IsDateTimeFormat(combinedArray(row, col)) Then
                            row_pos = row_pos + 1
                            tmstpPosArray(row_pos, 1) = row
                            tmstpPosArray(row_pos, 2) = col
                        End If
                    Next col
                Loop
                .Close
            End With
            ' ƒtƒ@ƒCƒ‹–¼
            combinedArray(row - lineCount, 1 + two_spaces) = file.Name
            Debug.Print "CSV import completed. " & row & " rows processed.", vbInformation
        Next file

        rsltArray(1, 1) = combinedArray
        rsltArray(2, 1) = tmstpPosArray
        GetRsltArray = rsltArray
End Function

Function IsDateTimeFormat(Byval strValue As String) As Boolean
    Dim regEx As Object
    Dim strPattern As String

    Set regEx = CreateObject("VBScript.RegExp")
    strPattern = "^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}\.\d{3}$"
    With regEx
        .Global = False
        .MultiLine = False
        .IgnoreCase = False
        .Pattern = strPattern
    End With

    IsDateTimeFormat = regEx.test(strValue)

    Set regEx = Nothing
End Function

Function GetEncoding() As String
    Dim response As VbMsgBoxResult
    response = MsgBox("SJIS", vbYesNo + vbQuestion, "Confirmation")
    GetEncoding = IIf(response = vbYes, "SJIS", "UTF-8")
End Function

Function ExpandColumns(combinedArray As Variant, arrLine As Variant, two_spaces As Long, maxCols As Long)
    If UBound(arrLine) + 1 + two_spaces > maxCols Then
        maxCols = UBound(arrLine) + 1 + two_spaces
        ReDim Preserve combinedArray(1 To 1000, 1 To maxCols)
    End If
End Function

