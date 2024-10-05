Attribute VB_Name = "Module1"
Option Explicit

Sub ImportAndPasteCSVs()

    Dim result As Variant
    Dim combinedArray As Variant
    Dim dateTimePositions As Variant
    Dim folderPath As String
    Dim i As Long

    ' Specify the folder path containing CSV files
    folderPath = "C:\DEV_v02\my_XVBA\csv_files"

    ' Import CSV files
    result = ImportMultipleCSVs(folderPath)

    ' Extract combined array And date-time positions
    combinedArray = result(0)
    dateTimePositions = result(1)

    ' Paste combined array To the active sheet
'    ActiveSheet.Range("A1").Resize(UBound(combinedArray, 1), UBound(combinedArray, 2)).Value = combinedArray

    ' Highlight cells With date-time format (Optional)
    For i = 1 To UBound(dateTimePositions, 1)
'        ActiveSheet.Cells(dateTimePositions(i, 1), dateTimePositions(i, 2)).Interior.Color = RGB(255, 255, 0)
    Next i

    MsgBox "Import complete. " & UBound(dateTimePositions, 1) & " date-time formatted cells found."
End Sub

Function ImportMultipleCSVs(folderPath As String) As Variant
'    Dim fso As Object
    Dim fso As New FileSystemObject
    Dim folder As Object
    Dim file As file
    Dim fileStream As Object
    Dim combinedArray() As Variant
    Dim dateTimePositions() As Variant
    Dim lineText As String
    Dim lineArray As Variant
    Dim i As Long, j As Long, k As Long
    Dim totalRows As Long
    Dim maxCols As Long

    ' Create FileSystemObject
'    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' Initialize arrays
    ReDim combinedArray(1 To 1000, 1 To 1)
    ReDim dateTimePositions(1 To 1, 1 To 2)

    totalRows = 0
    maxCols = 0

    ' Loop through each CSV file in the folder
    For Each file In folder.Files
    
    



    strPath = file.Path

    Dim i As Long, j As Long
    Dim strLine As String
    Dim arrLine As Variant 'カンマでsplitして格納

    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")

    i = 1
    With adoSt
        .Charset = "UTF-8"        'Streamで扱う文字コートをutf-8に設定
        .Open                             'Streamをオープン
        .LoadFromFile (strPath) 'ファイルからStreamにデータを読み込む

        Do Until .EOS           'Streamの末尾まで繰り返す

            strLine = .ReadText(adReadLine) 'Streamから1行取り込み

            arrLine = Split(Replace(replaceColon(strLine), """", ""), ":") 'strLineをカンマで区切りarrLineに格納

            For j = 0 To UBound(arrLine)

                ws.Cells(i, j + 1).Value = arrLine(j)

            Next j
            i = i + 1
        Loop

        .Close
    End With
    
        
    
    
    
    
    
    
    
    
        Debug.Print file.Type
        If LCase(Right(file.Name, 4)) = ".csv" Then
            ' Open the CSV file
            Set fileStream = fso.OpenTextFile(file.Path, 1, False, -1)

            ' Read the file And populate the array
            Do While Not fileStream.AtEndOfStream
                lineText = fileStream.ReadLine
                lineArray = Split(lineText, ",")

                ' Resize combinedArray If necessary
                If UBound(lineArray) + 1 > maxCols Then
                    maxCols = UBound(lineArray) + 1
                    ReDim Preserve combinedArray(1 To UBound(combinedArray, 1), 1 To maxCols)
                End If

                totalRows = totalRows + 1

                '                ReDim Preserve combinedArray(1 To 1000, 1 To maxCols)

                ' Populate the array And check For date-time format
                For j = 0 To UBound(lineArray)
                    combinedArray(totalRows, j + 1) = Trim(Replace(lineArray(j), """", ""))

                    ' Check If the value matches the date-time format
                    If IsDateTimeFormat(combinedArray(totalRows, j + 1)) Then
                        k = k + 1
                        ReDim Preserve dateTimePositions(1 To k, 1 To 2)
                        dateTimePositions(k, 1) = totalRows
                        dateTimePositions(k, 2) = j + 1
                    End If
                Next j
            Loop

            ' Close the file
            fileStream.Close
        End If
    Next file

    ' Return the populated array And date-time positions
    ImportMultipleCSVs = Array(combinedArray, dateTimePositions)
End Function

Function IsDateTimeFormat(ByVal strValue As String) As Boolean
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

