Option Explicit

Type IDDobj
  name As String
  firstField As Long
  lastField As Long
End Type

Dim objToMod() As IDDobj
Dim numObjToMod As Integer

Dim iddField() As String
Dim numIddFields As Long
Dim sizeIddFields As Long

Dim idfObjectStrings() As String
Dim numIdfObjectStrings As Long
Dim sizeIdfObjectStrings As Long

Sub testDrumlin()
Debug.Print ""
Debug.Print ""
Debug.Print "================= TEST at ", Time()
Call getListOfObjects
Call ReadIDDforFields
Call readIDF
Call putObjectsOnTab
End Sub

Sub getListOfObjects()
numObjToMod = 2
ReDim objToMod(numObjToMod)
objToMod(1).name = "LIGHTS"
objToMod(2).name = "FENESTRATIONSURFACE:DETAILED"
End Sub

Sub readIDF()
Dim InFN As Integer
Dim DataLine As String
Dim lineCount As Integer

Dim exclaimPos As Integer
Dim commaPos As Integer
Dim semiPos As Integer
Dim lineNoComment As String

Dim inObject As Boolean
Dim objString As String
Dim objParts() As String
Dim iNumObjToMod As Integer
Dim found As Boolean

InFN = FreeFile()
'Open "d:\personal\dev\drumlin-EnergyPlusSpreadsheet\5ZoneAirCooled.idf" For Input As InFN
'Open "C:\Users\jglazer\Documents\personal\dev\drumlin-EnergyPlusSpreadsheet\5ZoneAirCooled.idf" For Input As InFN
Dim spreadSheetLocation As String
spreadSheetLocation = Application.ActiveWorkbook.Path
Open spreadSheetLocation & "\5ZoneAirCooled.idf" For Input As InFN
lineCount = 0
inObject = False

numIdfObjectStrings = 0
sizeIdfObjectStrings = 500
ReDim idfObjectStrings(sizeIdfObjectStrings)
Do While Not EOF(InFN)
    Line Input #InFN, DataLine ' read in data 1 line at a time
    lineCount = lineCount + 1
    DataLine = Trim(DataLine)
    exclaimPos = InStr(DataLine, "!")
    commaPos = InStr(DataLine, ",")
    semiPos = InStr(DataLine, ";")
    If exclaimPos > 0 Then
        lineNoComment = Left(DataLine, exclaimPos - 1)
    Else
        lineNoComment = DataLine
    End If
    lineNoComment = Trim(lineNoComment)
    objString = objString & lineNoComment
    If semiPos > 0 Then
        If Len(objString) > 0 Then
            If Right(objString, 1) = ";" Then
                objString = Left(objString, Len(objString) - 1)
            End If
            'MsgBox objString
            objParts = Split(objString, ",")
            'MsgBox objParts(0)
            For iNumObjToMod = 1 To numObjToMod
                If UCase(objParts(0)) = objToMod(iNumObjToMod).name Then
                    found = True
                    numIdfObjectStrings = numIdfObjectStrings + 1
                    If numIdfObjectStrings > sizeIdfObjectStrings Then
                        sizeIdfObjectStrings = sizeIdfObjectStrings + 500
                        ReDim Preserve idfObjectStrings(sizeIdfObjectStrings)
                    End If
                    idfObjectStrings(numIdfObjectStrings) = objString
                    'MsgBox objString
                End If
            Next iNumObjToMod
        End If
        objString = ""
    End If
Loop 'not eof(infn)
Debug.Print "number of IDF lines: " & lineCount
End Sub

Sub ReadIDDforFields()
Dim iddFN As Integer
Dim fileLine As String
Dim lineCount As Long
Dim iObj As Long
Dim withinObject As Boolean
Dim iddFileName As String
Dim posSlash As Long
Dim posSemi As Long
Dim fieldNm As String
Dim posField As Long
Dim numObjFound As Long
Dim objFound As Long

numIddFields = 0
sizeIddFields = 2000
ReDim iddField(sizeIddFields)

numObjFound = 0
iddFN = FreeFile()
iddFileName = "C:\EnergyPlusV8-7-0\Energy+.idd"
Open iddFileName For Input As iddFN
withinObject = False
Do While Not EOF(iddFN)
    Line Input #iddFN, fileLine ' read in data 1 line at a time
    lineCount = lineCount + 1
    If Not withinObject Then
        For iObj = 1 To numObjToMod
            If UCase(fileLine) = objToMod(iObj).name + "," Then
                withinObject = True
                objFound = iObj
                Debug.Print "found", objToMod(iObj).name, lineCount
            End If
        Next iObj
    Else 'within the object
        posField = InStr(fileLine, "\field")
        If posField > 0 Then
            fieldNm = Mid(fileLine, posField + 7)
            numIddFields = numIddFields + 1
            iddField(numIddFields) = fieldNm
            If objToMod(objFound).firstField = 0 Then objToMod(objFound).firstField = numIddFields
            objToMod(objFound).lastField = numIddFields
            Debug.Print "field["; fieldNm; "]"
        End If
        'switch to outside of object if semicolon is found
        posSlash = InStr(fileLine, "\")
        posSemi = InStr(fileLine, ";")
        If posSemi > 0 Then
            If posSlash > posSemi Then
                withinObject = False
                numObjFound = numObjFound + 1
            End If
        End If
    End If
    If numObjFound = numObjToMod Then
        Exit Do
    End If
Loop 'not eof(iddfn)
Debug.Print "last line scanned in IDD: " & lineCount
Close iddFN
End Sub


Sub putObjectsOnTab()
Dim iObj As Long
Dim jField As Long
Dim kObjToMod As Long
Dim pieces() As String
Dim nRow As Long
Dim nCol As Long
Dim maxRowsForObj As Long
Dim fieldStart As Long

Application.ScreenUpdating = False
nRow = 10
For kObjToMod = 1 To numObjToMod
    maxRowsForObj = 0
    nCol = 3
    For iObj = 1 To numIdfObjectStrings
        pieces = Split(idfObjectStrings(iObj), ",")
        If UCase(pieces(0)) = objToMod(kObjToMod).name Then
            nCol = nCol + 1
            For jField = LBound(pieces) To UBound(pieces)
                Cells(nRow + jField, nCol).Value = pieces(jField)
            Next jField
            If UBound(pieces) > maxRowsForObj Then
                maxRowsForObj = UBound(pieces)
            End If
        End If
    Next iObj
    'put in the field names
    fieldStart = objToMod(kObjToMod).firstField
    For jField = 0 To maxRowsForObj - 1
        Cells(nRow + jField + 1, 3).Value = iddField(fieldStart + jField)
    Next jField
    nRow = nRow + maxRowsForObj + 5
Next kObjToMod
Application.ScreenUpdating = True
End Sub


