' add unit conversions
' training

Option Explicit

Dim iddFileName As String
Dim originalFileName As String
Dim revisedFileName As String

Type IDDobj
    name As String
    nameUC As String
    firstField As Long
    lastField As Long
    revisedRow As Long
    revisedCol As Long
End Type
Dim objToMod() As IDDobj
Dim numObjToMod As Integer
Dim sizeObjToMod As Integer

Type IDDfieldType
    name As String
    unitIndex As Long
End Type
Dim IDDfield() As IDDfieldType
Dim numIddFields As Long
Dim sizeIddFields As Long

Type unitsType
    siName As String
    ipName As String
    multiplier As Double
    offset As Double
End Type
Dim units() As unitsType
Dim numUnits As Long
Dim sizeUnits As Long

Const modeRead = 1
Const modeWrite = 2



Dim idfObjectStrings() As String
Dim numIdfObjectStrings As Long
Dim sizeIdfObjectStrings As Long

Sub runDrumlin()
Debug.Print ""
Debug.Print ""
Debug.Print "================= Begin Drumlin at ", Time()
Call initializeParameters
Call ReadIDDforFields
Call readWriteIDF(modeRead)
Call putObjectsOnTab
'recalc the tab
Call readWriteIDF(modeWrite)
Debug.Print "================= End Drumlin at ", Time()
End Sub

Sub initializeParameters()
Dim rowNum As Integer
Dim cParam As String
Dim cValue As String
numObjToMod = 0
sizeObjToMod = 25
ReDim objToMod(sizeObjToMod)
'Debug.Print Cells(1, 1).Value
If UCase(Trim(Cells(1, 1).Value)) = "DRUMLIN" Then
    rowNum = 2
    Do
        cParam = Trim(Cells(rowNum, 1).Value)
        cValue = Trim(Cells(rowNum, 2).Value)
        Debug.Print cParam, cValue
        If cParam = "" Or cValue = "" Then Exit Do
        Select Case cParam
            Case "IDD"
                iddFileName = addPathIfLocal(cValue)
            Case "ORIGINAL"
                originalFileName = addPathIfLocal(cValue)
            Case "REVISED"
                revisedFileName = addPathIfLocal(cValue)
            Case "OBJECT"
                numObjToMod = numObjToMod + 1
                If numObjToMod > sizeObjToMod Then
                    sizeObjToMod = sizeObjToMod + 25
                    ReDim Preserve objToMod(sizeObjToMod)
                End If
                objToMod(numObjToMod).nameUC = UCase(cValue)
        End Select
        rowNum = rowNum + 1
    Loop
End If
End Sub

Function addPathIfLocal(fileNameIn As String) As String
Dim spreadsheetlocation As String
spreadsheetlocation = Application.ActiveWorkbook.Path
If Mid(fileNameIn, 2, 2) = ":\" Then
    addPathIfLocal = fileNameIn
Else
    addPathIfLocal = spreadsheetlocation & "\" & fileNameIn
End If
End Function

Sub readWriteIDF(mode As Integer)
' mode = 1 for read
' mode = 2 for write
Dim InFN As Integer
Dim OutFN As Integer
Dim wrt As Boolean
Dim DataLine As String
Dim trimDataLine As String
Dim lineCount As Integer

Dim lineBuffer() As String
Dim numLineBuffer As Long
Dim sizeLineBuffer As Long

Dim exclaimPos As Integer
Dim commaPos As Integer
Dim semiPos As Integer
Dim lineNoComment As String

Dim inObject As Boolean
Dim objString As String
Dim objParts() As String
Dim iNumObjToMod As Integer
Dim jLineBuffer As Integer
Dim found As Boolean

InFN = FreeFile()
Open originalFileName For Input As InFN
If mode = modeRead Then
    Debug.Print "Read IDF"
    wrt = False
Else
    Debug.Print "Write IDF"
    OutFN = FreeFile()
    Open revisedFileName For Output As OutFN
    wrt = True
    'establish the buffer for saving unchanged lines into output file
    numLineBuffer = 0
    sizeLineBuffer = 500
    ReDim lineBuffer(sizeLineBuffer)
End If
lineCount = 0
inObject = False
found = False

numIdfObjectStrings = 0
sizeIdfObjectStrings = 500
ReDim idfObjectStrings(sizeIdfObjectStrings)
Do While Not EOF(InFN)
    Line Input #InFN, DataLine ' read in data 1 line at a time
    If wrt Then
        numLineBuffer = numLineBuffer + 1
        If numLineBuffer > sizeLineBuffer Then
            sizeLineBuffer = sizeLineBuffer + 500
            ReDim Preserve lineBuffer(sizeLineBuffer)
        End If
        lineBuffer(numLineBuffer) = DataLine
    End If
    lineCount = lineCount + 1
    trimDataLine = Trim(DataLine)
    exclaimPos = InStr(trimDataLine, "!")
    commaPos = InStr(trimDataLine, ",")
    semiPos = InStr(trimDataLine, ";")
    If exclaimPos > 0 Then
        lineNoComment = Left(trimDataLine, exclaimPos - 1)
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
            objParts = Split(objString, ",")
            found = False
            For iNumObjToMod = 1 To numObjToMod
                If UCase(objParts(0)) = objToMod(iNumObjToMod).nameUC Then
                    found = True
                    numIdfObjectStrings = numIdfObjectStrings + 1
                    If numIdfObjectStrings > sizeIdfObjectStrings Then
                        sizeIdfObjectStrings = sizeIdfObjectStrings + 500
                        ReDim Preserve idfObjectStrings(sizeIdfObjectStrings)
                    End If
                    idfObjectStrings(numIdfObjectStrings) = objString
                End If
            Next iNumObjToMod
        End If
        objString = ""
        If wrt Then
            If found Then
                Call writeSubstitutedBuffer(OutFN, numIdfObjectStrings, lineBuffer)
            Else
                For jLineBuffer = 1 To numLineBuffer
                    Print #OutFN, lineBuffer(jLineBuffer)
                Next jLineBuffer
            End If
            numLineBuffer = 0
            found = False
        End If
    End If
Loop 'not eof(infn)
Debug.Print "number of IDF lines: " & lineCount
Close InFN
Close OutFN
End Sub

Sub ReadIDDforFields()
Dim iddFN As Integer
Dim fileLine As String
Dim lineCount As Long
Dim iObj As Long
Dim withinObject As Boolean
Dim posSlash As Long
Dim posSemi As Long
Dim fieldNm As String
Dim posField As Long
Dim numObjFound As Long
Dim objFound As Long

Dim curSI As String
Dim curIP As String
Dim curMult As String
numUnits = 0
sizeUnits = 150
ReDim units(sizeUnits)

numIddFields = 0
sizeIddFields = 2000
ReDim IDDfield(sizeIddFields)

numObjFound = 0
iddFN = FreeFile()
Open iddFileName For Input As iddFN
' read the unit conversions as the top of the IDD fiel
Do While Not EOF(iddFN)
    Line Input #iddFN, fileLine ' read in data 1 line at a time
    lineCount = lineCount + 1
    If Left(fileLine, 1) = "!" Then
        If Mid(fileLine, 31, 2) = "=>" Then
            'Debug.Print fileLine
            curSI = Trim(Mid(fileLine, 2, 28))
            curIP = Trim(Mid(fileLine, 33, 20))
            curMult = Trim(Mid(fileLine, 55))
            numUnits = numUnits + 1
            units(numUnits).siName = curSI
            units(numUnits).ipName = curIP
            If curSI = "C" And curIP = "F" Then
                units(numUnits).multiplier = 1.8
                units(numUnits).offset = 32
            Else
                units(numUnits).multiplier = Val(curMult)
                units(numUnits).offset = 0
            End If
            'Debug.Print curSI, curIP, curMult
        End If
    Else
        Exit Do
    End If
Loop 'not eof(iddfn)
Debug.Print "Number of units: ", numUnits

' read the objects and fields
withinObject = False
Do While Not EOF(iddFN)
    Line Input #iddFN, fileLine ' read in data 1 line at a time
    lineCount = lineCount + 1
    If Not withinObject Then
        For iObj = 1 To numObjToMod
            If UCase(fileLine) = objToMod(iObj).nameUC + "," Then
                withinObject = True
                objFound = iObj
                objToMod(iObj).name = Left(fileLine, Len(fileLine) - 1)
                Debug.Print "found: "; objToMod(iObj).name; " at: "; lineCount
            End If
        Next iObj
    Else 'within the object
        posField = InStr(fileLine, "\field")
        If posField > 0 Then
            curSI = ""
            curIP = ""
            fieldNm = Mid(fileLine, posField + 7)
            numIddFields = numIddFields + 1
            IDDfield(numIddFields).name = fieldNm
            If objToMod(objFound).firstField = 0 Then objToMod(objFound).firstField = numIddFields
            objToMod(objFound).lastField = numIddFields
        End If
        posField = InStr(fileLine, "\units")
        If posField > 0 Then
            curSI = Trim(Mid(fileLine, posField + 7))
            IDDfield(numIddFields).unitIndex = getUnits(curSI, curIP)
            'Debug.Print "\units", curSI, curIP, IDDfield(numIddFields).unitIndex
        End If
        posField = InStr(fileLine, "\ip-units")
        If posField > 0 Then
            curIP = Trim(Mid(fileLine, posField + 10))
            IDDfield(numIddFields).unitIndex = getUnits(curSI, curIP)
            'Debug.Print "\ip-units", curSI, curIP, IDDfield(numIddFields).unitIndex
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

Function getUnits(SIstring, IPstring) As Long
Dim found As Long
Dim i As Long
found = 0
If SIstring <> "" And IPstring <> "" Then
    For i = 1 To numUnits
        If SIstring = units(i).siName And IPstring = units(i).ipName Then
            found = i
            Exit For
        End If
    Next i
ElseIf SIstring <> "" Then
    For i = 1 To numUnits
        If SIstring = units(i).siName Then
            found = i
            Exit For
        End If
    Next i
Else
    found = 0
End If
getUnits = found
End Function

Sub putObjectsOnTab()
Dim iObj As Long
Dim jField As Long
Dim kObjToMod As Long
Dim pieces() As String
Dim nRow As Long
Dim nCol As Long
Dim maxRowsForObj As Long
Dim fieldStart As Long
Dim numOfFieldsInObj As Long
Dim formulaRowOffset As Long
Dim simpleCopyFormula As String
Dim origCellFormula As String

Application.ScreenUpdating = False
nRow = 10
For kObjToMod = 1 To numObjToMod
    maxRowsForObj = 0
    nCol = 4
    ' write the ORIGINAL rows
    Cells(nRow + 1, 2).Value = objToMod(kObjToMod).name + " [ORIGINAL]"
    For iObj = 1 To numIdfObjectStrings
        pieces = Split(idfObjectStrings(iObj), ",")
        If UCase(pieces(0)) = objToMod(kObjToMod).nameUC Then
            nCol = nCol + 1
            For jField = 1 To UBound(pieces)
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
        Cells(nRow + jField + 1, 3).Value = IDDfield(fieldStart + jField).name
        Cells(nRow + jField + 1, 4).Value = units(IDDfield(fieldStart + jField).unitIndex).siName
    Next jField
    formulaRowOffset = maxRowsForObj + 3
    simpleCopyFormula = "=R[" + Trim(Str(-formulaRowOffset)) + "]C"
    nRow = nRow + formulaRowOffset
    ' now write the REVISED rows
    Cells(nRow + 1, 2).Value = objToMod(kObjToMod).name + " [REVISED]"
    numOfFieldsInObj = 1 + objToMod(kObjToMod).lastField - objToMod(kObjToMod).firstField
    'put in the field names - go all the way to the end of the object
    For jField = 0 To (numOfFieldsInObj - 1)
        Cells(nRow + jField + 1, 3).Value = IDDfield(fieldStart + jField).name
        Cells(nRow + jField + 1, 4).Value = units(IDDfield(fieldStart + jField).unitIndex).siName
    Next jField
    ' insert formulas
    nCol = 4
    objToMod(kObjToMod).revisedCol = nCol + 1
    objToMod(kObjToMod).revisedRow = nRow + 1
    For iObj = 1 To numIdfObjectStrings
        pieces = Split(idfObjectStrings(iObj), ",")
        If UCase(pieces(0)) = objToMod(kObjToMod).nameUC Then
            nCol = nCol + 1
            For jField = 1 To UBound(pieces)
                origCellFormula = Cells(nRow + jField, nCol).FormulaR1C1
                Debug.Print origCellFormula, simpleCopyFormula
                Cells(nRow + jField, nCol).Interior.ColorIndex = 0
                If origCellFormula = "" Or origCellFormula = simpleCopyFormula Then
                    Cells(nRow + jField, nCol).Formula = simpleCopyFormula
                Else
                    Cells(nRow + jField, nCol).Interior.Color = RGB(240, 240, 240)
                End If
            Next jField
        End If
    Next iObj
    nRow = nRow + numOfFieldsInObj + 5
Next kObjToMod
Application.ScreenUpdating = True
End Sub

Sub writeSubstitutedBuffer(OutFN, objStrNum, lineBuffer)
Dim pieces() As String
Dim jField As Long
Dim kObjToMod As Long
Dim objFound As Long
Dim colIndex As Long
Dim nameRow As Long
Dim fieldStart As Long
Dim numOfFieldsInObj As Long
Dim commaOrSemi As String
Dim lastNonBlank As Long
Dim lastFieldOut As Long
Dim valueOfCellString As String

pieces = Split(idfObjectStrings(objStrNum), ",")
Print #OutFN,
' find the object
For kObjToMod = 1 To numObjToMod
    If UCase(pieces(0)) = objToMod(kObjToMod).nameUC Then
        objFound = kObjToMod
        Exit For
    End If
Next kObjToMod
' find the column with the matching name in the spreadsheet grid
colIndex = objToMod(objFound).revisedCol
nameRow = objToMod(objFound).revisedRow
Do
    If pieces(1) = Cells(nameRow, colIndex).Value Then Exit Do
    If Cells(nameRow, colIndex).Value = "" Then
        colIndex = -1
        Exit Do
    End If
    colIndex = colIndex + 1
Loop
If colIndex > 0 Then
    Print #OutFN, "  "; pieces(0); ","
    numOfFieldsInObj = 1 + objToMod(objFound).lastField - objToMod(objFound).firstField
    fieldStart = objToMod(kObjToMod).firstField
    'determine last field that is not blank
    lastNonBlank = 0
    For jField = (numOfFieldsInObj - 1) To 0 Step -1
        If Cells(nameRow + jField, colIndex).Value <> "" Then
            lastNonBlank = jField
            Exit For
        End If
    Next jField
    If lastNonBlank < (numOfFieldsInObj - 1) Then
        lastFieldOut = lastNonBlank
    Else
        lastFieldOut = (numOfFieldsInObj - 1)
    End If
    'write the object
    commaOrSemi = ","
    For jField = 0 To lastFieldOut
        If jField = lastFieldOut Then commaOrSemi = ";"
        valueOfCellString = Trim(Cells(nameRow + jField, colIndex).Value)
        If valueOfCellString = "0" And pieces(jField + 1) = "" Then valueOfCellString = ""
        Print #OutFN, "    "; valueOfCellString; commaOrSemi; Tab(30); "!- "; IDDfield(fieldStart + jField).name;
        If IDDfield(fieldStart + jField).unitIndex <> 0 Then
            Print #OutFN, " {"; units(IDDfield(fieldStart + jField).unitIndex).siName; "}"
        Else
            Print #OutFN, ""
        End If
    Next jField
End If
End Sub
