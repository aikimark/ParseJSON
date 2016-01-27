Attribute VB_Name = "modTests"
Option Explicit

Sub DataSave(parmObject As Object)
    '* * * Can not save objects with a Put statement.
    'ToDo: use some form of the Iterate code to save parmObject
    Dim intFN As Integer
    Dim vKeys As Variant
    Dim vData As Variant
    Dim vItem As Variant
    
    intFN = FreeFile
    Open "C:\Users\Mark\Downloads\Small-REST-Output.bin" For Binary As #intFN
    Select Case TypeName(parmObject)
        Case "Collection"
            For Each vItem In parmObject
            Next
        Case "Dictionary"
            vKeys = parmObject.keys
            vData = parmObject.items
            Put #intFN, , vKeys
            Put #intFN, , vData    'fails if vData is an object (dictionary or collection)
        Case Else
        
    End Select
    Close #intFN
End Sub

Sub testParseJSON()
    Dim oThing As Object
    'Set oThing = parseJSON("C:\Users\Mark\Downloads\Small-REST-Output.txt")
    'Set oThing = parseJSON("C:\Users\Mark\Downloads\Q_28906582.txt")
    Set oThing = parseJSONfile("C:\Users\Mark\Downloads\Q_28918483.JSON.txt")

'Note: DataSave does not currently serialize object data types
'    DataSave oThing(1)

    testIterateObject oThing, 0
End Sub

Function testIterateObject(parmObject As Object, parmDepth As Long)
    Dim vItem As Variant
    Dim oItem As Object
    Dim strDelim As String
    
    Select Case TypeName(parmObject)
        Case "Dictionary"
            For Each vItem In parmObject
                If VarType(parmObject(vItem)) = vbObject Then
                    Debug.Print String(parmDepth, vbTab); vItem, "Count: " & parmObject(vItem).Count
                    testIterateObject parmObject(vItem), parmDepth + 1
                Else
                    Select Case VarType(parmObject(vItem))
                        Case VbVarType.vbString
                            strDelim = """"
                        Case VbVarType.vbDate
                            strDelim = "#"
                        Case Else
                            strDelim = vbNullString
                    End Select
                    Debug.Print String(parmDepth, vbTab); vItem, strDelim & parmObject(vItem) & strDelim
                End If
            Next
        Case "Collection"
            For Each vItem In parmObject
                If VarType(vItem) = vbObject Then
                    'Debug.Print vItem
                    Set oItem = vItem
                    testIterateObject oItem, parmDepth + 1
                Else
                    Select Case VarType(vItem)
                        Case VbVarType.vbString
                            strDelim = """"
                        Case VbVarType.vbDate
                            strDelim = "#"
                        Case Else
                            strDelim = vbNullString
                    End Select
                    Debug.Print String(parmDepth, vbTab); strDelim & vItem & strDelim
                End If
            Next
    End Select
End Function

Sub testStack()
    Dim colThing As Collection
    Dim oThing As Object
    Dim vThing As Variant
    Dim clsM As clsMatch
    Dim vItem As Variant
    Dim lngLoop As Long
    Dim sngStart As Single
'    sngStart = Timer
    Set colThing = New Collection
    Set clsM = New clsMatch
    clsM.lngM = -999
    colThing.Add clsM
    vThing = Array(1, 2, 3)
    colThing.Add vThing
    Set oThing = CreateObject("scripting.dictionary")
    oThing!Mark = "self"
    colThing.Add oThing
    Set vThing = CreateObject("scripting.dictionary")
    vThing!Fred = "brother"
    colThing.Add vThing
    Set vThing = New Collection
    vThing.Add 123456
    vThing.Add 7890
    colThing.Add vThing
    Stop
'    For lngLoop = 1 To 1000
'        Set clsM = New clsMatch
'        clsM.lngM = lngLoop
'        If colThing.Count = 0 Then
'            colThing.Add clsM
'        Else
'            colThing.Add clsM, , 1
'        End If
'    Next
'    Debug.Print "Population time: " & Timer - sngStart
'    Set clsM = colThing(1)
'    clsM.lngM = -999
'    Debug.Print clsM.lngM
'    Set clsM = colThing(1)
'    Debug.Print clsM.lngM
    
'    Set clsM = New clsMatch
'    clsM.lngM = 2
'    colThing.Add clsM, , 1
'    Debug.Print colThing.Count
'    sngStart = Timer
'    For Each vitem In colThing
'        Set clsM = vitem
'        Debug.Print clsM.lngM, ;
'    Next
'    Debug.Print vbCrLf; "Print time: " & Timer - sngStart
'    sngStart = Timer
'    Do Until colThing.Count = 0
''        Set clsM = colThing(1)
''        Set clsM = Nothing
''        Set colThing(1) = Nothing
'        colThing.Remove 1
'    Loop
'    Debug.Print vbCrLf; "Clear collection time: " & Timer - sngStart
End Sub


Sub DataSave(parmObject As Object)
'To Do: correctly persist the parsed object data in a friendly format
'Solution paths:
'    * intrinsic VB I/O
'    * ADODB recordset and stream
'    * XML - seems like a cheat
End Sub