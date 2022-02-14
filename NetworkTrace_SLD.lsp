Option Explicit
Dim i As Integer
Public iTypeFieldIndex As Integer
Public iPillar_NumberFieldIndex As Integer
Public iCKT_NO As Integer
Public iSTATUS As Integer
Public i2_CKTNO As Integer
Public i4m_CKTNO As Integer
Public iPillar4m As Integer
Public iPillar2 As Integer
Dim PillarArray() As Long
Dim LT_Array() As Long
Dim Name_String() As String
Dim iName As Integer
Public iLT_array As Integer
Public iPillararray As Integer
Public SSet4LT_name As Variant
Dim ssetLineObj As AcadSelectionSet
Public Function Get_Entry_Pillar()
ZoomExtents 'Full extent the current drawing

iTypeFieldIndex = GetPillar_TypeFieldIndexes() 'get field index of type field in pillar object data
iPillar_NumberFieldIndex = GetPillar_PillarNo_FieldIndexes() 'get the pillar number index of pillartype field in pillar object data

Dim acadOBJ As Object
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
Dim boolVal As Boolean

Dim OD_PB As ODTable ' variable to connect to PB table
Set OD_PB = acad.Projects(ThisDrawing).ODTables.Item("PILLAR_BOUNDARY") ' setting up the table for use

Dim ODrcs_PB As ODRecords
Set ODrcs_PB = OD_PB.GetODRecords ' get all records in the table


Dim i10w As Integer
Dim i8w As Integer
Dim i6w As Integer
Dim i4w As Integer
Dim i2w As Integer
i10w = i8w = i6w = i4w = i2w = 0

Dim PillarType As String
PillarType = ""
'count the different types of pillars so as to decide entry point of the LT
For Each acadOBJ In ThisDrawing.ModelSpace
    boolVal = ODrcs_PB.Init(acadOBJ, True, False)
        Do While ODrcs_PB.IsDone = False
            If ODrcs_PB.Record.Item(iTypeFieldIndex).Value = "10W" Then
                i10w = i10w + 1
            End If
            If ODrcs_PB.Record.Item(iTypeFieldIndex).Value = "8W" Then
                i8w = i8w + 1
            End If
            If ODrcs_PB.Record.Item(iTypeFieldIndex).Value = "6W" Then
                i6w = i6w + 1
            End If
            If ODrcs_PB.Record.Item(iTypeFieldIndex).Value = "4W" Then
                i4w = i4w + 1
            End If
            If ODrcs_PB.Record.Item(iTypeFieldIndex).Value = "2W" Then
                i2w = i2w + 1
            End If
          ODrcs_PB.Next
        Loop
Next
If i10w > 0 Then
    PillarType = "10W"
    Call Get_8W_PillarAttributes(PillarType)
Else
    If i8w > 0 Then
        PillarType = "8W"
        Call Get_8W_PillarAttributes(PillarType)
    Else
        If i6w > 0 Then
            PillarType = "6W"
            Call Get_8W_PillarAttributes(PillarType)
        Else
            If i4w > 0 Then
                PillarType = "4W"
                Call Get_8W_PillarAttributes(PillarType)
            Else
                If i2w > 0 Then
                    PillarType = "2W"
                    Call Get_8W_PillarAttributes(PillarType)
                End If
            End If
        End If
    End If
End If
End Function
'Entry level function
Private Function Get_8W_PillarAttributes(PillarType As String)
 
ZoomExtents 'Full extent the current drawing
 
iPillararray = 0
iLT_array = 0
iName = 0
ReDim Preserve LT_Array(iLT_array) As Long
LT_Array(iLT_array) = 0
ReDim Preserve Name_String(iName) As String
Name_String(iName) = ""

iTypeFieldIndex = GetPillar_TypeFieldIndexes() 'get field index of type field in pillar object data
iPillar_NumberFieldIndex = GetPillar_PillarNo_FieldIndexes() 'get the pillar number index of pillartype field in pillar object data
iPillar4m = GetLT_Pillar4m_Index()
iPillar2 = GetLT_Pillar2_Index()
i4m_CKTNO = GetLT_4m_CKTNO_Index()
i2_CKTNO = GetLT_2_CKTNO_Index()

Dim acadOBJ As Object
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
Dim boolVal As Boolean

Dim OD_PB As ODTable ' variable to connect to PB table
Set OD_PB = acad.Projects(ThisDrawing).ODTables.Item("PILLAR_BOUNDARY") ' setting up the table for use

Dim ODrcs_PB As ODRecords
'Set ODrcs_PB = OD_PB.GetODRecords ' get all records in the table

Dim pillar_8W_objectID() As Long 'store pillar objectID in an array
Dim iPillarCount As Integer
Dim pillarno() As Long
Dim iPillarNo As Integer 'store pillarnumber in this array

For Each acadOBJ In ThisDrawing.ModelSpace
    If Left$(acadOBJ.Layer, 6) = "PILLAR" Then
        Set ODrcs_PB = OD_PB.GetODRecords
        ODrcs_PB.Init acadOBJ, False, False
        If ODrcs_PB.IsDone = True Then
            boolVal = True
        End If
    Do Until ODrcs_PB.IsDone 'This runs for number of circuit breakers in the given pillar
        Set ODrcs_PB = ODrc_PB.Record
            If ODrcs_PB.Record.Item(iTypeFieldIndex).Value = PillarType Then
                MsgBox "Sucess"
                iPillarCount = iPillarCount + 1
                iPillarNo = iPillarNo + 1
                ReDim Preserve pillar_8W_objectID(iPillarCount) As Long
                ReDim Preserve pillarno(iPillarNo) As Long
                pillar_8W_objectID(iPillarCount - 1) = ODrcs_PB.Record.ObjectID
                pillarno(iPillarNo - 1) = ODrcs_PB.Record.Item(iPillar_NumberFieldIndex).Value
            End If
        ODrcs_PB.Next
    Loop
'    Exit Function
Next
Call GetObjects_in_Pillars(pillar_8W_objectID(), pillarno()) 'passing pillar objID and pillarno
End Function
Private Function GetPillar_TypeFieldIndexes()

Dim ODtbs As ODTables
Dim iType As Integer

Set ODtbs = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application").Projects(ThisDrawing).ODTables
'get the field names

For i = 0 To ODtbs.Item("PILLAR_BOUNDARY").ODFieldDefs.Count - 1
    'get the field index of "TYPE" field
        If ODtbs.Item("PILLAR_BOUNDARY").ODFieldDefs.Item(i).Name = "TYPE" Then
            iType = i
        End If
Next i

GetPillar_TypeFieldIndexes = iType

End Function
Private Function GetPillar_PillarNo_FieldIndexes()

Dim ODtbs As ODTables
Dim iPillarNo As Integer

Set ODtbs = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application").Projects(ThisDrawing).ODTables

'get the field names
For i = 0 To ODtbs.Item("PILLAR_BOUNDARY").ODFieldDefs.Count - 1
    'get the field index of "PILLAR_NO" field
        If ODtbs.Item("PILLAR_BOUNDARY").ODFieldDefs.Item(i).Name = "PILLAR_NO" Then
            iPillarNo = i
        End If
Next i

GetPillar_PillarNo_FieldIndexes = iPillarNo

End Function
Private Function GetObjects_in_Pillars(pillar_8W_objectID() As Long, pillarno() As Long)

Dim minExt As Variant
Dim maxExt As Variant
Dim ssetObj As AcadSelectionSet
Dim iNo_of_Pillars As Integer

Dim mode As Integer
mode = acSelectionSetCrossing

If ThisDrawing.SelectionSets.Count > 0 Then
    For i = (ThisDrawing.SelectionSets.Count - 1) To 0 Step -1
        ThisDrawing.SelectionSets.Item(i).Delete
    Next i
End If

If ThisDrawing.SelectionSets.Count = 0 Then
    For iNo_of_Pillars = 0 To (UBound(pillar_8W_objectID()) - 1) Step 1
        Dim acadLWPOLY As AcadLWPolyline
        Set acadLWPOLY = Nothing
        Set acadLWPOLY = ThisDrawing.ObjectIdToObject(pillar_8W_objectID(iNo_of_Pillars))
        acadLWPOLY.GetBoundingBox minExt, maxExt
        'declaration of selesection set
        Set ssetObj = Nothing
        MsgBox "Pillar Number: " & pillarno(iNo_of_Pillars)
        'Create a filter for getting only CIRCUIT BLOCKS
        Set ssetObj = ThisDrawing.SelectionSets.Add(pillar_8W_objectID(iNo_of_Pillars))
        Dim gpCode(0) As Integer
        Dim dataValue(0) As Variant
        gpCode(0) = 8
        Dim iDataValue As Integer
        iDataValue = 0
        
        For iDataValue = 0 To (ThisDrawing.Layers.Count - 1) Step 1
            If Left$(ThisDrawing.Layers.Item(iDataValue).Name, 3) = "CIR" Then
                dataValue(0) = ThisDrawing.Layers.Item(iDataValue).Name
            End If
        Next iDataValue
        
        Dim groupCode As Variant, dataCode As Variant
        groupCode = gpCode
        dataCode = dataValue
        ssetObj.Select mode, maxExt, minExt, groupCode, dataCode
        Call GetCBObjects(ssetObj, pillarno(iNo_of_Pillars))
    Next
End If

End Function
Private Sub GetCBObjects(ssetObj As AcadSelectionSet, pillarno As Long)

Dim acadOBJ As Object

Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")

Dim boolVal As Boolean

Dim ODT As ODTables
Set ODT = acad.Projects(ThisDrawing).ODTables

Dim OD_CB As ODTable
Set OD_CB = ODT.Item("CIRCUIT_BNDY")

Dim minExt As Variant
Dim maxExt As Variant
Dim acblock As AcadBlockReference

Dim ODrcs_CB As ODRecords
Dim ODrc_CB As ODRecord
Set ODrcs_CB = Nothing

iCKT_NO = GetCB_CKTNO_Index()
iSTATUS = GetCB_Status_Index()

Dim iCKT45 As Integer, iCKT4 As Integer, iCKT5 As Integer

For Each acadOBJ In ssetObj
    If Left$(acadOBJ.Layer, 3) = "CIR" Then
        Set ODrcs_CB = OD_CB.GetODRecords
        ODrcs_CB.Init acadOBJ, False, False
        If ODrcs_CB.IsDone = True Then
            boolVal = True
        End If
        
        Do Until ODrcs_CB.IsDone 'This runs for number of circuit breakers in the given pillar
            Set ODrc_CB = ODrcs_CB.Record
            If ODrc_CB.Item(iCKT_NO).Value = "CKT4" And ODrc_CB.Item(iSTATUS).Value = "CLOSE" Then
                iCKT45 = iCKT45 + 1
                iCKT4 = iCKT4 + 1
            End If
            If ODrc_CB.Item(iCKT_NO).Value = "CKT5" And ODrc_CB.Item(iSTATUS).Value = "CLOSE" Then
                iCKT45 = iCKT45 + 1
                iCKT5 = iCKT5 + 1
            End If
            ODrcs_CB.Next
        Loop
    
    End If
Next

'Check which side of pillar is feeded
If iCKT45 = 2 Then
    MsgBox "Entry Level Pillar: " & pillarno
    Call For_Each_CB_EntryPillar(ssetObj, pillarno)
Else:
    MsgBox "More than one entry level pillars"
        If iCKT4 = 1 And iCKT5 = 0 Then
            MsgBox "Left side of entry level pillar is feeded by DT"
        Else: MsgBox "Right side of entry level pillar is feeded by DT"
        End If
End If

End Sub
Private Function GetCB_CKTNO_Index()

Dim ODtbs As ODTables
Dim iCKT As Integer
Set ODtbs = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application").Projects(ThisDrawing).ODTables

'get the field names
For i = 0 To ODtbs.Item("CIRCUIT_BNDY").ODFieldDefs.Count - 1
        If ODtbs.Item("CIRCUIT_BNDY").ODFieldDefs.Item(i).Name = "CKT_NO" Then
           iCKT = i
        End If
Next i

GetCB_CKTNO_Index = iCKT

End Function
Private Function GetCB_Status_Index()

Dim ODtbs As ODTables
Dim iSTAT As Integer

Set ODtbs = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application").Projects(ThisDrawing).ODTables

'get the field names
For i = 0 To ODtbs.Item("CIRCUIT_BNDY").ODFieldDefs.Count - 1
        If ODtbs.Item("CIRCUIT_BNDY").ODFieldDefs.Item(i).Name = "STATUS" Then
           iSTAT = i
        End If
Next i

GetCB_Status_Index = iSTAT

End Function
Private Function For_Each_CB_EntryPillar(ssetObj As AcadSelectionSet, pillarno As Long)

Dim acadOBJ As Object
Dim acad As AcadMap

Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")

Dim boolVal As Boolean

Dim ODT As ODTables
Set ODT = acad.Projects(ThisDrawing).ODTables

Dim OD_CB As ODTable
Set OD_CB = ODT.Item("CIRCUIT_BNDY")

Dim minExt As Variant
Dim maxExt As Variant
Dim acblock As AcadBlockReference
Dim ODrcs_CB As ODRecords
Dim ODrc_CB As ODRecord

Dim mode As Integer
mode = acSelectionSetCrossing

Set ODrcs_CB = Nothing

'MsgBox "Number of CB: " & ssetObj.Count & " in" & PillarNo

For Each acadOBJ In ssetObj
    If Left$(acadOBJ.Layer, 3) = "CIR" Then
        Set ODrcs_CB = OD_CB.GetODRecords
        ODrcs_CB.Init acadOBJ, False, False
        If ODrcs_CB.IsDone = True Then
            boolVal = True
        End If
        Do Until ODrcs_CB.IsDone 'This runs for number of circuit breakers in the given pillar
        Set ODrc_CB = ODrcs_CB.Record
            If ODrc_CB.Item(iSTATUS).Value = "OPEN" Then
                GoTo nextCB
            Else:
            Set acblock = Nothing
            Set acblock = ThisDrawing.ObjectIdToObject(ODrcs_CB.Record.ObjectID)
            acblock.GetBoundingBox minExt, maxExt
'            MsgBox PillarNo & " ." & ODrc_CB.Item(iCKT_NO).Value
            Call LT_4_Each_CB(acblock, minExt, maxExt, pillarno, ODrc_CB.Item(iCKT_NO).Value)
            GoTo nextCB
            End If
nextCB:
            ODrcs_CB.Next
        Loop
    End If
Next

End Function
Private Function GetLT_Pillar4m_Index()
Dim ODtbs As ODTables
Dim iPillar4m As Integer

Set ODtbs = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application").Projects(ThisDrawing).ODTables

'get the field names
For i = 0 To ODtbs.Item("LT_CABLE").ODFieldDefs.Count - 1
        If ODtbs.Item("LT_CABLE").ODFieldDefs.Item(i).Name = "FROM_PILLARNO" Then
           iPillar4m = i
        End If
Next i

GetLT_Pillar4m_Index = iPillar4m

End Function
Private Function GetLT_Pillar2_Index()
Dim ODtbs As ODTables
Dim iPillar2 As Integer

Set ODtbs = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application").Projects(ThisDrawing).ODTables

'get the field names
For i = 0 To ODtbs.Item("LT_CABLE").ODFieldDefs.Count - 1
        If ODtbs.Item("LT_CABLE").ODFieldDefs.Item(i).Name = "TO_PILLARNO" Then
           iPillar2 = i
        End If
Next i

GetLT_Pillar2_Index = iPillar2

End Function
Private Function GetLT_4m_CKTNO_Index()
Dim ODtbs As ODTables
Dim i4CKTNO As Integer

Set ODtbs = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application").Projects(ThisDrawing).ODTables

'get the field names
For i = 0 To ODtbs.Item("LT_CABLE").ODFieldDefs.Count - 1
        If ODtbs.Item("LT_CABLE").ODFieldDefs.Item(i).Name = "FROM_CKTNO" Then
           i4CKTNO = i
        End If
Next i

GetLT_4m_CKTNO_Index = i4CKTNO

End Function
Private Function GetLT_2_CKTNO_Index()

Dim ODtbs As ODTables
Dim i2CKTNO As Integer

Set ODtbs = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application").Projects(ThisDrawing).ODTables

'get the field names
For i = 0 To ODtbs.Item("LT_CABLE").ODFieldDefs.Count - 1
        If ODtbs.Item("LT_CABLE").ODFieldDefs.Item(i).Name = "TO_CKTNO" Then
           i2CKTNO = i
        End If
Next i

GetLT_2_CKTNO_Index = i2CKTNO

End Function
Private Function LT_4_Each_CB(acblock As AcadBlockReference, minExt As Variant, maxExt As Variant, pillarno As Long, CB_CKT_No As Variant)
Dim acadOBJ As Object

Dim mode As Integer
mode = acSelectionSetCrossing

Dim SSet4LT As AcadSelectionSet

'MsgBox PillarNo & "." & CB_CKT_No
Set SSet4LT = ThisDrawing.SelectionSets.Add(pillarno & "." & CB_CKT_No)

'Create a filter for getting only LT_Cable
Dim gpCode(0) As Integer
gpCode(0) = 8

Dim dataValue(0) As Variant

Dim iDataValue As Integer
iDataValue = 0

For iDataValue = 0 To (ThisDrawing.Layers.Count - 1) Step 1
    If Left$(ThisDrawing.Layers.Item(iDataValue).Name, 2) = "LT" Then
        dataValue(0) = ThisDrawing.Layers.Item(iDataValue).Name
    End If
Next iDataValue

Dim groupCode As Variant, dataCode As Variant
groupCode = gpCode
dataCode = dataValue

SSet4LT.Select mode, maxExt, minExt, groupCode, dataCode
'MsgBox "Number of LT: " & PillarNo & "." & CB_CKT_No & "=" & SSet4LT.Count

'for each CB
'MsgBox SSet4LT.Count
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")

Dim boolVal As Boolean

Dim ODT As ODTables
Set ODT = acad.Projects(ThisDrawing).ODTables

Dim OD_LT As ODTable
Set OD_LT = ODT.Item("LT_CABLE")

Dim ODrcs_LT As ODRecords
Dim ODrc_LT As ODRecord
Dim iRecord As Integer
iRecord = 0

For Each acadOBJ In SSet4LT
    If Left$(acadOBJ.Layer, 2) = "LT" Then
        Set ODrcs_LT = OD_LT.GetODRecords
        ODrcs_LT.Init acadOBJ, False, False
        
        If ODrcs_LT.IsDone = True Then
            boolVal = True
        End If
        
         Do Until ODrcs_LT.IsDone 'This runs for number of circuit breakers in the given pillar
            Set ODrc_LT = ODrcs_LT.Record
            ReDim Preserve LT_Array(iLT_array) As Long
            For i = 0 To UBound(LT_Array) Step 1
                If ODrc_LT.ObjectID = LT_Array(i) Then
'                    MsgBox "Same LT_Cable:" & ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & ODrc_LT.Item(iPillar2).Value & "." & ODrc_LT.Item(i2_CKTNO).Value
                    GoTo nextrecord
                End If
            Next i
            iLT_array = iLT_array + 1
            ReDim Preserve LT_Array(iLT_array) As Long
            LT_Array(iLT_array) = ODrc_LT.ObjectID
'            Add this LT_Cable to the existing array
                If ODrc_LT.Item(i2_CKTNO).Value <> "MP" Or ODrc_LT.Item(i4m_CKTNO).Value <> "MP" Then
                    If ODrc_LT.Item(iPillar2).Value = "SERVICE" Then
                        MsgBox ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & "SERVICE POINT"
                        GoTo nextrecord
                    Else:
                        MsgBox ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & ODrc_LT.Item(iPillar2).Value & "." & ODrc_LT.Item(i2_CKTNO).Value
                        Call To_Pillar(ODrc_LT.Item(iPillar2).Value, ODrc_LT.Item(i2_CKTNO).Value, LT_Array())
                        GoTo nextrecord
                    End If
                
'                    Call To_MP
'                    Exit Function
                End If
           
nextrecord:
            ODrcs_LT.Next
        Loop
    End If
Next
End Function
Private Function To_Pillar(pillarno As Long, CKT_No As Variant, LT_Array() As Long)

If CKT_No = "MP" Then
    GoTo MP_Trace
    Exit Function
End If
'MsgBox "PillarNo: " & PillarNo & " " & "CKT Number: " & CKT_No
Dim acadOBJ As Object
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
Dim boolVal As Boolean
Dim boolBC As Boolean
Dim OD_PB As ODTable
Set OD_PB = acad.Projects(ThisDrawing).ODTables.Item("PILLAR_BOUNDARY")
Dim ODrcs_PB As ODRecords
Dim ODrc_PB As ODRecord
Dim iPillar As Long
Dim Pillar_ObjectID As Long
For Each acadOBJ In ThisDrawing.ModelSpace
    If Left$(acadOBJ.Layer, 6) = "PILLAR" Then
        Set ODrcs_PB = OD_PB.GetODRecords
        ODrcs_PB.Init acadOBJ, False, False
        If ODrcs_PB.IsDone = True Then
            boolVal = True
        End If
        Do Until ODrcs_PB.IsDone
          Set ODrc_PB = ODrcs_PB.Record
          iPillar = ODrc_PB.Item(iPillar_NumberFieldIndex).Value
            If iPillar = pillarno Then
                Pillar_ObjectID = ODrc_PB.ObjectID
                If ODrc_PB.Item(iTypeFieldIndex).Value = "6W" Or ODrc_PB.Item(iTypeFieldIndex).Value = "4W" Or ODrc_PB.Item(iTypeFieldIndex).Value = "2W" Then
'                    MsgBox "The pillar is " & ODrc_PB.Item(iTypeFieldIndex).Value & " TYPE"
                    'create selection set for this pillar
                    Dim ssetpillar As AcadSelectionSet
                    Dim acadLWPOLY As AcadLWPolyline
                    Dim minExt As Variant
                    Dim maxExt As Variant
                    Set acadLWPOLY = Nothing
                    Set acadLWPOLY = ThisDrawing.ObjectIdToObject(Pillar_ObjectID)
                    acadLWPOLY.GetBoundingBox minExt, maxExt
                    'declaration of selesection set
                    ReDim Preserve PillarArray(iPillararray) As Long
                    PillarArray(iPillararray) = Pillar_ObjectID
                    'Create a filter for getting only CIRCUIT BLOCKS
                    Dim mode As Integer
                    mode = acSelectionSetCrossing
'                    MsgBox PillarArray(iPillararray) & "." & (iPillararray + 1)
                    Set ssetpillar = Nothing
                    Set ssetpillar = ThisDrawing.SelectionSets.Add(PillarArray(iPillararray) & "." & (iPillararray + 1))
                    Dim gpCode(0) As Integer
                    Dim dataValue(0) As Variant
                    gpCode(0) = 8
                    Dim iDataValue As Integer
                    iDataValue = 0
                        For iDataValue = 0 To (ThisDrawing.Layers.Count - 1) Step 1
                            If Left$(ThisDrawing.Layers.Item(iDataValue).Name, 3) = "CIR" Then
                                dataValue(0) = ThisDrawing.Layers.Item(iDataValue).Name
                            End If
                        Next iDataValue
                    Dim groupCode As Variant, dataCode As Variant
                    groupCode = gpCode
                    dataCode = dataValue
                    ssetpillar.Select mode, maxExt, minExt, groupCode, dataCode
                    iPillararray = iPillararray + 1
                    GoTo Record
                Else:
'                    MsgBox "Inside MP: Pillar"
MP_Trace:
                    Call Inside_MP(pillarno)
                    Exit Function
                End If
            End If
        ODrcs_PB.Next
'        Exit Function
        Loop
    End If
Next
Record:
Call Get_CB_in_Pillar(ssetpillar, pillarno, LT_Array())
Call Get_BB_LT(pillarno)
End Function
Private Function Get_CB_in_Pillar(ssetpillar As AcadSelectionSet, pillarno As Long, LT_Array() As Long)
Dim acadOBJ As Object
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
Dim boolVal As Boolean
Dim ODT As ODTables
Set ODT = acad.Projects(ThisDrawing).ODTables
Dim OD_CB As ODTable
Set OD_CB = ODT.Item("CIRCUIT_BNDY")
Dim minExt As Variant
Dim maxExt As Variant
Dim acblock As AcadBlockReference
Dim ODrcs_CB As ODRecords
Dim ODrc_CB As ODRecord
Dim mode As Integer
mode = acSelectionSetCrossing
Dim sset_LT As AcadSelectionSet

Set ODrcs_CB = Nothing
'MsgBox "Number of CB: " & ssetpillar.Count & " in" & pillarno
For Each acadOBJ In ssetpillar
    If Left$(acadOBJ.Layer, 3) = "CIR" Then
        Set ODrcs_CB = OD_CB.GetODRecords
        ODrcs_CB.Init acadOBJ, False, False
        If ODrcs_CB.IsDone = True Then
            boolVal = True
        End If
        Do Until ODrcs_CB.IsDone 'This runs for number of circuit breakers in the given pillar
            Set ODrc_CB = ODrcs_CB.Record
            Dim Name As Variant
            Dim gpCode(0) As Integer
            Dim dataValue(0) As Variant
            Dim groupCode As Variant, dataCode As Variant
            If ODrc_CB.Item(iSTATUS).Value = "OPEN" Then
                Set acblock = Nothing
                Set acblock = ThisDrawing.ObjectIdToObject(ODrc_CB.ObjectID)
                acblock.GetBoundingBox minExt, maxExt
                Name = ODrc_CB.ObjectID & "." & ODrc_CB.Item(iCKT_NO).Value
                'create selection set for each Circuit Breaker
                ReDim Preserve Name_String(iName) As String
                For i = 0 To UBound(Name_String) Step 1
                    If Name = Name_String(i) Then
                        GoTo nextCB_pillar
                    End If
                Next i
                Name_String(iName) = Name
                Set sset_LT = ThisDrawing.SelectionSets.Add(Name_String(iName))
                iName = iName + 1
                'Create a filter for getting only LT_Cable
                gpCode(0) = 8
                Dim iDataValue As Integer
                iDataValue = 0
                For iDataValue = 0 To (ThisDrawing.Layers.Count - 1) Step 1
                    If Left$(ThisDrawing.Layers.Item(iDataValue).Name, 2) = "LT" Then
                        dataValue(0) = ThisDrawing.Layers.Item(iDataValue).Name
                    End If
                Next iDataValue
                groupCode = gpCode
                dataCode = dataValue
                sset_LT.Select mode, maxExt, minExt, groupCode, dataCode
                If sset_LT.Count > 1 Then
                    Call OPEN_CB(sset_LT)
                End If
                GoTo nextCB_pillar
            Else:
                Set acblock = Nothing
                Set acblock = ThisDrawing.ObjectIdToObject(ODrc_CB.ObjectID)
                acblock.GetBoundingBox minExt, maxExt
                Name = ODrc_CB.ObjectID & "." & ODrc_CB.Item(iCKT_NO).Value
                'create selection set for each Circuit Breaker
                ReDim Preserve Name_String(iName) As String
                For i = 0 To UBound(Name_String) Step 1
                    If Name = Name_String(i) Then
                        GoTo nextCB_pillar
                    End If
                Next i
                Name_String(iName) = Name
                Set sset_LT = ThisDrawing.SelectionSets.Add(Name_String(iName))
                iName = iName + 1
                'Create a filter for getting only LT_Cable
                gpCode(0) = 8
                iDataValue = 0
                For iDataValue = 0 To (ThisDrawing.Layers.Count - 1) Step 1
                    If Left$(ThisDrawing.Layers.Item(iDataValue).Name, 2) = "LT" Then
                        dataValue(0) = ThisDrawing.Layers.Item(iDataValue).Name
                    End If
                Next iDataValue
                groupCode = gpCode
                dataCode = dataValue
                sset_LT.Select mode, maxExt, minExt, groupCode, dataCode
                Call LT_4_Each_CB(acblock, minExt, maxExt, pillarno, ODrc_CB.Item(iCKT_NO).Value)
'                End If
                GoTo nextCB_pillar
            End If
nextCB_pillar:
             ODrcs_CB.Next
        Loop
    End If
 Next
End Function
Private Function Inside_MP(pillarno As Long)
Dim acadOBJ As Object
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
Dim boolVal As Boolean
Dim boolBC As Boolean
Dim OD_PB As ODTable
Set OD_PB = acad.Projects(ThisDrawing).ODTables.Item("PILLAR_BOUNDARY")
Dim ODrcs_PB As ODRecords
Dim ODrc_PB As ODRecord
Dim iPillar As Long
Dim Pillar_ObjectID As Long
For Each acadOBJ In ThisDrawing.ModelSpace
    If Left$(acadOBJ.Layer, 6) = "PILLAR" Then
        Set ODrcs_PB = OD_PB.GetODRecords
        ODrcs_PB.Init acadOBJ, False, False
        If ODrcs_PB.IsDone = True Then
            boolVal = True
        End If
        Do Until ODrcs_PB.IsDone
          Set ODrc_PB = ODrcs_PB.Record
          iPillar = ODrc_PB.Item(iPillar_NumberFieldIndex).Value
            If iPillar = pillarno Then
                Pillar_ObjectID = ODrc_PB.ObjectID
                'create selection set for this pillar
                Dim ssetpillar As AcadSelectionSet
                Dim acadLWPOLY As AcadLWPolyline
                Dim minExt As Variant
                Dim maxExt As Variant
                Set acadLWPOLY = Nothing
                Set acadLWPOLY = ThisDrawing.ObjectIdToObject(Pillar_ObjectID)
                acadLWPOLY.GetBoundingBox minExt, maxExt
                'declaration of selesection set
                ReDim Preserve PillarArray(iPillararray) As Long
                PillarArray(iPillararray) = Pillar_ObjectID
                'Create a filter for getting only CIRCUIT BLOCKS
                Dim mode As Integer
                mode = acSelectionSetCrossing
'                MsgBox PillarArray(iPillararray) & "." & (iPillararray + 1)
                Set ssetpillar = Nothing
                Set ssetpillar = ThisDrawing.SelectionSets.Add(PillarArray(iPillararray) & "." & (iPillararray + 1))
                Dim gpCode(0) As Integer
                Dim dataValue(0) As Variant
                gpCode(0) = 8
                Dim iDataValue As Integer
                iDataValue = 0
                    For iDataValue = 0 To (ThisDrawing.Layers.Count - 1) Step 1
                        If Left$(ThisDrawing.Layers.Item(iDataValue).Name, 2) = "LT" Then
                            dataValue(0) = ThisDrawing.Layers.Item(iDataValue).Name
                        End If
                    Next iDataValue
                Dim groupCode As Variant, dataCode As Variant
                groupCode = gpCode
                dataCode = dataValue
                ssetpillar.Select mode, maxExt, minExt, groupCode, dataCode
                If ssetpillar.Count = 0 Then
                    Exit Function
                Else:
'                    MsgBox ssetpillar.Count
                    Call Get_LT_in_MP(ssetpillar)
                    Exit Function
                End If
            End If
        ODrcs_PB.Next
        Loop
    End If
Next
End Function
Private Function Get_LT_in_MP(sset_LT As AcadSelectionSet)
Dim acadOBJ As Object
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
Dim boolVal As Boolean
Dim ODT As ODTables
Set ODT = acad.Projects(ThisDrawing).ODTables
Dim OD_LT As ODTable
Set OD_LT = ODT.Item("LT_CABLE")
Dim ODrcs_LT As ODRecords
Dim ODrc_LT As ODRecord
For Each acadOBJ In sset_LT
    If Left$(acadOBJ.Layer, 2) = "LT" Then
        Set ODrcs_LT = OD_LT.GetODRecords
        ODrcs_LT.Init acadOBJ, False, False
        If ODrcs_LT.IsDone = True Then
            boolVal = True
        End If
         Do Until ODrcs_LT.IsDone 'This runs for number of circuit breakers in the given pillar
            Set ODrc_LT = ODrcs_LT.Record
            ReDim Preserve LT_Array(iLT_array) As Long
            For i = 0 To UBound(LT_Array) Step 1
                If ODrc_LT.ObjectID = LT_Array(i) Then
'                    MsgBox "Same LT_Cable:" & ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & ODrc_LT.Item(iPillar2).Value & "." & ODrc_LT.Item(i2_CKTNO).Value
                    GoTo nextrecord
                End If
            Next i
            iLT_array = iLT_array + 1
            ReDim Preserve LT_Array(iLT_array) As Long
            LT_Array(iLT_array) = ODrc_LT.ObjectID
'            Add this LT_Cable to the existing array
                If ODrc_LT.Item(iPillar2).Value = "SERVICE" Then
                    MsgBox ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & "SERVICE POINT"
                    GoTo nextrecord
                Else:
                    MsgBox ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & ODrc_LT.Item(iPillar2).Value & "." & ODrc_LT.Item(i2_CKTNO).Value
                    Call To_Pillar(ODrc_LT.Item(iPillar2).Value, ODrc_LT.Item(i2_CKTNO).Value, LT_Array())
                    GoTo nextrecord
                End If
nextrecord:
            ODrcs_LT.Next
        Loop
    End If
Next
End Function
Private Function OPEN_CB(sset_OPEN As AcadSelectionSet)
Dim acadOBJ As Object
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
Dim boolVal As Boolean
Dim ODT As ODTables
Set ODT = acad.Projects(ThisDrawing).ODTables
Dim OD_LT As ODTable
Set OD_LT = ODT.Item("LT_CABLE")
Dim ODrcs_LT As ODRecords
Dim ODrc_LT As ODRecord
For Each acadOBJ In sset_OPEN
    If Left$(acadOBJ.Layer, 2) = "LT" Then
        Set ODrcs_LT = OD_LT.GetODRecords
        ODrcs_LT.Init acadOBJ, False, False
        If ODrcs_LT.IsDone = True Then
            boolVal = True
        End If
         Do Until ODrcs_LT.IsDone 'This runs for number of circuit breakers in the given pillar
            Set ODrc_LT = ODrcs_LT.Record
            ReDim Preserve LT_Array(iLT_array) As Long
            For i = 0 To UBound(LT_Array) Step 1
                If ODrc_LT.ObjectID = LT_Array(i) Then
'                    MsgBox "Same LT_Cable:" & ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & ODrc_LT.Item(iPillar2).Value & "." & ODrc_LT.Item(i2_CKTNO).Value
                    GoTo nextrecord
                End If
            Next i
            iLT_array = iLT_array + 1
            ReDim Preserve LT_Array(iLT_array) As Long
            LT_Array(iLT_array) = ODrc_LT.ObjectID
'            Add this LT_Cable to the existing array
                If ODrc_LT.Item(iPillar2).Value = "SERVICE" Then
                    MsgBox ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & "SERVICE POINT"
                    GoTo nextrecord
                Else:
                    MsgBox ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & ODrc_LT.Item(iPillar2).Value & "." & ODrc_LT.Item(i2_CKTNO).Value
                    Call To_Pillar(ODrc_LT.Item(iPillar2).Value, ODrc_LT.Item(i2_CKTNO).Value, LT_Array())
                    GoTo nextrecord
                End If
nextrecord:
            ODrcs_LT.Next
        Loop
    End If
Next
End Function
Private Function Get_BB_LT(pillarno As Long)
Dim acadOBJ As Object
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
Dim boolVal As Boolean
Dim boolBC As Boolean
Dim OD_PB As ODTable
Set OD_PB = acad.Projects(ThisDrawing).ODTables.Item("PILLAR_BOUNDARY")
Dim ODrcs_PB As ODRecords
Dim ODrc_PB As ODRecord
Dim iPillar As Long
Dim Pillar_ObjectID As Long
For Each acadOBJ In ThisDrawing.ModelSpace
    If Left$(acadOBJ.Layer, 6) = "PILLAR" Then
        Set ODrcs_PB = OD_PB.GetODRecords
        ODrcs_PB.Init acadOBJ, False, False
        If ODrcs_PB.IsDone = True Then
            boolVal = True
        End If
        Do Until ODrcs_PB.IsDone
          Set ODrc_PB = ODrcs_PB.Record
          iPillar = ODrc_PB.Item(iPillar_NumberFieldIndex).Value
            If iPillar = pillarno Then
                Pillar_ObjectID = ODrc_PB.ObjectID
                'create selection set for this pillar
                If ODrc_PB.Item(iTypeFieldIndex).Value <> "MP" Then
                    Dim SS_LT_BB As AcadSelectionSet
                    Dim acadLWPOLY As AcadLWPolyline
                    Dim minExt As Variant
                    Dim maxExt As Variant
                    Set acadLWPOLY = Nothing
                    Set acadLWPOLY = ThisDrawing.ObjectIdToObject(Pillar_ObjectID)
                    acadLWPOLY.GetBoundingBox minExt, maxExt
                    'declaration of selesection set
                    ReDim Preserve PillarArray(iPillararray) As Long
                    PillarArray(iPillararray) = Pillar_ObjectID
                    'Create a filter for getting only CIRCUIT BLOCKS
                    Dim mode As Integer
                    mode = acSelectionSetCrossing
'                   MsgBox PillarArray(iPillararray) & "." & (iPillararray + 1)
                    Set SS_LT_BB = Nothing
                    Set SS_LT_BB = ThisDrawing.SelectionSets.Add(PillarArray(iPillararray) & "." & (iPillararray + 1))
                    Dim gpCode(0) As Integer
                    Dim dataValue(0) As Variant
                    gpCode(0) = 8
                    Dim iDataValue As Integer
                    iDataValue = 0
                    For iDataValue = 0 To (ThisDrawing.Layers.Count - 1) Step 1
                        If Left$(ThisDrawing.Layers.Item(iDataValue).Name, 2) = "LT" Then
                            dataValue(0) = ThisDrawing.Layers.Item(iDataValue).Name
                        End If
                    Next iDataValue
                    Dim groupCode As Variant, dataCode As Variant
                    groupCode = gpCode
                    dataCode = dataValue
                    SS_LT_BB.Select mode, maxExt, minExt, groupCode, dataCode
                    If SS_LT_BB.Count = 0 Then
                        Exit Function
                    Else:
                    Call GET_LT_BB(SS_LT_BB)
                        Exit Function
                    End If
                Else:
                    Exit Function
                End If
            End If
            ODrcs_PB.Next
        Loop
    End If
Next
End Function
Private Function GET_LT_BB(SS_LT_BB As AcadSelectionSet)
Dim acadOBJ As Object
Dim acad As AcadMap
Set acad = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
Dim boolVal As Boolean
Dim ODT As ODTables
Set ODT = acad.Projects(ThisDrawing).ODTables
Dim OD_LT As ODTable
Set OD_LT = ODT.Item("LT_CABLE")
Dim ODrcs_LT As ODRecords
Dim ODrc_LT As ODRecord
For Each acadOBJ In SS_LT_BB
    If Left$(acadOBJ.Layer, 2) = "LT" Then
        Set ODrcs_LT = OD_LT.GetODRecords
        ODrcs_LT.Init acadOBJ, False, False
        If ODrcs_LT.IsDone = True Then
            boolVal = True
        End If
         Do Until ODrcs_LT.IsDone 'This runs for number of circuit breakers in the given pillar
            Set ODrc_LT = ODrcs_LT.Record
            ReDim Preserve LT_Array(iLT_array) As Long
            For i = 0 To UBound(LT_Array) Step 1
                If ODrc_LT.ObjectID = LT_Array(i) Then
'                    MsgBox "Same LT_Cable:" & ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & ODrc_LT.Item(iPillar2).Value & "." & ODrc_LT.Item(i2_CKTNO).Value
                    GoTo nextrecord
                End If
            Next i
            iLT_array = iLT_array + 1
            ReDim Preserve LT_Array(iLT_array) As Long
            LT_Array(iLT_array) = ODrc_LT.ObjectID
'            Add this LT_Cable to the existing array
                If ODrc_LT.Item(i4m_CKTNO).Value = "BB" Or ODrc_LT.Item(i2_CKTNO).Value = "BB" Then
                    MsgBox ODrc_LT.Item(iPillar4m).Value & "." & ODrc_LT.Item(i4m_CKTNO).Value & "-->" & ODrc_LT.Item(iPillar2).Value & "." & ODrc_LT.Item(i2_CKTNO).Value
                    Call To_Pillar(ODrc_LT.Item(iPillar2).Value, ODrc_LT.Item(i2_CKTNO).Value, LT_Array())
                    Exit Function
                    GoTo nextrecord
                End If
nextrecord:
            ODrcs_LT.Next
        Loop
    End If
Next

End Function