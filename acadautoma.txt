'02.LoopLayer

Public Function SpecPoint(i As Integer) As Variant
	Dim varPt As Variant
	If i Mod 1 = 0 Then
	varPt = ThisDrawing.Utility.GetPoint(, "Specify the starting point")
	Else
	varPt = ThisDrawing.Utility.GetPoint(varPt, "Specify the ending point")
	End If
	SpecPoint = varPt
End Function

Public Sub LoopLines()
	Dim strPmt As String
	Dim varPt1 As Variant
	Dim varPt2 As Variant
	Dim i As Integer
	i = 1
	Do
	On Error Resume Next
	varPt1 = SpecPoint(i)
	If Err Then
	Err.Clear
	Exit Do
	End If
	On Error GoTo 0
	i = i + 1
	varPt2 = SpecPoint(i)
	ThisDrawing.ModelSpace.AddLine varPt1, varPt2
	Loop
	On Error GoTo 0
End Sub

Public Sub MyMacro()
	Dim Count As Integer
	Dim NumericString As String
	Count = 100
	NumericString = "555"
	MsgBox "Integer: " & Count & vbCrLf & _
	"String: " & NumericString
	Count = NumericString
	MsgBox "Integer: " & Count
End Sub

Public Sub IterateLayers()
	Dim Layer As AcadLayer
	For Each Layer In ThisDrawing.Layers
	MsgBox Layer.Name
	Next Layer
End Sub

Public Sub WhatColor()
	Dim Layer As AcadLayer
	Dim Answer As String
	For Each Layer In ThisDrawing.Layers
	Answer = InputBox("Enter color name: ")
	'user pressed cancel
	If Answer = "" Then Exit Sub
	Select Case UCase(Answer)
	Case "RED"
	Layer.color = acRed
	Case "YELLOW"
	Layer.color = acYellow
	Case "GREEN"
	Layer.color = acGreen
	Case "CYAN"
	Layer.color = acCyan
	Case "BLUE"
	Layer.color = acBlue
	Case "MAGENTA"
	Layer.color = acMagenta
	Case "WHITE"
	Layer.color = acWhite
	Case Else
	If CInt(Answer) > 0 And CInt(Answer) < 256 Then
	Layer.color = CInt(Answer)
	Else
	MsgBox UCase(Answer) & " is an invalid color name", _
	vbCritical, "Invalid Color Selected"
	End If
	End Select
	Next Layer
End Sub

Sub With_()
	Dim mylayer As AcadLayer
	mylayer = ThisDrawing.ActiveLayer
	mylayer.color = acBlue
	mylayer.Linetype = "continuous"
	mylayer.Lineweight = acLnWtByLwDefault
	mylayer.Freeze = False
	mylayer.LayerOn = True
	mylayer.Lock = False
	'Using the With ... End With statement, you do less typing:
	mylayer = ThisDrawing.ActiveLayer
	With mylayer
	.color = acBlue
	.Linetype = "continuous"
	.Lineweight = acLnWtByLwDefault
	.Freeze = False
	.LayerOn = True
	.Lock = False
	End With
End Sub


Sub selectCase()
	Select Case UCase(ColorName)
	Case "RED"
	Layer.color = acRed
	Case "YELLOW"
	Layer.color = acYellow
	Case "GREEN"
	Layer.color = acGreen
	Case "CYAN"
	Layer.color = acCyan
	Case "BLUE"
	Layer.color = acBlue
	Case "MAGENTA"
	Layer.color = acMagenta
	Case "WHITE"
	Layer.color = acWhite
	Case Else
	If CInt(ColorName) > 0 And CInt(ColorName) < 256 Then
	Layer.color = CInt(ColorName)
	Else
	MsgBox UCase(ColorName) & " is an invalid color name", _
	vbCritical, "Invalid Color Selected"
	End If
	End Select
End Sub
Sub do_While()
	'Dim Length(0 To 2) As Double
	Dim Index As Integer
	Index = 1
	Length = 4
	Do While Index - 1 < Length
	Character = Asc(Mid(Name, Index, 1))
	Select Case Character
	Case 36, 45, 48 To 57, 65 To 90, 95
	IsOK = True
	Case Else
	IsOK = False
	Exit Sub
	End Select
	Index = Index + 1
	Loop
End Sub

Sub for_Next()
	Dim Point(0 To 2) As Double
	Dim Index As Integer
	For Index = 0 To 2
	Point(Index) = 0
	Next Index
End Sub

Sub for_each_next()
	Dim Layer As AcadLayer
	For Each Layer In ThisDrawing.Layers
	MsgBox Layer.Name
	Next Layer
End Sub


'06.LayerControl



Sub Example_DeleteLayers()
	' This example creates a new layer called "A". Draw a polyline on this layer
	' Create another layer called "B", draw a line on it.
	' Create the third layer called "C", Draw a circle on it.
	' Then delete layer "A" and "B", leave "C" and its object to save as a new drawing
	Dim objLayer As AcadLayer
	'1~~~~~~~~~~~~~~~~~~
	'Create the new layer A
	Set objLayer = ThisDrawing.Layers.Add("A")
	objLayer.LayerOn = True
	ThisDrawing.ActiveLayer = ThisDrawing.Layers("A")
	'Draw polyline
	Dim objPline As AcadLWPolyline
	Dim varpts(15) As Double
	varpts(0) = 1: varpts(1) = 0
	varpts(2) = 2: varpts(3) = 0
	varpts(4) = 3: varpts(5) = 1
	varpts(6) = 3: varpts(7) = 2
	varpts(7) = 2: varpts(9) = 3
	varpts(10) = 1: varpts(11) = 3
	varpts(12) = 0: varpts(13) = 2
	varpts(14) = 0: varpts(15) = 1
	Set objPline = ThisDrawing.ModelSpace.AddLightWeightPolyline(varpts)
	objPline.Closed = True
	objPline.SetBulge 1, 0.5
	objPline.SetBulge 3, 0.5
	objPline.SetBulge 5, 0.5
	objPline.SetBulge 7, 0.5
	objPline.Update
	'2~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'Create layer "B"
	Set objLayer = ThisDrawing.Layers.Add("B")
	objLayer.LayerOn = True
	ThisDrawing.ActiveLayer = ThisDrawing.Layers("B")
	'Draw a line
	Dim dblStart(2) As Double
	Dim dblEnd(2) As Double
	Dim objLine As AcadLine
	dblStart(0) = 2: dblStart(1) = 2: dblStart(2) = 0
	dblEnd(0) = 5: dblEnd(1) = 5: dblEnd(2) = 0
	Set objLine = ThisDrawing.ModelSpace.AddLine(dblStart, dblEnd)
	ThisDrawing.Regen acAllViewports
	'3~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	'Create layer "C"
	Set objLayer = ThisDrawing.Layers.Add("C")
	objLayer.LayerOn = True
	ThisDrawing.ActiveLayer = ThisDrawing.Layers("C")
	'Draw circle
	Dim aCircle As AcadCircle
	Dim Center(0 To 2) As Double
	Dim Radius As Double
	Center(0) = 4
	Center(1) = 4
	Radius = 2
	Set aCircle = ThisDrawing.ModelSpace.AddCircle(Center, Radius)
	ThisDrawing.Regen acAllViewports
	'4~~~~~~~~~~~~~~~~~~~~~~~~
	'Define a selection set
	-1-D:\_S\Excel workbook\CAD_Col\06.LayerControl\mk_DeleteLayers.bas Thursday, May 20, 2021 5:29 PM
	Dim SSet As AcadSelectionSet
	Dim intCode(1) As Integer
	Dim varData(1) As Variant
	'Delete objects on layer B
	intCode(0) = 8: varData(0) = "B" 'only select items on layer "B"
	intCode(1) = 67: varData(1) = 0 'only select items in modelspace - error without this filter
	On Error Resume Next
	ThisDrawing.SelectionSets.Item("B").Delete
	On Error GoTo 0
	Set SSet = ThisDrawing.SelectionSets.Add("B")
	SSet.Select acSelectionSetAll, , , intCode, varData
	SSet.Highlight True
	SSet.Erase
	SSet.Delete
	'Delete objects on layer A
	intCode(0) = 8: varData(0) = "A" 'only select items on layer "A"
	intCode(1) = 67: varData(1) = 0 'only select items in modelspace - error without this filter
	On Error Resume Next
	ThisDrawing.SelectionSets.Item("A").Delete
	On Error GoTo 0
	Set SSet = ThisDrawing.SelectionSets.Add("A")
	SSet.Select acSelectionSetAll, , , intCode, varData
	SSet.Highlight True
	SSet.Erase
	SSet.Delete
	'Delete not used layers A and B
	ThisDrawing.ActiveLayer = ThisDrawing.Layers("C")
	ThisDrawing.Layers.Item("A").Delete
	ThisDrawing.Layers.Item("B").Delete
End Sub

Public Sub AddLayer()
	Dim strLayerName As String
	Dim objLayer As AcadLayer
	strLayerName = InputBox("Name of Layer to add: ")
	If "" = strLayerName Then Exit Sub ' exit if no name entered
	On Error Resume Next ' handle exceptions inline
	'check to see if layer already exists
	Set objLayer = ThisDrawing.Layers(strLayerName)
	If objLayer Is Nothing Then
	Set objLayer = ThisDrawing.Layers.Add(strLayerName)
	If objLayer Is Nothing Then ' check if obj has been set
	MsgBox "Unable to Add '" & strLayerName & "'"
	Else
	MsgBox "Added Layer '" & objLayer.Name & "'"
	End If
	Else
	MsgBox "Layer already existed"
	End If
End Sub

Public Sub ChangeEntityLayer()
	On Error Resume Next ' handle exceptions inline
	Dim objEntity As AcadEntity
	Dim varPick As Variant
	Dim strLayerName As String
	Dim objLayer As AcadLayer
	ThisDrawing.Utility.GetEntity objEntity, varPick, "Select an entity"
	If objEntity Is Nothing Then
	MsgBox "No entity was selected"
	Exit Sub ' exit if no entity picked
	End If
	strLayerName = InputBox("Enter a new Layer name: ")
	If "" = strLayerName Then Exit Sub ' exit if no name entered
	Set objLayer = ThisDrawing.Layers(strLayerName)
	If objLayer Is Nothing Then
	MsgBox "Layer was not recognized"
	Exit Sub ' exit if layer not found
	End If
	objEntity.Layer = strLayerName ' else change entity layer
End Sub


'check to see if layer is active or not....
	Public Function IsLayerActive(strLayerName As String) As Boolean
	IsLayerActive = False 'assume failure
	If 0 = StrComp(ThisDrawing.ActiveLayer.Name, strLayerName, _
	vbTextCompare) Then
	IsLayerActive = True
	End If


Public Sub LayerActive()
	Dim strLayerName As String
	strLayerName = InputBox("Name of the Layer to check: ")
	If IsLayerActive(strLayerName) Then
	MsgBox "'" & strLayerName & "' is active"
	Else
	MsgBox "'" & strLayerName & "' is not active"
	End If
End Sub

Public Sub CheckForLayerByIteration()
	Dim objLayer As AcadLayer
	Dim strLayerName As String
	strLayerName = InputBox("Enter a Layer name to search for: ")
	If "" = strLayerName Then Exit Sub ' exit if no name entered
	For Each objLayer In ThisDrawing.Layers ' iterate layers
	If 0 = StrComp(objLayer.Name, strLayerName, vbTextCompare) Then
	MsgBox "Layer '" & strLayerName & "' exists"
	Exit Sub ' exit after finding layer
	End If
	Next objLayer
	MsgBox "Layer '" & strLayerName & "' does not exist"
End Sub

Public Sub CheckForLayerByException()
	Dim strLayerName As String
	Dim objLayer As AcadLayer
	strLayerName = InputBox("Enter a Layer name to search for: ")
	If "" = strLayerName Then Exit Sub ' exit if no name entered
	On Error Resume Next ' handle exceptions inline
	Set objLayer = ThisDrawing.Layers(strLayerName)
	If objLayer Is Nothing Then ' check if obj has been set
	MsgBox "Layer '" & strLayerName & "' does not exist"
	Else
	MsgBox "Layer '" & objLayer.Name & "' exists"
	End If
End Sub

Public Sub DeleteLayer()
	On Error Resume Next ' handle exceptions inline
	Dim strLayerName As String
	Dim objLayer As AcadLayer
	strLayerName = InputBox("Layer name to delete: ")
	If "" = strLayerName Then Exit Sub ' exit if no old name
	Set objLayer = ThisDrawing.Layers(strLayerName)
	If objLayer Is Nothing Then ' exit if not found
	MsgBox "Layer '" & strLayerName & "' not found"
	Exit Sub
	End If
	objLayer.Delete ' try to delete it
	If Err Then ' check if it worked
	MsgBox "Unable to delete layer: " & vbCr & Err.Description
	Else
	MsgBox "Layer '" & strLayerName & "' deleted"
	End If
End Sub

Public Sub SaveAsType()
	Dim iSaveAsType As Integer
	iSaveAsType = ThisDrawing.Application.Preferences.OpenSave.SaveAsType
	Select Case iSaveAsType
	Case acR12_dxf
	MsgBox "Current save as format is R12_DXF", vbInformation
	Case ac2000_dwg
	MsgBox "Current save as format is 2000_DWG", vbInformation
	Case ac2000_dxf
	MsgBox "Current save as format is 2000_DXF", vbInformation
	Case ac2000_Template
	MsgBox "Current save as format is 2000_Template", vbInformation
	Case ac2004_dwg, acNative
	MsgBox "Current save as format is 2004_DWG", vbInformation
	Case ac2004_dxf
	MsgBox "Current save as format is 2004_DXF", vbInformation
	Case ac2004_Template
	MsgBox "Current save as format is 2004_Template", vbInformation
	Case acUnknown
	MsgBox "Current save as format is Unknown or Read-Only", vbInformation
	End Select
End Sub


Public Sub ListLayers()
	Dim objLayer As AcadLayer
	For Each objLayer In ThisDrawing.Layers
	Debug.Print objLayer.Name
	Next
	End Sub
	Public Sub ListLayersManually()
	Dim objLayers As AcadLayers
	Dim objLayer As AcadLayer
	Dim intI As Integer
	Set objLayers = ThisDrawing.Layers
	For intI = 0 To objLayers.Count - 1
	Set objLayer = objLayers(intI)
	MsgBox objLayer.Name
	Next
End Sub

Public Sub ListLayersBackwards()
	Dim objLayers As AcadLayers
	Dim objLayer As AcadLayer
	Dim intI As Integer
	Set objLayers = ThisDrawing.Layers
	For intI = objLayers.Count - 1 To 0 Step -1
	Set objLayer = objLayers(intI)
	Debug.Print objLayer.Name
	Next
End Sub

Public Sub ShowOnlyLayer()
	On Error Resume Next ' handle exceptions inline
	Dim strLayerName As String
	Dim objLayer As AcadLayer
	strLayerName = InputBox("Enter a Layer name to show: ")
	If "" = strLayerName Then Exit Sub ' exit if no name entered
	For Each objLayer In ThisDrawing.Layers
	objLayer.LayerOn = False ' turn off all the layers
	Next objLayer
	Set objLayer = ThisDrawing.Layers(strLayerName)
	If objLayer Is Nothing Then
	MsgBox "Layer does not exist"
	Exit Sub ' exit if layer not found
	End If
	objLayer.LayerOn = True ' turn on the desired layer
End Sub


Public Sub RenameLayer()
	On Error Resume Next ' handle exceptions inline
	Dim strLayerName As String
	Dim objLayer As AcadLayer
	strLayerName = InputBox("Original Layer name: ")
	If "" = strLayerName Then Exit Sub ' exit if no old name
	Set objLayer = ThisDrawing.Layers(strLayerName)
	If objLayer Is Nothing Then ' exit if not found
	MsgBox "Layer '" & strLayerName & "' not found"
	Exit Sub
	End If
	strLayerName = InputBox("New Layer name: ")
	If "" = strLayerName Then Exit Sub ' exit if no new name
	objLayer.Name = strLayerName ' try and change name
	If Err Then ' check if it worked
	MsgBox "Unable to rename layer: " & vbCr & Err.Description
	Else
	MsgBox "Layer renamed to '" & strLayerName & "'"
	End If
End Sub


'07.UtilityObjects

Sub pointWithinPoly()
	Dim basePnt As Variant
	Dim selectedPoly As AcadObject
	Dim point As Variant
	Dim count As Integer
	Dim pl_coordinates As Variant
	Dim pi As Double
	Dim numberVertices As Integer
	Dim verticesCount As Integer
	Dim deltaAngleSpread As Double
	Dim deltaX As Double
	Dim deltaY As Double
	pi = 3.141592654
	ThisDrawing.Utility.GetEntity selectedPoly, basePnt, "Select a polyline:"
	point = ThisDrawing.Utility.GetPoint(, vbCrLf & "Click on point:")
	pl_coordinates = selectedPoly.Coordinates
	count = 0
	numberVertices = (UBound(pl_coordinates) + 1) / 2
	ReDim angle(0 To numberVertices - 1) As Double
	verticesCount = 0
	Do While count <= (UBound(pl_coordinates))
	deltaX = (pl_coordinates(count) - point(0))
	deltaY = (pl_coordinates(count + 1) - point(1))
	If deltaX <> 0 And deltaY = 0 Then
	If deltaX < 0 Then
	angle(verticesCount) = pi
	Else
	angle(verticesCount) = 0
	End If
	ElseIf deltaX = 0 And deltaY <> 0 Then
	If deltaY < 0 Then
	angle(verticesCount) = 1.5 * pi
	Else
	angle(verticesCount) = pi / 2
	End If
	ElseIf deltaX = 0 And deltaY = 0 Then
	deltaAngleSpread = 2 * pi
	GoTo PASTWHILELOOP
	Else
	If deltaX > 0 And deltaY > 0 Then
	angle(verticesCount) = Atn(Abs(deltaY) / Abs(deltaX))
	ElseIf deltaX < 0 And deltaY > 0 Then
	angle(verticesCount) = pi - Atn(Abs(deltaY) / Abs(deltaX))
	ElseIf deltaX > 0 And deltaY < 0 Then
	angle(verticesCount) = 2 * pi - Atn(Abs(deltaY) / Abs(deltaX))
	ElseIf deltaX < 0 And deltaY < 0 Then
	angle(verticesCount) = Atn(Abs(deltaY) / Abs(deltaX)) + pi
	End If
	End If
	count = count + 2
	verticesCount = verticesCount + 1
	Loop
	verticesCount = 1
	Do While verticesCount <= numberVertices
	If verticesCount = numberVertices Then
	If Abs(angle(0) - angle(verticesCount - 1)) > pi Then
	If angle(0) < angle(verticesCount - 1) Then
	-1-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_CheckPt_zinPolyline_orNot.bas Thursday, May 20, 2021 5:34 PM
	deltaAngleSpread = (2 * pi - angle(verticesCount - 1) + angle(0)) +
	deltaAngleSpread
	Else
	deltaAngleSpread = -1 * (2 * pi - angle(0) + angle(verticesCount - 1))
	+ deltaAngleSpread
	End If
	ElseIf Abs(angle(0) - angle(verticesCount - 1)) = pi Then
	deltaAngleSpread = 2 * pi
	Exit Do
	Else
	deltaAngleSpread = angle(0) - angle(verticesCount - 1) + deltaAngleSpread
	End If
	Else
	If Abs(angle(verticesCount) - angle(verticesCount - 1)) > pi Then
	If angle(verticesCount) < angle(verticesCount - 1) Then
	deltaAngleSpread = (2 * pi - angle(verticesCount - 1) +
	angle(verticesCount)) + deltaAngleSpread
	Else
	deltaAngleSpread = -1 * (2 * pi - angle(verticesCount) +
	angle(verticesCount - 1)) + deltaAngleSpread
	End If
	ElseIf Abs(angle(verticesCount) - angle(verticesCount - 1)) = pi Then
	deltaAngleSpread = 2 * pi
	Exit Do
	Else
	deltaAngleSpread = angle(verticesCount) - angle(verticesCount - 1) +
	deltaAngleSpread
	End If
	End If
	verticesCount = verticesCount + 1
	Loop
	PASTWHILELOOP:
	If Abs(deltaAngleSpread) >= 6.28 Then
	MsgBox "Ohh, this point is inside this poly baby!! :)"
	Else
	MsgBox "Sorry, this point is not inside this poly :("
	End If
End Sub

Function getDistance(vFirstPickPnt As Variant, vSecondPickPnt As Variant) As Double
	Dim objLine As AcadLine
	Dim dblDistance As Double
	Set objLine = ThisDrawing.ModelSpace.AddLine(vFirstPickPnt, vSecondPickPnt)
	dblDistance = objLine.Length
	objLine.Delete
	getDistance = dblDistance
End Function

Sub getDistance_Sample()
	Dim vFirstPickPnt, vSecondPickPnt As Variant
	vFirstPickPnt = ThisDrawing.Utility.GetPoint(, "Pick first point: ")
	vSecondPickPnt = ThisDrawing.Utility.GetPoint(, "Pick Second point: ")
	MsgBox getDistance(vFirstPickPnt, vSecondPickPnt)
End Sub

Option Explicit
Function IsExcelRunning() As Boolean
	Dim objXL As Object
	On Error Resume Next
	Set objXL = GetObject(, "Excel.Application")
	IsExcelRunning = (Err.Number = 0)
	Set objXL = Nothing
	Err.Clear
End Function

Public Sub WriteCustoms()
	Dim Cust() As String
	Dim col As Long
	Dim row As Long
	Cust = GetCustoms("D:\TEST.dwg")
	Dim sXlFilNam As String
	sXlFilNam = "D:\TEST.xls"
	'***Begin code from Randall Rath******
	Dim oXL As Object
	Dim blnXLRunning As Boolean
	blnXLRunning = IsExcelRunning()
	If blnXLRunning Then
	Set oXL = GetObject(, "Excel.Application")
	Else
	Set oXL = CreateObject("Excel.Application")
	oXL.Visible = False
	oXL.UserControl = False
	oXL.DisplayAlerts = False
	End If
	'***End code from Randall Rath******
	Dim oWb As Object
	Dim oWs As Object
	Set oWb = oXL.Workbooks.Open(sXlFilNam)
	If oWb Is Nothing Then
	MsgBox "The Excel file " & sXlFilNam & " not found" & _
	"Try again."
	GoTo Exit_Here
	End If
	Set oWs = oWb.Worksheets(1)
	oWs.Activate
	' write data to Excel
	' headers:
	oWs.Cells(row + 1, 1) = "NAME"
	oWs.Cells(row + 1, 2) = "VALUE"
	For row = 0 To UBound(Cust, 1)
	oWs.Cells(row + 2, 1) = Cust(row, 0)
	oWs.Cells(row + 2, 2) = Cust(row, 1)
	Next
	oWs.Columns.autofit
	-1-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_DrawingProperties.bas Thursday, May 20, 2021 5:34 PM
	Exit_Here:
	Set oWs = Nothing
	oWb.Save: oWb.Close
	Set oWb = Nothing
	oXL.Quit
	Set oXL = Nothing
	MsgBox "Done"
End Sub

Function GetCustoms(fileName As String) As Variant
	Dim Value As String
	Dim Cnt As Long
	Dim Num As Long
	Dim Index As Long
	Dim CustomKey As String
	Dim CustomValue As String
	Dim doc As AcadDocument
	Dim Sum As AcadSummaryInfo
	Set doc = ThisDrawing.Application.Documents.Open(fileName)
	Set Sum = doc.SummaryInfo
	Num = Sum.NumCustomInfo
	ReDim Cust(0 To Num - 1, 0 To 1) As String
	For Index = 0 To Num - 1
	Sum.GetCustomByIndex Index, CustomKey, CustomValue
	Cust(Cnt, 0) = CustomKey
	Cust(Cnt, 1) = CustomValue
	Cnt = Cnt + 1
	Next Index
	doc.Close
	Set Sum = Nothing
	Set doc = Nothing
	GetCustoms = Cust
End Function

Option Explicit
Function IsExcelRunning() As Boolean
	Dim objXL As Object
	On Error Resume Next
	Set objXL = GetObject(, "Excel.Application")
	IsExcelRunning = (Err.Number = 0)
	Set objXL = Nothing
	Err.Clear
End Function

Public Sub WriteCustom()
	Dim Value As String
	Value = GetCustom("D:\TEST.dwg", _
	"PROJECT1") ' change dwg name and key here
	Dim sXlFilNam As String
	-2-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_DrawingProperties.bas Thursday, May 20, 2021 5:34 PM
	sXlFilNam = "D:\Test.xls" ' change the existing Excel file name here
	'***Begin code from Randall Rath******
	Dim oXL As Object
	Dim blnXLRunning As Boolean
	blnXLRunning = IsExcelRunning()
	If blnXLRunning Then
	Set oXL = GetObject(, "Excel.Application")
	Else
	Set oXL = CreateObject("Excel.Application")
	oXL.Visible = False
	oXL.UserControl = False
	oXL.DisplayAlerts = False
	End If
	'***End code from Randall Rath******
	Dim oWb As Object
	Dim oWs As Object
	Set oWb = oXL.Workbooks.Open(sXlFilNam)
	If oWb Is Nothing Then
	MsgBox "The Excel file " & sXlFilNam & " not found" & _
	"Try again."
	GoTo Exit_Here
	End If
	Set oWs = oWb.Worksheets(1)
	oWs.Activate
	oWs.Cells(11, 2) = Value
	'Or:
	' oWs.Range("B11").Value2 = Value
	Exit_Here:
	Set oWs = Nothing
	oWb.Save: oWb.Close
	Set oWb = Nothing
	oXL.Quit
	Set oXL = Nothing
End Sub

Private Function GetCustom(fileName As String, Key As String) As String
	Dim Value As String
	Dim doc As AcadDocument
	Set doc = ThisDrawing.Application.Documents.Open(fileName)
	doc.SummaryInfo.GetCustomByKey Key, Value
	doc.Close
	Set doc = Nothing
	GetCustom = Value
End Function

Option Explicit
Function IsExcelRunning() As Boolean
	Dim objXL As Object
	On Error Resume Next
	Set objXL = GetObject(, "Excel.Application")
	IsExcelRunning = (Err.Number = 0)
	Set objXL = Nothing
	Err.Clear
End Function

Public Sub WriteCustom()
	Dim Value As String
	Value = GetCustom("D:\TEST.dwg", _
	"PROJECT1") ' change dwg name and key here
	Dim sXlFilNam As String
	sXlFilNam = "D:\Test.xls" ' change the existing Excel file name here
	'***Begin code from Randall Rath******
	Dim oXL As Object
	Dim blnXLRunning As Boolean
	blnXLRunning = IsExcelRunning()
	If blnXLRunning Then
	Set oXL = GetObject(, "Excel.Application")
	Else
	Set oXL = CreateObject("Excel.Application")
	oXL.Visible = False
	oXL.UserControl = False
	oXL.DisplayAlerts = False
	End If
	'***End code from Randall Rath******
	Dim oWb As Object
	Dim oWs As Object
	Set oWb = oXL.Workbooks.Open(sXlFilNam)
	If oWb Is Nothing Then
	MsgBox "The Excel file " & sXlFilNam & " not found" & _
	"Try again."
	GoTo Exit_Here
	End If
	Set oWs = oWb.Worksheets(1)
	oWs.Activate
	oWs.Cells(11, 2) = Value
	'Or:
	' oWs.Range("B11").Value2 = Value
	Exit_Here:
	Set oWs = Nothing
	oWb.Save: oWb.Close
	Set oWb = Nothing
	oXL.Quit
	-1-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_dwgProperties.bas Thursday, May 20, 2021 5:34 PM
	Set oXL = Nothing
End Sub

Private Function GetCustom(fileName As String, Key As String) As String
	Dim Value As String
	Dim doc As AcadDocument
	Set doc = ThisDrawing.Application.Documents.Open(fileName)
	doc.SummaryInfo.GetCustomByKey Key, Value
	doc.Close
	Set doc = Nothing
	GetCustom = Value
End Function

Sub FilletLines()
	Dim objLine1 As AcadLine
	Dim objLine2 As AcadLine
	Dim varPt As Variant
	Dim dblFill, newSysVar As Double
	Dim comString As String
	dblFill = ThisDrawing.GetVariable("FILLETRAD")
	On Error GoTo ProblemHere
	MsgBox CStr(dblFill)
	newSysVar = CDbl(InputBox("Enter fillet radius:", "Fillet Radius Value", "2,5"))
	ThisDrawing.Utility.GetEntity objLine1, varPt, "Select first line"
	ThisDrawing.Utility.GetEntity objLine2, varPt, "Select second line"
	If objLine1 Is Nothing Or objLine2 Is Nothing Then
	Exit Sub
	Else
	If TypeOf objLine1 Is AcadLine And _
	TypeOf objLine2 Is AcadLine Then
	ThisDrawing.SetVariable "FILLETRAD", newSysVar
	comString = "_FILLET" & vbCr & "(HANDENT " & Chr(34) & CStr(objLine1.Handle) & Chr(34) & ")" &
	vbCr & _
	"(HANDENT " & Chr(34) & CStr(objLine2.Handle) & Chr(34) & ")" & vbCr
	ThisDrawing.SendCommand comString
	Else
	MsgBox "Incorrect object type"
	Exit Sub
	End If
	End If
	ThisDrawing.SetVariable "FILLETRAD", dblFill
	ProblemHere:
	If Err Then
	ThisDrawing.SetVariable "FILLETRAD", dblFill
	MsgBox vbCr & Err.Description
	End If
End Sub

Sub IPNT()
	Dim Poly1 As AcadLWPolyline
	Dim Poly2 As AcadLWPolyline
	Dim pts As Variant
	On Error Resume Next
	ThisDrawing.Utility.GetEntity Poly1, pickPt, "Pick Poly1:"
	If Err Then Exit Sub
	ThisDrawing.Utility.GetEntity Poly2, pickPt, "Pick Poly2:"
	If Err Then Exit Sub
	pts = Poly1.IntersectWith(Poly2, acExtendNone)
	MsgBox "Found " & (UBound(pts) + 1) / 3 & " intersection/s."
	'-----------------------------------use this to use only the first intersection
	ReDim Preserve pts(2)
	ThisDrawing.ModelSpace.AddPoint pts
	'-----------------------------------use this to use only the first intersection
	'-----------------------------------------use this to display all intersections
	'Dim ptCoord(0 To 2) As Double
	'For i = 0 To UBound(pts) Step 3
	'ptCoord(0) = pts(i): ptCoord(1) = pts(i + 1): ptCoord(2) = pts(i + 2)
	'Debug.Print ptCoord(0); ptCoord(1); ptCoord(2)
	'ThisDrawing.ModelSpace.AddPoint ptCoord
	'Next
	'-----------------------------------------use this to display all intersections
End Sub

Function Is2DPointsEqual(p1 As Variant, p2 As Variant, gap As Double) As Boolean
	Is2DPointsEqual = False
	Dim a, b
	a = Abs(CDbl(p1(0)) - CDbl(p2(0)))
	b = Abs(CDbl(p1(1)) - CDbl(p2(1)))
	If a <= gap And b <= gap Then Is2DPointsEqual = True
End Function

Sub JoinLines()
	' based on idea by Norman Yuan
	' Fatty T.O.H. () 2007 * all rights removed
	' edited 02.04.2008
	Dim oSsets As AcadSelectionSets
	Dim pSset As AcadSelectionSet
	Dim oSset As AcadSelectionSet
	Dim setName As String
	Dim fType(0) As Integer
	Dim fData(0) As Variant
	Dim varPt As Variant
	Dim pickPt As Variant
	Dim fLine As AcadLine
	Dim oLine As AcadEntity
	Dim oEnt As AcadEntity
	Dim commStr As String
	Dim stPt(1) As Double
	Dim endPt(1) As Double
	Dim dxftype, dxfcode
	Dim n As Integer
	Dim sp As Variant
	Dim ep As Variant
	Dim ps(1) As Double
	Dim pe(1) As Double
	Dim vexs As Variant
	Dim oSpace As AcadBlock
	With ThisDrawing
	If .ActiveSpace = acModelSpace Then
	Set oSpace = .ModelSpace
	Else
	Set oSpace = .PaperSpace
	End If
	End With
	On Error GoTo Error_Trapp
	Dim osm
	osm = ThisDrawing.GetVariable("OSMODE")
	ThisDrawing.SetVariable "OSMODE", 1
	ThisDrawing.SetVariable "PICKBOX", 1
	pickPt = ThisDrawing.Utility.GetPoint(, vbCr & "Select the starting point of the chain of lines
	:")
	ZoomExtents
	Set oSsets = ThisDrawing.SelectionSets
	fType(0) = 0: fData(0) = "LINE"
	dxftype = fType: dxfcode = fData
	setName = "FirstLine"
	With ThisDrawing.SelectionSets
	While .Count > 0
	.Item(0).Delete
	-1-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_JoinLines_Best.bas Thursday, May 20, 2021 5:34 PM
	Wend
	End With
	setName = "LineSset"
	Set pSset = oSsets.Add("FirstLine")
	pSset.SelectAtPoint pickPt, dxftype, dxfcode
	If pSset.Count > 1 Then
	MsgBox "More than one line selected" & vbCr & _
	"Error"
	Exit Sub
	ElseIf pSset.Count = 1 Then
	Set fLine = pSset.Item(0)
	ElseIf pSset.Count = 0 Then
	MsgBox "Nothing selected" & vbCr & _
	"Error"
	Exit Sub
	End If
	sp = fLine.StartPoint
	ep = fLine.EndPoint
	ps(0) = sp(0): ps(1) = sp(1)
	pe(0) = ep(0): pe(1) = ep(1)
	If Is2DPointsEqual(pickPt, ps, 0.01) Then
	stPt(0) = ps(0): stPt(1) = ps(1)
	endPt(0) = pe(0): endPt(1) = pe(1)
	ElseIf Is2DPointsEqual(pickPt, pe, 0.01) Then
	stPt(0) = pe(0): stPt(1) = pe(1)
	endPt(0) = ps(0): endPt(1) = ps(1)
	End If
	Dim oPline As AcadLWPolyline
	Dim coors(3) As Double
	coors(0) = stPt(0): coors(1) = stPt(1)
	coors(2) = endPt(0): coors(3) = endPt(1)
	Set oPline = oSpace.AddLightWeightPolyline(coors)
	pSset.Delete
	Set pSset = Nothing
	Set oSset = oSsets.Add("LineSset")
	Dim remLine(0) As AcadEntity
	Set remLine(0) = fLine
	oSset.Select acSelectionSetAll, , , dxftype, dxfcode
	oSset.RemoveItems remLine
	fLine.Delete
	Dim i As Long
	i = 1
	Dim Pokey As Boolean
	Pokey = True
	Do Until Not Pokey
	Pokey = False
	Gumby:
	For n = oSset.Count - 1 To 0 Step -1
	Set oLine = oSset.Item(n)
	sp = oLine.StartPoint
	ep = oLine.EndPoint
	ps(0) = sp(0): ps(1) = sp(1)
	pe(0) = ep(0): pe(1) = ep(1)
	If Is2DPointsEqual(ps, endPt, 0.01) Then
	i = i + 1
	oPline.AddVertex i, pe
	Set remLine(0) = oLine
	-2-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_JoinLines_Best.bas Thursday, May 20, 2021 5:34 PM
	oSset.RemoveItems remLine
	oLine.Delete
	vexs = oPline.Coordinate(i)
	endPt(0) = vexs(0): endPt(1) = vexs(1)
	Pokey = True
	Exit For
	ElseIf Is2DPointsEqual(pe, endPt, 0.01) Then
	i = i + 1
	oPline.AddVertex i, ps
	Set remLine(0) = oLine
	oSset.RemoveItems remLine
	oLine.Delete
	vexs = oPline.Coordinate(i)
	endPt(0) = vexs(0): endPt(1) = vexs(1)
	Pokey = True
	Exit For
	End If
	Next n
	If oSset.Count > 0 Then
	GoTo Gumby
	Else
	Exit Do
	End If
	Loop
	oSset.Delete
	Set oSset = Nothing
	Error_Trapp:
	ZoomPrevious
	If Err.Number <> 0 Then
	MsgBox "Error number: " & Err.Number & vbCr & Err.Description
	End If
	On Error Resume Next
	ThisDrawing.SetVariable "OSMODE", osm
	ThisDrawing.SetVariable "PICKBOX", 4 '<--change size to your suit
End Sub


Sub IPNT()
	Dim Poly1 As AcadLWPolyline
	Dim Poly2 As AcadLWPolyline
	Dim pts As Variant
	On Error Resume Next
	ThisDrawing.Utility.GetEntity Poly1, pickPt, "Pick Poly1:"
	If Err Then Exit Sub
	ThisDrawing.Utility.GetEntity Poly2, pickPt, "Pick Poly2:"
	If Err Then Exit Sub
	pts = Poly1.IntersectWith(Poly2, acExtendNone)
	MsgBox "Found " & (UBound(pts) + 1) / 3 & " intersection/s."
	'-----------------------------------use this to use only the first intersection
	ReDim Preserve pts(2)
	ThisDrawing.ModelSpace.AddPoint pts
	'-----------------------------------use this to use only the first intersection
	'-----------------------------------------use this to display all intersections
	'Dim ptCoord(0 To 2) As Double
	'For i = 0 To UBound(pts) Step 3
	'ptCoord(0) = pts(i): ptCoord(1) = pts(i + 1): ptCoord(2) = pts(i + 2)
	'Debug.Print ptCoord(0); ptCoord(1); ptCoord(2)
	'ThisDrawing.ModelSpace.AddPoint ptCoord
	'Next
	'-----------------------------------------use this to display all intersections
End Sub

Sub IPNT()
	Dim objSS As AcadSelectionSet
	Dim objSS2 As AcadSelectionSet
	Dim Poly1 As AcadLWPolyline
	Dim Poly2 As AcadLWPolyline
	On Error Resume Next
	ThisDrawing.SelectionSets("TempSSet1").Delete
	On Error Resume Next
	Set objSS = ThisDrawing.SelectionSets.Add("TempSSet1")
	If Err Then Exit Sub
	MsgBox "Select Poly 1"
	objSS.SelectOnScreen
	For Each Poly1 In objSS
	Exit For: Next
	On Error Resume Next
	ThisDrawing.SelectionSets("TempSSet2").Delete
	On Error Resume Next
	Set objSS2 = ThisDrawing.SelectionSets.Add("TempSSet2")
	If Err Then Exit Sub
	MsgBox "Select Poly 2"
	objSS2.SelectOnScreen
	For Each Poly2 In objSS2
	Exit For: Next
	pts = Poly1.IntersectWith(Poly2, acExtendNone)
	MsgBox "X= " & pts(0) & vbCr & "Y= " & pts(1), vbInformation, "Intersection Point"
End Sub


Private Sub cmdTest_Click()
	Dim Points(5) As Double
	Dim PLine As AcadLWPolyline
	Dim SelectionSet As AcadSelectionSet
	Points(0) = 20: Points(1) = 20: Points(2) = 20
	Points(3) = 20: Points(4) = 20
	Set AcadPolyline = ThisDrawing.ModelSpace.AddLightWeightPolyline(Points)
	ZoomAll
End Sub

Option Explicit
Sub DummyTable()
	Dim minp As Variant
	Dim maxp As Variant
	Dim pt(2) As Double
	Dim acTable As AcadTable
	Dim pickPt As Variant
	pickPt = ThisDrawing.Utility.GetPoint(, vbLf & "Enter a table inserion point: ")
	pt(0) = pickPt(0): pt(1) = pickPt(1): pt(2) = 0#
	Set acTable = ThisDrawing.ActiveLayout.Block.AddTable(pt, 10, 4, 10, 60)
	With acTable
	.RegenerateTableSuppressed = True
	.RecomputeTableBlock False
	.TitleSuppressed = False
	.HeaderSuppressed = False
	.SetTextStyle AcRowType.acTitleRow, "Standard" '<-- change text style name here
	.SetTextStyle AcRowType.acHeaderRow, "Standard"
	.SetTextStyle AcRowType.acDataRow, "Standard"
	Dim i As Double, j As Double
	Dim col As New AcadAcCmColor
	col.SetRGB 255, 0, 255
	'title
	.SetCellTextHeight i, j, 6.4
	.SetCellAlignment i, j, acMiddleCenter
	col.SetRGB 194, 212, 235
	.SetCellBackgroundColor i, j, col
	col.SetRGB 127, 0, 0
	.SetCellContentColor i, j, col
	.SetCellType i, j, acTextCell
	.SetText 0, 0, "Table Title"
	i = i + 1
	'headers
	For j = 0 To .Columns - 1
	col.SetRGB 203, 220, 183
	.SetCellBackgroundColor i, j, col
	col.SetRGB 0, 0, 255
	.SetCellContentColor i, j, col
	.SetCellTextHeight i, j, 5.2
	.SetCellAlignment i, j, acMiddleCenter
	.SetCellType i, j, acTextCell
	.SetText i, j, "Header" & CStr(j + 1)
	Next
	'data rows
	For i = 2 To .Rows - 1
	For j = 0 To .Columns - 1
	.SetCellTextHeight i, j, 4.5
	.SetCellAlignment i, j, acMiddleCenter
	If i Mod 2 = 0 Then
	col.SetRGB 239, 235, 195
	Else
	col.SetRGB 227, 227, 227
	End If
	.SetCellBackgroundColor i, j, col
	col.SetRGB 0, 76, 0
	.SetCellContentColor i, j, col
	.SetCellType i, j, acTextCell
	-1-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_Table_on_CAD.bas Thursday, May 20, 2021 5:32 PM
	.SetText i, j, "Row" & CStr(i - 1) & " Column" & CStr(j + 1)
	Next j
	Next i
	' set color for grid
	col.SetRGB 0, 127, 255
	.SetGridColor AcGridLineType.acHorzBottom, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acHorzInside, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acHorzTop, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acVertInside, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acVertLeft, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acVertRight, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acHorzBottom, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acHorzInside, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acHorzTop, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acVertInside, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acVertLeft, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acVertRight, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acHorzBottom, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acHorzInside, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acHorzTop, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acVertInside, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acVertLeft, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acVertRight, AcRowType.acDataRow, col
	'set line weights for grid
	.SetGridLineWeight AcGridLineType.acVertLeft, AcRowType.acTitleRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertRight, AcRowType.acTitleRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzTop, AcRowType.acTitleRow, AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzBottom, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzInside, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzTop, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertInside, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertLeft, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertRight, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzBottom, AcRowType.acDataRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzInside, AcRowType.acDataRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzTop, AcRowType.acDataRow, AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertInside, AcRowType.acDataRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertLeft, AcRowType.acDataRow, AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertRight, AcRowType.acDataRow,
	AcLineWeight.acLnWt040
	.RegenerateTableSuppressed = False
	.RecomputeTableBlock True
	.Update
	.GetBoundingBox minp, maxp
	ZoomWindow minp, maxp
	ZoomScaled 0.9, acZoomScaledRelative
	ThisDrawing.SetVariable "LWDISPLAY", 1
	Set col = Nothing
	-2-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_Table_on_CAD.bas Thursday, May 20, 2021 5:32 PM
	End With
	MsgBox "Done"
End Sub

Option Explicit
Sub DummyTable()
	Dim minp As Variant
	Dim maxp As Variant
	Dim pt(2) As Double
	Dim acTable As AcadTable
	Dim pickPt As Variant
	pickPt = ThisDrawing.Utility.GetPoint(, vbLf & "Enter a table inserion point: ")
	pt(0) = pickPt(0): pt(1) = pickPt(1): pt(2) = 0#
	Set acTable = ThisDrawing.ActiveLayout.Block.AddTable(pt, 10, 4, 10, 60)
	With acTable
	.RegenerateTableSuppressed = True
	.RecomputeTableBlock False
	.TitleSuppressed = False
	.HeaderSuppressed = False
	.SetTextStyle AcRowType.acTitleRow, "Standard" '<-- change text style name here
	.SetTextStyle AcRowType.acHeaderRow, "Standard"
	.SetTextStyle AcRowType.acDataRow, "Standard"
	Dim i As Double, j As Double
	Dim col As New AcadAcCmColor
	col.SetRGB 255, 0, 255
	'title
	.SetCellTextHeight i, j, 6.4
	.SetCellAlignment i, j, acMiddleCenter
	col.SetRGB 194, 212, 235
	.SetCellBackgroundColor i, j, col
	col.SetRGB 127, 0, 0
	.SetCellContentColor i, j, col
	.SetCellType i, j, acTextCell
	.SetText 0, 0, "Table Title"
	i = i + 1
	'headers
	For j = 0 To .Columns - 1
	col.SetRGB 203, 220, 183
	.SetCellBackgroundColor i, j, col
	col.SetRGB 0, 0, 255
	.SetCellContentColor i, j, col
	.SetCellTextHeight i, j, 5.2
	.SetCellAlignment i, j, acMiddleCenter
	.SetCellType i, j, acTextCell
	.SetText i, j, "Header" & CStr(j + 1)
	Next
	'data rows
	For i = 2 To .Rows - 1
	For j = 0 To .Columns - 1
	.SetCellTextHeight i, j, 4.5
	.SetCellAlignment i, j, acMiddleCenter
	If i Mod 2 = 0 Then
	col.SetRGB 239, 235, 195
	Else
	col.SetRGB 227, 227, 227
	End If
	.SetCellBackgroundColor i, j, col
	col.SetRGB 0, 76, 0
	.SetCellContentColor i, j, col
	.SetCellType i, j, acTextCell
	-1-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_Table_on_CAD2.bas Thursday, May 20, 2021 5:32 PM
	.SetText i, j, "Row" & CStr(i - 1) & " Column" & CStr(j + 1)
	Next j
	Next i
	' set color for grid
	col.SetRGB 0, 127, 255
	.SetGridColor AcGridLineType.acHorzBottom, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acHorzInside, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acHorzTop, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acVertInside, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acVertLeft, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acVertRight, AcRowType.acTitleRow, col
	.SetGridColor AcGridLineType.acHorzBottom, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acHorzInside, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acHorzTop, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acVertInside, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acVertLeft, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acVertRight, AcRowType.acHeaderRow, col
	.SetGridColor AcGridLineType.acHorzBottom, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acHorzInside, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acHorzTop, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acVertInside, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acVertLeft, AcRowType.acDataRow, col
	.SetGridColor AcGridLineType.acVertRight, AcRowType.acDataRow, col
	'set line weights for grid
	.SetGridLineWeight AcGridLineType.acVertLeft, AcRowType.acTitleRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertRight, AcRowType.acTitleRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzTop, AcRowType.acTitleRow, AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzBottom, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzInside, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzTop, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertInside, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertLeft, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertRight, AcRowType.acHeaderRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzBottom, AcRowType.acDataRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzInside, AcRowType.acDataRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acHorzTop, AcRowType.acDataRow, AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertInside, AcRowType.acDataRow,
	AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertLeft, AcRowType.acDataRow, AcLineWeight.acLnWt040
	.SetGridLineWeight AcGridLineType.acVertRight, AcRowType.acDataRow,
	AcLineWeight.acLnWt040
	.RegenerateTableSuppressed = False
	.RecomputeTableBlock True
	.Update
	.GetBoundingBox minp, maxp
	ZoomWindow minp, maxp
	ZoomScaled 0.9, acZoomScaledRelative
	ThisDrawing.SetVariable "LWDISPLAY", 1
	Set col = Nothing
	-2-D:\_S\Excel workbook\CAD_Col\07.UtilityObjects\mk_Table_on_CAD2.bas Thursday, May 20, 2021 5:32 PM
	End With
	MsgBox "Done"
End Sub

Public Sub TestAngleFromXAxis()
	Dim varStart As Variant
	Dim varEnd As Variant
	Dim dblAngle As Double
	With ThisDrawing.Utility
	varStart = .GetPoint(, vbCr & "Pick the start point: ")
	varEnd = .GetPoint(varStart, vbCr & "Pick the end point: ")
	dblAngle = .AngleFromXAxis(varStart, varEnd)
	.Prompt vbCr & "The angle from the X-axis is " _
	& .AngleToString(dblAngle, acDegrees, 2) & " degrees"
	End With
End Sub

Public Sub TestAngleToReal()
	Dim strInput As String
	Dim dblAngle As Double
	With ThisDrawing.Utility
	strInput = .GetString(True, vbCr & "Enter an angle: ")
	dblAngle = .AngleToReal(strInput, acDegrees)
	.Prompt vbCr & "Radians: " & dblAngle
	End With
End Sub

Public Sub TestAngleToString()
	Dim strInput As String
	Dim strOutput As String
	Dim dblAngle As Double
	With ThisDrawing.Utility
	strInput = .GetString(True, vbCr & "Enter an angle: ")
	dblAngle = .AngleToReal(strInput, acDegrees)
	.Prompt vbCr & "Radians: " & dblAngle
	strOutput = .AngleToString(dblAngle, acDegrees, 4)
	.Prompt vbCrLf & "Degrees: " & strOutput
	End With
End Sub

Public Sub TestDistanceToReal()
	Dim strInput As String
	Dim dblDist As Double
	With ThisDrawing.Utility
	strInput = .GetString(True, vbCr & "Enter a distance: ")
	dblDist = .DistanceToReal(strInput, acArchitectural)
	.Prompt vbCr & "Distance: " & dblDist
	End With


Public Sub TestRealToString()
	Dim strInput As String
	Dim strOutput As String
	Dim dblDist As Double
	With ThisDrawing.Utility
	strInput = .GetString(True, vbCr & "Enter a distance: ")
	dblDist = .DistanceToReal(strInput, acArchitectural)
	.Prompt vbCr & "Double: " & dblDist
	strOutput = .RealToString(dblDist, acArchitectural, 4)
	.Prompt vbCrLf & "Distance: " & strOutput
	End With
End Sub

Public Sub TestGetAngle()
	Dim dblInput As Double
	ThisDrawing.SetVariable "DIMAUNIT", acDegrees
	With ThisDrawing.Utility
	dblInput = .GetAngle(, vbCr & "Enter an angle: ")
	.Prompt vbCr & "Angle in radians: " & dblInput
	End With
End Sub

Public Sub TestGetCorner()
	Dim varBase As Variant
	Dim varPick As Variant
	With ThisDrawing.Utility
	varBase = .GetPoint(, vbCr & "Pick the first corner: ")
	.Prompt vbCrLf & varBase(0) & "," & varBase(1)
	varPick = .GetCorner(varBase, vbLf & "Pick the second: ")
	.Prompt vbCr & varPick(0) & "," & varPick(1)
	End With
End Sub

Public Sub TestGetDistance()
	Dim dblInput As Double
	Dim dblBase(2) As Double
	dblBase(0) = 0: dblBase(1) = 0: dblBase(2) = 0
	With ThisDrawing.Utility
	dblInput = .GetDistance(dblBase, vbCr & "Enter a distance: ")
	.Prompt vbCr & "You entered " & dblInput
	End With
End Sub

Public Sub TestGetEntity()
	Dim objEnt As AcadEntity
	Dim varPick As Variant
	On Error Resume Next
	With ThisDrawing.Utility
	.GetEntity objEnt, varPick, vbCr & "Pick an entity: "
	If objEnt Is Nothing Then 'check if object was picked.
	.Prompt vbCrLf & "You did not pick as entity"
	Exit Sub
	End If
	.Prompt vbCr & "You picked a " & objEnt.ObjectName
	.Prompt vbCrLf & "At " & varPick(0) & "," & varPick(1)
	End With
End Sub

Public Sub TestGetInput()
	Dim intInput As Integer
	Dim strInput As String
	On Error Resume Next ' handle exceptions inline
	With ThisDrawing.Utility
	strInput = .GetInput()
	.InitializeUserInput 0, "Line Arc Circle"
	intInput = .GetInteger(vbCr & "Integer or [Line/Arc/Circle]: ")
	If Err.Description Like "*error*" Then
	.Prompt vbCr & "Input Cancelled"
	ElseIf Err.Description Like "*keyword*" Then
	strInput = .GetInput()
	Select Case strInput
	Case "Line": ThisDrawing.SendCommand "_Line" & vbCr
	Case "Arc": ThisDrawing.SendCommand "_Arc" & vbCr
	Case "Circle": ThisDrawing.SendCommand "_Circle" & vbCr
	Case Else: .Prompt vbCr & "Null Input entered"
	End Select
	Else
	.Prompt vbCr & "You entered " & intInput
	End If
	End With
End Sub

Public Sub TestGetInputBug()
	On Error Resume Next ' handle exceptions inline
	With ThisDrawing.Utility
	'' first keyword input
	.InitializeUserInput 1, "Alpha Beta Ship"
	.GetInteger vbCr & "Option [Alpha/Beta/Ship]: "
	MsgBox "You entered: " & .GetInput()
	'' second keyword input - hit Enter here
	.InitializeUserInput 0, "Bug May Slip"
	.GetInteger vbCr & "Hit enter [Bug/May/Slip]: "
	MsgBox "GetInput still returns: " & .GetInput()
	End With
End Sub

Public Sub TestGetInputWorkaround()
	Dim strBeforeKeyword As String
	Dim strKeyword As String
	On Error Resume Next ' handle exceptions inline
	With ThisDrawing.Utility
	'' first keyword input
	.InitializeUserInput 1, "This Bug Stuff"
	.GetInteger vbCrLf & "Option [This/Bug/Stuff]: "
	MsgBox "You entered: " & .GetInput()
	'' get lingering keyword
	strBeforeKeyword = .GetInput()
	'' second keyword input - press Enter
	.InitializeUserInput 0, "Make Life Rough"
	.GetInteger vbCrLf & "Hit enter [Make/Life/Rough]: "
	strKeyword = .GetInput()
	'' if input = lingering, it might be null input
	If strKeyword = strBeforeKeyword Then
	MsgBox "Looks like null input: " & strKeyword
	Else
	MsgBox "This time you entered: " & strKeyword
	End If
	End With
End Sub

Public Sub TestGetInteger()
	Dim intInput As Integer
	With ThisDrawing.Utility
	intInput = .GetInteger(vbCr & "Enter an integer: ")
	.Prompt vbCr & "You entered " & intInput
	End With
End Sub

Sub TestGetKeyword()
	Dim strInput As String
	With ThisDrawing.Utility
	.InitializeUserInput 0, "Line Arc Circle"
	strInput = .GetKeyword(vbCr & "Command [Line/Arc/Circle]: ")
	End With
	Select Case strInput
	Case "Line": ThisDrawing.SendCommand "_Line" & vbCr
	Case "Arc": ThisDrawing.SendCommand "_Arc" & vbCr
	Case "Circle": ThisDrawing.SendCommand "_Circle" & vbCr
	Case Else: MsgBox "You pressed Enter."
	End Select
End Sub


Public Sub TestGetPoint()
	Dim varPick As Variant
	With ThisDrawing.Utility
	varPick = .GetPoint(, vbCr & "Pick a point: ")
	.Prompt vbCr & varPick(0) & "," & varPick(1)
	End With
End Sub

Public Sub TestGetReal()
	Dim dblInput As Double
	With ThisDrawing.Utility
	dblInput = .GetReal(vbCrLf & "Enter an real: ")
	.Prompt vbCr & "You entered " & dblInput
	End With
End Sub

Public Sub TestGetString()
	Dim strInput As String
	With ThisDrawing.Utility
	strInput = .GetString(True, vbCr & "Enter a string: ")
	.Prompt vbCr & "You entered '" & strInput & "' "
	End With
End Sub

Public Sub TestGetSubEntity()
	Dim objEnt As AcadEntity
	Dim varPick As Variant
	Dim varMatrix As Variant
	Dim varParents As Variant
	Dim intI As Integer
	Dim intJ As Integer
	Dim varID As Variant
	With ThisDrawing.Utility
	'' get the subentity from the user
	.GetSubEntity objEnt, varPick, varMatrix, varParents, _
	vbCr & "Pick an entity: "
	'' print some information about the entity
	.Prompt vbCr & "You picked a " & objEnt.ObjectName
	.Prompt vbCrLf & "At " & varPick(0) & "," & varPick(1)
	'' dump the varMatrix
	If Not IsEmpty(varMatrix) Then
	.Prompt vbLf & "MCS to WCS Translation varMatrix:"
	'' format varMatrix row
	For intI = 0 To 3
	.Prompt vbLf & "["
	'' format varMatrix column
	For intJ = 0 To 3
	.Prompt "(" & varMatrix(intI, intJ) & ")"
	Next intJ
	.Prompt "]"
	Next intI
	.Prompt vbLf
	End If
	'' if it has a parent nest
	If Not IsEmpty(varParents) Then
	.Prompt vbLf & "Block nesting:"
	'' depth counter
	intI = -1
	'' traverse most to least deep (reverse order)
	For intJ = UBound(varParents) To LBound(varParents) Step -1
	'' increment depth
	intI = intI + 1
	'' indent output
	.Prompt vbLf & Space(intI * 2)
	'' parent object ID
	varID = varParents(intJ)
	'' parent entity
	Set objEnt = ThisDrawing.ObjectIdToObject(varID)
	'' print info about parent
	.Prompt objEnt.ObjectName & " : " & objEnt.Name
	Next intJ
	.Prompt vbLf
	End If
	.Prompt vbCr
	End With
End Sub


Public Sub TestUserInput()
	Dim strInput As String
	With ThisDrawing.Utility
	.InitializeUserInput 1, "Line Arc Circle laSt"
	strInput = .GetKeyword(vbCr & "Option [Line/Arc/Circle/laSt]: ")
	.Prompt "You selected '" & strInput & "'"
	End With
End Sub

Private Sub CommandButton1_Click()
	Dim strInput As String
	'Me.Hide
	With ThisDrawing.Utility
	.InitializeUserInput 1, "Line Arc Circle laSt"
	strInput = .GetKeyword(vbCr & "Option [Line/Arc/Circle/laSt]: ")
	MsgBox "You selected '" & strInput & "'"
	End With
	'Me.Show
End Sub

Public Sub TestPolarPoint()
	Dim varpnt1 As Variant
	Dim varpnt2 As Variant
	Dim varpnt3 As Variant
	Dim varpnt4 As Variant
	Dim dblAngle As Double
	Dim dblLength As Double
	Dim dblHeight As Double
	Dim dbl90Deg As Double
	'' get the point, length, height, and angle from user
	With ThisDrawing.Utility
	'' get point, length, height, and angle from user
	varpnt1 = .GetPoint(, vbCr & "Pick the start point: ")
	dblLength = .GetDistance(varpnt1, vbCr & "Enter the length: ")
	dblHeight = .GetDistance(varpnt1, vbCr & "Enter the height: ")
	dblAngle = .GetAngle(varpnt1, vbCr & "Enter the angle: ")
	'' calculate remaining rectangle points
	dbl90Deg = .AngleToReal("90d", acDegrees)
	varpnt2 = .PolarPoint(varpnt1, dblAngle, dblLength)
	varpnt3 = .PolarPoint(varpnt2, dblAngle + dbl90Deg, dblHeight)
	varpnt4 = .PolarPoint(varpnt3, dblAngle + (dbl90Deg * 2), dblLength)
	End With
	'' draw the rectangle
	With ThisDrawing
	.ModelSpace.AddLine varpnt1, varpnt2
	.ModelSpace.AddLine varpnt2, varpnt3
	.ModelSpace.AddLine varpnt3, varpnt4
	.ModelSpace.AddLine varpnt4, varpnt1
	End With
End Sub

Public Sub TestPrompt()
	ThisDrawing.Utility.Prompt vbCrLf & "This is a simple message"
End Sub

Public Sub TestUserInput()
	Dim strInput As String
	With ThisDrawing.Utility
	.InitializeUserInput 1, "Line Arc Circle laSt"
	strInput = .GetKeyword(vbCr & "Option [Line/Arc/Circle/laSt]: ")
	.Prompt vbCr & "You selected '" & strInput & "'"
	End With
End Sub

Public Sub TestTranslateCoordinates()
	Dim varpnt1 As Variant
	Dim varpnt1Ucs As Variant
	Dim varpnt2 As Variant
	'' get the point, length, height, and angle from user
	With ThisDrawing.Utility
	'' get start point
	varpnt1 = .GetPoint(, vbCr & "Pick the start point: ")
	'' convert to UCS for use in the base point rubber-band line
	varpnt1Ucs = .TranslateCoordinates(varpnt1, acWorld, acUCS, False)
	'' get end point
	varpnt2 = .GetPoint(varpnt1Ucs, vbCr & "Pick the end point: ")
	End With
	'' draw the line
	With ThisDrawing
	.ModelSpace.AddLine varpnt1, varpnt2
	End With
End Sub


Public Sub tryer()
	ThisDrawing.SendCommand "pedit" & vbCr
End Sub


'08.DrawingObject


Public Sub TestAddPolyline()
	Dim objEnt As AcadPolyline
	Dim dblVertices(17) As Double
	'' setup initial points
	dblVertices(0) = 0: dblVertices(1) = 0: dblVertices(2) = 0
	dblVertices(3) = 10: dblVertices(4) = 0: dblVertices(5) = 0
	dblVertices(6) = 7: dblVertices(7) = 10: dblVertices(8) = 0
	dblVertices(9) = 5: dblVertices(10) = 7: dblVertices(11) = 0
	dblVertices(12) = 6: dblVertices(13) = 2: dblVertices(14) = 0
	dblVertices(15) = 0: dblVertices(16) = 4: dblVertices(17) = 0
	'' draw the entity
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddPolyline(dblVertices)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddPolyline(dblVertices)
	End If
	objEnt.Type = acFitCurvePoly
	objEnt.Closed = True
	objEnt.Update
End Sub

Public Sub TestPolylineType()
	Dim objEnt As AcadPolyline
	Dim varPick As Variant
	Dim strType As String
	Dim intType As Integer
	On Error Resume Next
	With ThisDrawing.Utility
	.GetEntity objEnt, varPick, vbCr & "Pick a polyline: "
	If Err Then
	MsgBox "That is not a Polyline"
	Exit Sub
	End If
	.InitializeUserInput 1, "Simple Fit Quad Cubic"
	strType = .GetKeyword(vbCr & "Change type [Simple/Fit/Quad/Cubic]: ")
	Select Case strType
	Case "Simple": intType = acSimplePoly
	Case "Fit": intType = acFitCurvePoly
	Case "Quad": intType = acQuadSplinePoly
	Case "Cubic": intType = acCubicSplinePoly
	End Select
	End With
	objEnt.Type = intType
	objEnt.Closed = True
	objEnt.Update
End Sub

Public Sub TestAddArc()
	Dim varCenter As Variant
	Dim dblRadius As Double
	Dim dblStart As Double
	Dim dblEnd As Double
	Dim objEnt As AcadArc
	On Error Resume Next
	'' get input from user
	With ThisDrawing.Utility
	varCenter = .GetPoint(, vbCr & "Pick the center point: ")
	dblRadius = .GetDistance(varCenter, vbCr & "Enter the radius: ")
	dblStart = .GetAngle(varCenter, vbCr & "Enter the start angle: ")
	dblEnd = .GetAngle(varCenter, vbCr & "Enter the end angle: ")
	End With
	'' draw the arc
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddArc(varCenter, dblRadius, _
	dblStart, dblEnd)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddArc(varCenter, dblRadius, dblStart, dblEnd)
	End If
	objEnt.Update
End Sub

Public Sub TestAddCircle()
	Dim varCenter As Variant
	Dim dblRadius As Double
	Dim objEnt As AcadCircle
	On Error Resume Next
	'' get input from user
	With ThisDrawing.Utility
	varCenter = .GetPoint(, vbCr & "Pick the centerpoint: ")
	dblRadius = .GetDistance(varCenter, vbCr & "Enter the radius: ")
	End With
	'' draw the entity
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddCircle(varCenter, dblRadius)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddCircle(varCenter, dblRadius)
	End If
	objEnt.Update
End Sub

Option Explicit
Option Private Module
Sub DrawPolyline()
	'Draws a polyline in AutoCAD using X and Y coordinates from sheet Coordinates.
	'By Christos Samaras
	'https://myengineeringworld.net/////
	'In order to use the macro you must enable the AutoCAD library from VBA editor:
	'Go to Tools -> References -> Autocad xxxx Type Library, where xxxx depends
	'on your AutoCAD version (i.e. 2010, 2011, 2012, etc.) you have installed to your PC.
	'Declaring the necessary variables.
	Dim acadApp As AcadApplication
	Dim acadDoc As AcadDocument
	Dim LastRow As Long
	Dim acadPol As AcadLWPolyline
	Dim dblCoordinates() As Double
	Dim i As Long
	Dim j As Long
	Dim k As Long
	Sheet6.Activate
	'Find the last row.
	With Sheet6
	LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
	End With
	'Check if there are at least two points.
	If LastRow < 3 Then
	MsgBox "There not enough points to draw the polyline!", vbCritical, "Points Error"
	Exit Sub
	End If
	'Check if AutoCAD is open.
	On Error Resume Next
	Set acadApp = GetObject(, "AutoCAD.Application")
	On Error GoTo 0
	'If AutoCAD is not opened create a new instance and make it visible.
	If acadApp Is Nothing Then
	Set acadApp = New AcadApplication
	acadApp.Visible = True
	End If
	'Check if there is an active drawing.
	On Error Resume Next
	Set acadDoc = acadApp.ActiveDocument
	On Error GoTo 0
	'No active drawing found. Create a new one.
	If acadDoc Is Nothing Then
	Set acadDoc = acadApp.Documents.Add
	acadApp.Visible = True
	End If
	'Get the array size.
	ReDim dblCoordinates(2 * (LastRow - 1) - 1)
	'Pass the coordinates to array.
	k = 0
	For i = 2 To LastRow
	-1-D:\_S\Excel workbook\CAD_Col\08.DrawingObject\_drawPolylne_Excel.bas Thursday, May 20, 2021 7:09 PM
	For j = 1 To 2
	dblCoordinates(k) = Sheet6.Cells(i, j)
	k = k + 1
	Next j
	Next i
	'Draw the polyline either at model space or at paper space.
	If acadDoc.ActiveSpace = acModelSpace Then
	Set acadPol = acadDoc.ModelSpace.AddLightWeightPolyline(dblCoordinates)
	Else
	Set acadPol = acadDoc.PaperSpace.AddLightWeightPolyline(dblCoordinates)
	End If
	'Leave the polyline open (the last point is not connected with the first point).
	'Set the next line to true if you need to connect the last point with the first one.
	acadPol.Closed = False
	acadPol.Update
	'Zooming in to the drawing area.
	acadApp.ZoomExtents
	'Inform the user that the polyline was created.
	MsgBox "The polyline was successfully created!", vbInformation, "Finished"
End Sub

Option Explicit
Public Sub DoWork()
	Dim i As Integer
	Dim pl As AcadLWPolyline
	Dim ent As AcadEntity
	Dim pt As Variant
	Dim line As AcadLine
	On Error Resume Next
	ThisDrawing.Utility.GetEntity ent, pt, vbCr & "Select polyline:"
	If ent Is Nothing Then Exit Sub
	If TypeOf ent Is AcadLWPolyline Then
	Set pl = ent
	Else
	MsgBox "Selected entity is not polyline!"
	Exit Sub
	End If
	On Error GoTo 0
	Dim startPt(0 To 2) As Double
	Dim endPt(0 To 2) As Double
	For i = 0 To UBound(pl.Coordinates) Step 2
	startPt(0) = pl.Coordinates(i): startPt(1) = pl.Coordinates(i + 1): startPt(2) = 0#
	endPt(0) = pl.Coordinates(i): endPt(1) = 0#: startPt(2) = 0#
	Set line = ThisDrawing.ModelSpace.AddLine(startPt, endPt)
	line.Update
	Next
End Sub

Sub Ch4_EditPolyline()
	Dim plineObj As AcadLWPolyline
	Dim points(0 To 9) As Double
	' Define the 2D polyline points
	points(0) = 1: points(1) = 1
	points(2) = 1: points(3) = 2
	points(4) = 2: points(5) = 2
	points(6) = 3: points(7) = 2
	points(8) = 4: points(9) = 4
	' Create a light weight Polyline object
	Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
	' Add a bulge to segment 3
	plineObj.SetBulge 3, -0.5
	' Define the new vertex
	Dim newVertex(0 To 1) As Double
	newVertex(0) = 4: newVertex(1) = 1
	' Add the vertex to the polyline
	plineObj.AddVertex 5, newVertex
	' Set the width of the new segment
	plineObj.SetWidth 4, 0.1, 0.5
	' Close the polyline
	plineObj.closed = True
	plineObj.Update
End Sub

Public Sub TestAddEllipse()
	Dim dblCenter(0 To 2) As Double
	Dim dblMajor(0 To 2) As Double
	Dim dblRatio As Double
	Dim dblStart As Double
	Dim dblEnd As Double
	Dim objEnt As AcadEllipse
	On Error Resume Next
	'' setup the ellipse parameters
	dblCenter(0) = 0: dblCenter(1) = 0: dblCenter(2) = 0
	dblMajor(0) = 10: dblMajor(1) = 0: dblMajor(2) = 0
	dblRatio = 0.5
	'' draw the ellipse
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddEllipse(dblCenter, dblMajor, dblRatio)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddEllipse(dblCenter, dblMajor, dblRatio)
	End If
	objEnt.Update
	'' get angular input from user
	With ThisDrawing.Utility
	dblStart = .GetAngle(dblCenter, vbCr & "Enter the start angle: ")
	dblEnd = .GetAngle(dblCenter, vbCr & "Enter the end angle: ")
	End With
	'' convert the ellipse into elliptical arc
	With objEnt
	.StartAngle = dblStart
	.EndAngle = dblEnd
	.Update
	End With
End Sub

Public Sub TestAddLine()
	Dim varStart As Variant
	Dim varEnd As Variant
	Dim objEnt As AcadLine
	On Error Resume Next
	'' get input from user
	With ThisDrawing.Utility
	varStart = .GetPoint(, vbCr & "Pick the start point: ")
	varEnd = .GetPoint(varStart, vbCr & "Pick the end point: ")
	End With
	'' draw the entity
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddLine(varStart, varEnd)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddLine(varStart, varEnd)
	End If
	objEnt.Update
End Sub

Public Sub TestAddLWPolyline()
	Dim objEnt As AcadLWPolyline
	Dim dblVertices() As Double
	'' setup initial points
	ReDim dblVertices(11)
	dblVertices(0) = 0#: dblVertices(1) = 0#
	dblVertices(2) = 10#: dblVertices(3) = 0#
	dblVertices(4) = 10#: dblVertices(5) = 10#
	dblVertices(6) = 5#: dblVertices(7) = 5#
	dblVertices(8) = 2#: dblVertices(9) = 2#
	dblVertices(10) = 0#: dblVertices(11) = 10#
	'' draw the entity
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddLightWeightPolyline(dblVertices)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddLightWeightPolyline(dblVertices)
	End If
	objEnt.Closed = True
	objEnt.Update
End Sub

Public Sub TestAddVertex()
	On Error Resume Next
	Dim objEnt As AcadLWPolyline
	Dim dblNew(0 To 1) As Double
	Dim lngLastVertex As Long
	Dim varPick As Variant
	Dim varWCS As Variant
	With ThisDrawing.Utility
	'' get entity from user
	.GetEntity objEnt, varPick, vbCr & "Pick a polyline <exit>: "
	'' exit if no pick
	If objEnt Is Nothing Then Exit Sub
	'' exit if not a lwpolyline
	If objEnt.ObjectName <> "AcDbPolyline" Then
	MsgBox "You did not pick a polyline"
	Exit Sub
	End If
	'' copy last vertex of pline into pickpoint to begin loop
	ReDim varPick(2)
	varPick(0) = objEnt.Coordinates(UBound(objEnt.Coordinates) - 1)
	varPick(1) = objEnt.Coordinates(UBound(objEnt.Coordinates))
	varPick(2) = 0
	'' append vertexes in a loop
	Do
	'' translate picked point to UCS for basepoint below
	varWCS = .TranslateCoordinates(varPick, acWorld, acUCS, True)
	'' get user point for new vertex, use last pick as basepoint
	varPick = .GetPoint(varWCS, vbCr & "Pick another point <exit>: ")
	'' exit loop if no point picked
	If Err Then Exit Do
	'' copy picked point X and Y into new 2d point
	dblNew(0) = varPick(0): dblNew(1) = varPick(1)
	'' get last vertex offset. it is one half the array size
	lngLastVertex = (UBound(objEnt.Coordinates) + 1) / 2
	'' add new vertex to pline at last offset
	objEnt.AddVertex lngLastVertex, dblNew
	Loop
	End With
	objEnt.Update
End Sub

Public Sub TestAddMLine()
	Dim objEnt As AcadMLine
	Dim dblVertices(17) As Double
	'' setup initial points
	dblVertices(0) = 0: dblVertices(1) = 0: dblVertices(2) = 0
	dblVertices(3) = 10: dblVertices(4) = 0: dblVertices(5) = 0
	dblVertices(6) = 10: dblVertices(7) = 10: dblVertices(8) = 0
	dblVertices(9) = 5: dblVertices(10) = 10: dblVertices(11) = 0
	dblVertices(12) = 5: dblVertices(13) = 5: dblVertices(14) = 0
	dblVertices(15) = 0: dblVertices(16) = 5: dblVertices(17) = 0
	'' draw the entity
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddMLine(dblVertices)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddMLine(dblVertices)
	End If
	objEnt.Update
End Sub

Public Sub activeSpaceTG()
	If ThisDrawing.ActiveSpace = acModelSpace Then
	MsgBox "The active space is model space"
	Else
	MsgBox "The active space is paper space"
	End If
	End Sub
Public Sub ToggleSpace()
	With ThisDrawing
	If .ActiveSpace = acModelSpace Then
	.ActiveSpace = acPaperSpace
	Else
	.ActiveSpace = acModelSpace
	End If
	End With
End Sub

Sub Example_Move()
	' This example creates a circle and then performs
	' a move on that circle.
	' Create the circle
	Dim circleObj As AcadCircle
	Dim center(0 To 2) As Double
	Dim radius As Double
	center(0) = 2#: center(1) = 2#: center(2) = 0#
	radius = 0.5
	Set circleObj = ThisDrawing.ModelSpace.AddCircle(center, radius)
	ZoomAll
	' Define the points that make up the move vector
	Dim point1(0 To 2) As Double
	Dim point2(0 To 2) As Double
	point1(0) = 0: point1(1) = 0: point1(2) = 0
	point2(0) = 2: point2(1) = 0: point2(2) = 0
	MsgBox "Move the circle 2 units in the X direction.", , "Move Example"
	' Move the circle
	circleObj.Move point1, point2
	ZoomAll
	MsgBox "Move completed.", , "Move Example"
End Sub

Public Sub TestAddMText()
	Dim varStart As Variant
	Dim dblWidth As Double
	Dim strText As String
	Dim objEnt As AcadMText
	On Error Resume Next
	'' get input from user
	With ThisDrawing.Utility
	varStart = .GetPoint(, vbCr & "Pick the start point: ")
	dblWidth = .GetDistance(varStart, vbCr & "Indicate the width: ")
	strText = .GetString(True, vbCr & "Enter the text: ")
	End With
	'' add font and size formatting
	strText = "\Fromand.shx;\H0.5;" & strText
	'' create the mtext
	Set objEnt = ThisDrawing.ModelSpace.AddMText(varStart, dblWidth, strText)
	objEnt.Update
End Sub


Public Sub TestAddPoint()
	Dim objEnt As AcadPoint
	Dim varPick As Variant
	Dim strType As String
	Dim intType As Integer
	Dim dblSize As Double
	On Error Resume Next
	With ThisDrawing.Utility
	'' get the pdmode center type
	.InitializeUserInput 1, "Dot None Cross X Tick"
	strType = .GetKeyword(vbCr & "Center type [Dot/None/Cross/X/Tick]: ")
	If Err Then Exit Sub
	Select Case strType
	Case "Dot": intType = 0
	Case "None": intType = 1
	Case "Cross": intType = 2
	Case "X": intType = 3
	Case "Tick": intType = 4
	End Select
	'' get the pdmode surrounding type
	.InitializeUserInput 1, "Circle Square Both"
	strType = .GetKeyword(vbCr & "Outer type [Circle/Square/Both]: ")
	If Err Then Exit Sub
	Select Case strType
	Case "Circle": intType = intType + 32
	Case "Square": intType = intType + 64
	Case "Both": intType = intType + 96
	End Select
	'' get the pdsize
	.InitializeUserInput 1, ""
	dblSize = .GetDistance(, vbCr & "Enter a point size: ")
	If Err Then Exit Sub
	'' set the system varibles
	With ThisDrawing
	.SetVariable "PDMODE", intType
	.SetVariable "PDSIZE", dblSize
	End With
	'' now add points in a loop
	Do
	'' get user point for new vertex, use last pick as basepoint
	varPick = .GetPoint(, vbCr & "Pick a point <exit>: ")
	'' exit loop if no point picked
	If Err Then Exit Do
	'' add new vertex to pline at last offset
	ThisDrawing.ModelSpace.AddPoint varPick
	Loop
	End With
End Sub

Public Sub TestAddBulge()
	Dim objEnt As AcadPolyline
	Dim dblVertices(17) As Double
	'' setup initial points
	dblVertices(0) = 0: dblVertices(1) = 0: dblVertices(2) = 0
	dblVertices(3) = 10: dblVertices(4) = 0: dblVertices(5) = 0
	dblVertices(6) = 7: dblVertices(7) = 10: dblVertices(8) = 0
	dblVertices(9) = 5: dblVertices(10) = 7: dblVertices(11) = 0
	dblVertices(12) = 6: dblVertices(13) = 2: dblVertices(14) = 0
	dblVertices(15) = 0: dblVertices(16) = 4: dblVertices(17) = 0
	'' draw the entity
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddPolyline(dblVertices)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddPolyline(dblVertices)
	End If
	objEnt.Type = acSimplePoly
	'add bulge to the fourth segment
	objEnt.SetBulge 3, 0.5
	objEnt.Update
End Sub

Public Sub TestAddRay()
	Dim varStart As Variant
	Dim varEnd As Variant
	Dim objEnt As AcadRay
	On Error Resume Next
	'' get input from user
	With ThisDrawing.Utility
	varStart = .GetPoint(, vbCr & "Pick the start point: ")
	varEnd = .GetPoint(varStart, vbCr & "Indicate a direction: ")
	End With
	'' draw the entity
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddRay(varStart, varEnd)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddRay(varStart, varEnd)
	End If
	objEnt.Update
End Sub

Public Sub TestAddRegion()
	Dim varCenter As Variant
	Dim varMove As Variant
	Dim dblRadius As Double
	Dim dblAngle As Double
	Dim varRegions As Variant
	Dim objEnts() As AcadEntity
	On Error Resume Next
	'' get input from user
	With ThisDrawing.Utility
	varCenter = .GetPoint(, vbCr & "Pick the center point: ")
	dblRadius = .GetDistance(varCenter, vbCr & "Indicate the radius: ")
	dblAngle = .AngleToReal("180", acDegrees)
	End With
	'' draw the entities
	With ThisDrawing.ModelSpace
	'' draw the outer region (circle)
	ReDim objEnts(2)
	Set objEnts(0) = .AddCircle(varCenter, dblRadius)
	'' draw the inner region (semicircle)
	Set objEnts(1) = .AddArc(varCenter, dblRadius * 0.5, 0, dblAngle)
	Set objEnts(2) = .AddLine(objEnts(1).StartPoint, objEnts(1).EndPoint)
	'' create the regions
	varRegions = .AddRegion(objEnts)
	End With
	'' get new position from user
	varMove = ThisDrawing.Utility.GetPoint(varCenter, vbCr & _
	"Pick a new location: ")
	'' subtract the inner region from the outer
	varRegions(1).Boolean acSubtraction, varRegions(0)
	'' move the composite region to a new location
	varRegions(1).Move varCenter, varMove
End Sub

Public Sub TestAddSolid()
	Dim varP1 As Variant
	Dim varP2 As Variant
	Dim varP3 As Variant
	Dim varP4 As Variant
	Dim objEnt As AcadSolid
	On Error Resume Next
	'' ensure that solid fill is enabled
	ThisDrawing.SetVariable "FILLMODE", 1
	'' get input from user
	With ThisDrawing.Utility
	varP1 = .GetPoint(, vbCr & "Pick the start point: ")
	varP2 = .GetPoint(varP1, vbCr & "Pick the second point: ")
	varP3 = .GetPoint(varP1, vbCr & "Pick a point opposite the start: ")
	varP4 = .GetPoint(varP3, vbCr & "Pick the last point: ")
	End With
	'' draw the entity
	Set objEnt = ThisDrawing.ModelSpace.AddSolid(varP1, varP2, varP3, varP4)
	objEnt.Update
End Sub

Public Sub TestAddSpline()
	Dim objEnt As AcadSpline
	Dim dblBegin(0 To 2) As Double
	Dim dblEnd(0 To 2) As Double
	Dim dblPoints(14) As Double
	'' set tangencies
	dblBegin(0) = 1.5: dblBegin(1) = 0#: dblBegin(2) = 0
	dblEnd(0) = 1.5: dblEnd(1) = 0#: dblEnd(2) = 0
	'' set the fit dblPoints
	dblPoints(0) = 0: dblPoints(1) = 0: dblPoints(2) = 0
	dblPoints(3) = 3: dblPoints(4) = 5: dblPoints(5) = 0
	dblPoints(6) = 5: dblPoints(7) = 0: dblPoints(8) = 0
	dblPoints(9) = 7: dblPoints(10) = -5: dblPoints(11) = 0
	dblPoints(12) = 10: dblPoints(13) = 0: dblPoints(14) = 0
	'' draw the entity
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddSpline(dblPoints, dblBegin, dblEnd)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddSpline(dblPoints, dblBegin, dblEnd)
	End If
	objEnt.Update
End Sub

Public Sub TestAddText()
	Dim varStart As Variant
	Dim dblHeight As Double
	Dim strText As String
	Dim objEnt As AcadText
	On Error Resume Next
	'' get input from user
	With ThisDrawing.Utility
	varStart = .GetPoint(, vbCr & "Pick the start point: ")
	dblHeight = .GetDistance(varStart, vbCr & "Indicate the height: ")
	strText = .GetString(True, vbCr & "Enter the text: ")
	End With
	'' create the text
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddText(strText, varStart, dblHeight)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddText(strText, varStart, dblHeight)
	End If
	objEnt.Update
End Sub

Public Sub TestAddXline()
	Dim varStart As Variant
	Dim varEnd As Variant
	Dim objEnt As AcadXline
	On Error Resume Next
	'' get input from user
	With ThisDrawing.Utility
	varStart = .GetPoint(, vbCr & "Pick the start point: ")
	varEnd = .GetPoint(varStart, vbCr & "Indicate an angle: ")
	End With
	'' draw the entity
	If ThisDrawing.ActiveSpace = acModelSpace Then
	Set objEnt = ThisDrawing.ModelSpace.AddXline(varStart, varEnd)
	Else
	Set objEnt = ThisDrawing.PaperSpace.AddXline(varStart, varEnd)
	End If
	objEnt.Update
End Sub


'09.3DObjects

Public Sub TestMassProperties()
	Dim objEnt As Acad3DSolid
	Dim varPick As Variant
	Dim strMassProperties As String
	Dim varProperty As Variant
	Dim intI As Integer
	On Error Resume Next
	'' let user pick a solid
	With ThisDrawing.Utility
	.GetEntity objEnt, varPick, vbCr & "Pick a solid: "
	If Err Then
	MsgBox "That is not an Acad3DSolid"
	Exit Sub
	End If
	End With
	'' format mass properties
	With objEnt
	strMassProperties = "Volume: "
	strMassProperties = strMassProperties & vbCr & " " & .Volume
	strMassProperties = strMassProperties & vbCr & vbCr & _
	"Center Of Gravity: "
	For Each varProperty In .Centroid
	strMassProperties = strMassProperties & vbCr & " " _
	& varProperty
	Next
	strMassProperties = strMassProperties & vbCr & vbCr & _
	"Moment Of Inertia: "
	For Each varProperty In .MomentOfInertia
	strMassProperties = strMassProperties & vbCr & " " & _
	varProperty
	Next
	strMassProperties = strMassProperties & vbCr & vbCr & _
	"Product Of Inertia: "
	For Each varProperty In .ProductOfInertia
	strMassProperties = strMassProperties & vbCr & " " & _
	varProperty
	Next
	strMassProperties = strMassProperties & vbCr & vbCr & _
	"Principal Moments: "
	For Each varProperty In .PrincipalMoments
	strMassProperties = strMassProperties & vbCr & " " & _
	varProperty
	Next
	strMassProperties = strMassProperties & vbCr & vbCr & _
	"Radii Of Gyration: "
	For Each varProperty In .RadiiOfGyration
	strMassProperties = strMassProperties & vbCr & " " & _
	varProperty
	Next
	strMassProperties = strMassProperties & vbCr & vbCr & _
	"Principal Directions: "
	For intI = 0 To UBound(.PrincipalDirections) / 3
	strMassProperties = strMassProperties & vbCr & " (" & _
	.PrincipalDirections((intI - 1) * 3) & ", " & _
	.PrincipalDirections((intI - 1) * 3 + 1) & "," & _
	.PrincipalDirections((intI - 1) * 3 + 2) & ")"
	Next
	End With
	'' highlight entity
	objEnt.Highlight True
	objEnt.Update
	'' display properties
	MsgBox strMassProperties, , "Mass Properties"
	'' dehighlight entity
	-1-D:\_S\Excel workbook\CAD_Col\09.3DObjects\mk9_AnalyzingSolidsMassProperty.bas Thursday, May 20, 2021 7:12 PM
	objEnt.Highlight False
	objEnt.Update
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub

Public Sub TestAddBox()
	Dim varPick As Variant
	Dim dblLength As Double
	Dim dblWidth As Double
	Dim dblHeight As Double
	Dim dblCenter(2) As Double
	Dim objEnt As Acad3DSolid
	'' set the default viewpoint
	SetViewpoint Zoom:=True
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varPick = .GetPoint(, vbCr & "Pick a corner point: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblLength = .GetDistance(varPick, vbCr & "Enter the X length: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblWidth = .GetDistance(varPick, vbCr & "Enter the Y width: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblHeight = .GetDistance(varPick, vbCr & "Enter the Z height: ")
	End With
	'' calculate center point from input
	dblCenter(0) = varPick(0) + (dblLength / 2)
	dblCenter(1) = varPick(1) + (dblWidth / 2)
	dblCenter(2) = varPick(2) + (dblHeight / 2)
	'' draw the entity
	Set objEnt = ThisDrawing.ModelSpace.AddBox(dblCenter, dblLength, _
	dblWidth, dblHeight)
	objEnt.Update
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub

Public Sub TestAddCone()
	Dim varPick As Variant
	Dim dblRadius As Double
	Dim dblHeight As Double
	Dim dblCenter(2) As Double
	Dim objEnt As Acad3DSolid
	'' set the default viewpoint
	SetViewpoint
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varPick = .GetPoint(, vbCr & "Pick the base center point: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblRadius = .GetDistance(varPick, vbCr & "Enter the radius: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblHeight = .GetDistance(varPick, vbCr & "Enter the Z height: ")
	End With
	'' calculate center point from input
	dblCenter(0) = varPick(0)
	dblCenter(1) = varPick(1)
	dblCenter(2) = varPick(2) + (dblHeight / 2)
	'' draw the entity
	Set objEnt = ThisDrawing.ModelSpace.AddCone(dblCenter, dblRadius, _
	dblHeight)
	objEnt.Update
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub

Public Sub TestAddCylinder()
	Dim varPick As Variant
	Dim dblRadius As Double
	Dim dblHeight As Double
	Dim dblCenter(2) As Double
	Dim objEnt As Acad3DSolid
	'' set the default viewpoint
	SetViewpoint
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varPick = .GetPoint(, vbCr & "Pick the base center point: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblRadius = .GetDistance(varPick, vbCr & "Enter the radius: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblHeight = .GetDistance(varPick, vbCr & "Enter the Z height: ")
	End With
	'' calculate center point from input
	dblCenter(0) = varPick(0)
	dblCenter(1) = varPick(1)
	dblCenter(2) = varPick(2) + (dblHeight / 2)
	'' draw the entity
	Set objEnt = ThisDrawing.ModelSpace.AddCylinder(dblCenter, dblRadius, dblHeight)
	objEnt.Update
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestBoolean()
	Dim objFirst As Acad3DSolid
	Dim objSecond As Acad3DSolid
	Dim varPick As Variant
	Dim strOp As String
	On Error Resume Next
	With ThisDrawing.Utility
	'' get first solid from user
	.GetEntity objFirst, varPick, vbCr & "Pick a solid to edit: "
	If Err Then
	MsgBox "That is not an Acad3DSolid"
	Exit Sub
	End If
	'' highlight entity
	objFirst.Highlight True
	objFirst.Update
	'' get second solid from user
	.GetEntity objSecond, varPick, vbCr & "Pick a solid to combine: "
	If Err Then
	MsgBox "That is not an Acad3DSolid"
	Exit Sub
	End If
	'' exit if they're the same
	If objFirst Is objSecond Then
	MsgBox "You must pick 2 different solids"
	Exit Sub
	End If
	'' highlight entity
	objSecond.Highlight True
	objSecond.Update
	'' get boolean operation
	.InitializeUserInput 1, "Intersect Subtract Union"
	strOp = .GetKeyword(vbCr & _
	"Boolean operation [Intersect/Subtract/Union]: ")
	'' combine the solids
	Select Case strOp
	Case "Intersect": objFirst.Boolean acIntersection, objSecond
	Case "Subtract": objFirst.Boolean acSubtraction, objSecond
	Case "Union": objFirst.Boolean acUnion, objSecond
	End Select
	'' highlight entity
	objFirst.Highlight False
	objFirst.Update
	End With
	'' shade the view, and start the interactive orbit command
	ThisDrawing.SendCommand "_shade" & vbCr & "_orbit" & vbCr
End Sub

Public Sub TestInterference()
	Dim objFirst As Acad3DSolid
	Dim objSecond As Acad3DSolid
	Dim objNew As Acad3DSolid
	Dim varPick As Variant
	Dim varNewPnt As Variant
	On Error Resume Next
	'' set default viewpoint
	SetViewpoint
	With ThisDrawing.Utility
	'' get first solid from user
	.GetEntity objFirst, varPick, vbCr & "Pick the first solid: "
	If Err Then
	MsgBox "That is not an Acad3DSolid"
	Exit Sub
	End If
	'' highlight entity
	objFirst.Highlight True
	objFirst.Update
	'' get second solid from user
	.GetEntity objSecond, varPick, vbCr & "Pick the second solid: "
	If Err Then
	MsgBox "That is not an Acad3DSolid"
	Exit Sub
	End If
	'' exit if they're the same
	If objFirst Is objSecond Then
	MsgBox "You must pick 2 different solids"
	Exit Sub
	End If
	'' highlight entity
	objSecond.Highlight True
	objSecond.Update
	'' combine the solids
	Set objNew = objFirst.CheckInterference(objSecond, True)
	If objNew Is Nothing Then
	MsgBox "Those solids don't intersect"
	Else
	'' highlight new solid
	objNew.Highlight True
	objNew.color = acWhite
	objNew.Update
	'' move new solid
	.InitializeUserInput 1
	varNewPnt = .GetPoint(varPick, vbCr & "Pick a new location: ")
	objNew.Move varPick, varNewPnt
	End If
	'' dehighlight entities
	objFirst.Highlight False
	objFirst.Update
	objSecond.Highlight False
	objSecond.Update
	End With
	'' shade the view, and start the interactive orbit command
	ThisDrawing.SendCommand "_shade" & vbCr & "_orbit" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub

Public Sub TestAddEllipticalCone()
	Dim varPick As Variant
	Dim dblXAxis As Double
	Dim dblYAxis As Double
	Dim dblHeight As Double
	Dim dblCenter(2) As Double
	Dim objEnt As Acad3DSolid
	'' set the default viewpoint
	SetViewpoint
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varPick = .GetPoint(, vbCr & "Pick a base center point: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblXAxis = .GetDistance(varPick, vbCr & "Enter the X eccentricity: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblYAxis = .GetDistance(varPick, vbCr & "Enter the Y eccentricity: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblHeight = .GetDistance(varPick, vbCr & "Enter the cone Z height: ")
	End With
	'' calculate center point from input
	dblCenter(0) = varPick(0)
	dblCenter(1) = varPick(1)
	dblCenter(2) = varPick(2) + (dblHeight / 2)
	'' draw the entity
	Set objEnt = ThisDrawing.ModelSpace.AddEllipticalCone(dblCenter, _
	dblXAxis, dblYAxis, dblHeight)
	objEnt.Update
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub
Public Sub TestAddEllipticalCylinder()
	Dim varPick As Variant
	Dim dblXAxis As Double
	Dim dblYAxis As Double
	Dim dblHeight As Double
	Dim dblCenter(2) As Double
	Dim objEnt As Acad3DSolid
	'' set the default viewpoint
	SetViewpoint
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varPick = .GetPoint(, vbCr & "Pick a base center point: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblXAxis = .GetDistance(varPick, vbCr & "Enter the X eccentricity: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblYAxis = .GetDistance(varPick, vbCr & "Enter the Y eccentricity: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblHeight = .GetDistance(varPick, vbCr & _
	"Enter the cylinder Z height: ")
	End With
	'' calculate center point from input
	dblCenter(0) = varPick(0)
	dblCenter(1) = varPick(1)
	dblCenter(2) = varPick(2) + (dblHeight / 2)
	'' draw the entity
	Set objEnt = ThisDrawing.ModelSpace.AddEllipticalCylinder(dblCenter, _
	dblXAxis, dblYAxis, dblHeight)
	objEnt.Update
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub
Public Sub TestAddExtrudedSolid()
	Dim varCenter As Variant
	Dim dblRadius As Double
	Dim dblHeight As Double
	Dim dblTaper As Double
	Dim strInput As String
	Dim varRegions As Variant
	Dim objEnts() As AcadEntity
	Dim objEnt As Acad3DSolid
	Dim varItem As Variant
	On Error GoTo Done
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varCenter = .GetPoint(, vbCr & "Pick the center point: ")
	.InitializeUserInput 1 + 2 + 4
	dblRadius = .GetDistance(varCenter, vbCr & "Indicate the radius: ")
	.InitializeUserInput 1 + 2 + 4
	dblHeight = .GetDistance(varCenter, vbCr & _
	"Enter the extrusion height: ")
	'' get the taper type
	.InitializeUserInput 1, "Expand Contract None"
	strInput = .GetKeyword(vbCr & _
	"Extrusion taper [Expand/Contract/None]: ")
	'' if none, taper = 0
	If strInput = "None" Then
	dblTaper = 0
	'' otherwise, get the taper angle
	Else
	.InitializeUserInput 1 + 2 + 4
	dblTaper = .GetReal("Enter the taper angle ( in degrees): ")
	dblTaper = .AngleToReal(CStr(dblTaper), acDegrees)
	'' if expanding, negate the angle
	If strInput = "Expand" Then dblTaper = -dblTaper
	End If
	End With
	'' draw the entities
	With ThisDrawing.ModelSpace
	'' draw the outer region (circle)
	ReDim objEnts(0)
	Set objEnts(0) = .AddCircle(varCenter, dblRadius)
	'' create the region
	varRegions = .AddRegion(objEnts)
	'' extrude the solid
	Set objEnt = .AddExtrudedSolid(varRegions(0), dblHeight, dblTaper)
	'' update the extruded solid
	objEnt.Update
	End With
	Done:
	If Err Then MsgBox Err.Description
	'' delete the temporary geometry
	For Each varItem In objEnts
	-1-D:\_S\Excel workbook\CAD_Col\09.3DObjects\mk9_ExtrudedSolid.bas Thursday, May 20, 2021 7:12 PM
	varItem.Delete
	Next
	For Each varItem In varRegions
	varItem.Delete
	Next
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub

Public Sub TestAddExtrudedSolidAlongPath()
	Dim objPath As AcadSpline
	Dim varPick As Variant
	Dim intI As Integer
	Dim dblCenter(2) As Double
	Dim dblRadius As Double
	Dim objCircle As AcadCircle
	Dim objEnts() As AcadEntity
	Dim objShape As Acad3DSolid
	Dim varRegions As Variant
	Dim varItem As Variant
	'' set default viewpoint
	SetViewpoint
	'' pick path and calculate shape points
	With ThisDrawing.Utility
	'' pick the path
	On Error Resume Next
	.GetEntity objPath, varPick, "Pick a Spline for the path"
	If Err Then
	MsgBox "You did not pick a spline"
	Exit Sub
	End If
	objPath.color = acGreen
	For intI = 0 To 2
	dblCenter(intI) = objPath.FitPoints(intI)
	Next
	.InitializeUserInput 1 + 2 + 4
	dblRadius = .GetDistance(dblCenter, vbCr & "Indicate the radius: ")
	End With
	'' draw the circular region, then extrude along path
	With ThisDrawing.ModelSpace
	'' draw the outer region (circle)
	ReDim objEnts(0)
	Set objCircle = .AddCircle(dblCenter, dblRadius)
	objCircle.Normal = objPath.StartTangent
	Set objEnts(0) = objCircle
	'' create the region
	varRegions = .AddRegion(objEnts)
	Set objShape = .AddExtrudedSolidAlongPath(varRegions(0), objPath)
	objShape.color = acRed
	End With
	'' delete the temporary geometry
	For Each varItem In objEnts: varItem.Delete: Next
	For Each varItem In varRegions: varItem.Delete: Next
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub
Public Sub TestAddRevolvedSolid()
	Dim objShape As AcadLWPolyline
	Dim varPick As Variant
	Dim objEnt As AcadEntity
	Dim varPnt1 As Variant
	Dim dblOrigin(2) As Double
	Dim varVec As Variant
	Dim dblAngle As Double
	Dim objEnts() As AcadEntity
	Dim varRegions As Variant
	Dim varItem As Variant
	'' set default viewpoint
	SetViewpoint
	'' draw the shape and get rotation from user
	With ThisDrawing.Utility
	'' pick a shape
	On Error Resume Next
	.GetEntity objShape, varPick, "pick a polyline shape"
	If Err Then
	MsgBox "You did not pick the correct type of shape"
	Exit Sub
	End If
	On Error GoTo Done
	objShape.Closed = True
	'' add pline to region input array
	ReDim objEnts(0)
	Set objEnts(0) = objShape
	'' get the axis points
	.InitializeUserInput 1
	varPnt1 = .GetPoint(, vbLf & "Pick an origin of revolution: ")
	.InitializeUserInput 1
	varVec = .GetPoint(dblOrigin, vbLf & _
	"Indicate the axis of revolution: ")
	'' get the angle to revolve
	.InitializeUserInput 1
	dblAngle = .GetAngle(, vbLf & "Angle to revolve: ")
	End With
	'' make the region, then revolve it into a solid
	With ThisDrawing.ModelSpace
	'' make region from closed pline
	varRegions = .AddRegion(objEnts)
	'' revolve solid about axis
	Set objEnt = .AddRevolvedSolid(varRegions(0), varPnt1, varVec, _
	dblAngle)
	objEnt.color = acRed
	End With
	Done:
	If Err Then MsgBox Err.Description
	'' delete the temporary geometry
	For Each varItem In objEnts: varItem.Delete: Next
	If Not IsEmpty(varRegions) Then
	-1-D:\_S\Excel workbook\CAD_Col\09.3DObjects\mk9_RevolvedSolid.bas Thursday, May 20, 2021 7:11 PM
	For Each varItem In varRegions: varItem.Delete: Next
	End If
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestSectionSolid()
	Dim objFirst As Acad3DSolid
	Dim objSecond As Acad3DSolid
	Dim objNew As AcadRegion
	Dim varPick As Variant
	Dim varPnt1 As Variant
	Dim varPnt2 As Variant
	Dim varPnt3 As Variant
	On Error Resume Next
	With ThisDrawing.Utility
	'' get first solid from user
	.GetEntity objFirst, varPick, vbCr & "Pick a solid to section: "
	If Err Then
	MsgBox "That is not an Acad3DSolid"
	Exit Sub
	End If
	'' highlight entity
	objFirst.Highlight True
	objFirst.Update
	.InitializeUserInput 1
	varPnt1 = .GetPoint(varPick, vbCr & "Pick first section point: ")
	.InitializeUserInput 1
	varPnt2 = .GetPoint(varPnt1, vbCr & "Pick second section point: ")
	.InitializeUserInput 1
	varPnt3 = .GetPoint(varPnt2, vbCr & "Pick last section point: ")
	'' section the solid
	Set objNew = objFirst.SectionSolid(varPnt1, varPnt2, varPnt3)
	If objNew Is Nothing Then
	MsgBox "Couldn't section using those points"
	Else
	'' highlight new solid
	objNew.Highlight False
	objNew.color = acWhite
	objNew.Update
	'' move section region to new location
	.InitializeUserInput 1
	varPnt2 = .GetPoint(varPnt1, vbCr & "Pick a new location: ")
	objNew.Move varPnt1, varPnt2
	End If
	'' dehighlight entities
	objFirst.Highlight False
	objFirst.Update
	End With
	'' shade the view
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Public Sub TestSliceSolid()
	Dim objFirst As Acad3DSolid
	Dim objSecond As Acad3DSolid
	Dim objNew As Acad3DSolid
	Dim varPick As Variant
	Dim varPnt1 As Variant
	Dim varPnt2 As Variant
	Dim varPnt3 As Variant
	Dim strOp As String
	Dim blnOp As Boolean
	On Error Resume Next
	With ThisDrawing.Utility
	'' get first solid from user
	.GetEntity objFirst, varPick, vbCr & "Pick a solid to slice: "
	If Err Then
	MsgBox "That is not a 3DSolid"
	Exit Sub
	End If
	'' highlight entity
	objFirst.Highlight True
	objFirst.Update
	.InitializeUserInput 1
	varPnt1 = .GetPoint(varPick, vbCr & "Pick first slice point: ")
	.InitializeUserInput 1
	varPnt2 = .GetPoint(varPnt1, vbCr & "Pick second slice point: ")
	.InitializeUserInput 1
	varPnt3 = .GetPoint(varPnt2, vbCr & "Pick last slice point: ")
	'' section the solid
	Set objNew = objFirst.SliceSolid(varPnt1, varPnt2, varPnt3, True)
	If objNew Is Nothing Then
	MsgBox "Couldn't slice using those points"
	Else
	'' highlight new solid
	objNew.Highlight False
	objNew.color = objNew.color + 1
	objNew.Update
	'' move section region to new location
	.InitializeUserInput 1
	varPnt2 = .GetPoint(varPnt1, vbCr & "Pick a new location: ")
	objNew.Move varPnt1, varPnt2
	End If
	End With
	'' shade the view
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub

Public Sub TestAddSphere()
	Dim varPick As Variant
	Dim dblRadius As Double
	Dim objEnt As Acad3DSolid
	'' set the default viewpoint
	SetViewpoint
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varPick = .GetPoint(, vbCr & "Pick the center point: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblRadius = .GetDistance(varPick, vbCr & "Enter the radius: ")
	End With
	'' draw the entity
	Set objEnt = ThisDrawing.ModelSpace.AddSphere(varPick, dblRadius)
	objEnt.Update
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub
Public Sub TestAddTorus()
	Dim pntPick As Variant
	Dim pntRadius As Variant
	Dim dblRadius As Double
	Dim dblTube As Double
	Dim objEnt As Acad3DSolid
	Dim intI As Integer
	'' set the default viewpoint
	SetViewpoint
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	pntPick = .GetPoint(, vbCr & "Pick the center point: ")
	.InitializeUserInput 1
	pntRadius = .GetPoint(pntPick, vbCr & "Pick a radius point: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblTube = .GetDistance(pntRadius, vbCr & "Enter the tube radius: ")
	End With
	'' calculate radius from points
	For intI = 0 To 2
	dblRadius = dblRadius + (pntPick(intI) - pntRadius(intI)) ^ 2
	Next
	dblRadius = Sqr(dblRadius)
	'' draw the entity
	Set objEnt = ThisDrawing.ModelSpace.AddTorus(pntPick, dblRadius, dblTube)
	objEnt.Update
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

Sub SetViewpoint(Optional Zoom As Boolean = False, _
	Optional X As Double = 1, _
	Optional Y As Double = -2, _
	Optional Z As Double = 1)
	Dim dblDirection(2) As Double
	dblDirection(0) = X: dblDirection(1) = Y: dblDirection(2) = Z
	With ThisDrawing
	.Preferences.ContourLinesPerSurface = 10 ' set surface countours
	.ActiveViewport.Direction = dblDirection ' assign new direction
	.ActiveViewport = .ActiveViewport ' force a viewport update
	If Zoom Then .Application.ZoomAll ' zoomall if requested
	End With
End Sub
Public Sub TestAddWedge()
	Dim varPick As Variant
	Dim dblLength As Double
	Dim dblWidth As Double
	Dim dblHeight As Double
	Dim dblCenter(2) As Double
	Dim objEnt As Acad3DSolid
	'' set the default viewpoint
	SetViewpoint
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varPick = .GetPoint(, vbCr & "Pick a base corner point: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblLength = .GetDistance(varPick, vbCr & "Enter the base X length: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblWidth = .GetDistance(varPick, vbCr & "Enter the base Y width: ")
	.InitializeUserInput 1 + 2 + 4, ""
	dblHeight = .GetDistance(varPick, vbCr & "Enter the base Z height: ")
	End With
	'' calculate center point from input
	dblCenter(0) = varPick(0) + (dblLength / 2)
	dblCenter(1) = varPick(1) + (dblWidth / 2)
	dblCenter(2) = varPick(2) + (dblHeight / 2)
	'' draw the entity
	Set objEnt = ThisDrawing.ModelSpace.AddWedge(dblCenter, dblLength, _
	dblWidth, dblHeight)
	objEnt.Update
	ThisDrawing.SendCommand "_shade" & vbCr
End Sub

'10.EditObjects



Sub Example_AngleFromXAxis()
	' This example finds the angle, in radians, between the X axis
	' and a line defined by two points.
	Dim pt1(0 To 2) As Double
	Dim pt2(0 To 2) As Double
	Dim retAngle As Double
	pt1(0) = 256833.405419725: pt1(1) = 5928142.14654379: pt1(2) = 0
	pt2(0) = 256890.630751811: pt2(1) = 5928014.31962269: pt2(2) = 0
	' Return the angle
	retAngle = ThisDrawing.Utility.AngleFromXAxis(pt1, pt2)
	' Create the line for a visual reference
	Dim lineObj As AcadLine
	Set lineObj = ThisDrawing.ModelSpace.AddLine(pt1, pt2)
	ZoomAll
	' Display the angle found
	MsgBox "The angle in radians between the X axis and the line is " & retAngle, ,
	"AngleFromXAxis Example"
End Sub

Public Sub ChangeLinetype()
	Dim objDrawingObject As AcadEntity
	Dim varEntityPickedPoint As Variant
	On Error Resume Next
	ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
	"Please pick an object"
	If objDrawingObject Is Nothing Then
	MsgBox "You did not choose an object"
	Exit Sub
	End If
	ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
	"Pick an entity to change linetype: "
	objDrawingObject.Linetype = "Continuous"
	objDrawingObject.Update
End Sub

Public Sub ColorGreen()
	Dim objSelectionSet As AcadSelectionSet
	Dim objDrawingObject As AcadEntity
	'choose a selection set name that you only use as temporary storage and
	'ensure that it does not currently exist
	On Error Resume Next
	ThisDrawing.SelectionSets("TempSSet").Delete
	Set objSelectionSet = ThisDrawing.SelectionSets.Add("TempSSet")
	'ask user to pick entities on the screen
	objSelectionSet.SelectOnScreen
	For Each objDrawingObject In objSelectionSet
	objDrawingObject.color = acGreen
	objDrawingObject.Update
	Next
	objSelectionSet.Delete
End Sub

Public Sub CopyObject()
	Dim objDrawingObject As AcadEntity
	Dim objCopiedObject As Object
	Dim varEntityPickedPoint As Variant
	Dim varCopyPoint As Variant
	On Error Resume Next
	ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
	"Pick an entity to copy: "
	If objDrawingObject Is Nothing Then
	MsgBox "You did not pick an object"
	Exit Sub
	End If
	'Copy the object
	Set objCopiedObject = objDrawingObject.Copy()
	varCopyPoint = ThisDrawing.Utility.GetPoint(, "Pick point to copy to: ")
	'put the object in its new position
	objCopiedObject.Move varEntityPickedPoint, varCopyPoint
	objCopiedObject.Update
End Sub

Public Sub Create2DRectangularArray()
	Dim objDrawingObject As AcadEntity
	Dim varEntityPickedPoint As Variant
	Dim lngNoRows As Long
	Dim lngNoColumns As Long
	Dim dblDistRows As Long
	Dim dblDistCols As Long
	Dim varRectangularArray As Variant
	Dim intCount As Integer
	On Error Resume Next
	ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
	"Please pick an entity to form the basis of a rectangular array: "
	If objDrawingObject Is Nothing Then
	MsgBox "You did not choose an object"
	Exit Sub
	End If
	lngNoRows = ThisDrawing.Utility.GetInteger( _
	"Enter the required number of rows: ")
	lngNoColumns = ThisDrawing.Utility.GetInteger( _
	"Enter the required number of columns: ")
	dblDistRows = ThisDrawing.Utility.GetReal( _
	"Enter the required distance between rows: ")
	dblDistCols = ThisDrawing.Utility.GetReal( _
	"Enter the required distance between columns: ")
	varRectangularArray = objDrawingObject.ArrayRectangular(lngNoRows, _
	lngNoColumns, 1, dblDistRows, dblDistCols, 0)
	For intCount = 0 To UBound(varRectangularArray)
	varRectangularArray(intCount).color = acRed
	varRectangularArray(intCount).Update
	Next
End Sub

Public Sub DeleteObject()
	Dim objDrawingObject As AcadEntity
	Dim varEntityPickedPoint As Variant
	On Error Resume Next
	ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
	"Pick an entity to delete: "
	If objDrawingObject Is Nothing Then
	MsgBox "You did not pick an object."
	Exit Sub
	End If
	'delete the object
	objDrawingObject.Delete
End Sub

Public Sub ToggleHighlight()
	Dim objSelectionSet As AcadSelectionSet
	Dim objDrawingObject As AcadEntity
	'choose a selection set name that you only use as temporary storage and
	'ensure that it does not currently exist
	On Error Resume Next
	ThisDrawing.SelectionSets("TempSSet").Delete
	Set objSelectionSet = ThisDrawing.SelectionSets.Add("TempSSet")
	'ask user to pick entities on the screen
	objSelectionSet.SelectOnScreen
	'change the highlight status of each entity selected
	For Each objDrawingObject In objSelectionSet
	objDrawingObject.Highlight True
	objDrawingObject.Update 'not required for 2006
	MsgBox "Notice that the entity is highlighted"
	objDrawingObject.Highlight False 'not required for 2006
	objDrawingObject.Update 'not required for 2006
	MsgBox "Notice that the entity is not highlighted"
	Next
	objSelectionSet.Delete
End Sub

Public Sub MirrorObjects()
	Dim objSelectionSet As AcadSelectionSet
	Dim objDrawingObject As AcadEntity
	Dim objMirroredObject As AcadEntity
	Dim varPoint1 As Variant
	Dim varPoint2 As Variant
	ThisDrawing.SetVariable "MIRRTEXT", 0
	'choose a selection set name that you only use as temporary storage and
	'ensure that it does not currently exist
	On Error Resume Next
	ThisDrawing.SelectionSets("TempSSet").Delete
	Set objSelectionSet = ThisDrawing.SelectionSets.Add("TempSSet")
	'ask user to pick entities on the screen
	ThisDrawing.Utility.Prompt "Pick objects to be mirrored." & vbCrLf
	objSelectionSet.SelectOnScreen
	'change the highlight status of each entity selected
	varPoint1 = ThisDrawing.Utility.GetPoint(, _
	"Select a point on the mirror axis")
	varPoint2 = ThisDrawing.Utility.GetPoint(varPoint1, _
	"Select a point on the mirror axis")
	For Each objDrawingObject In objSelectionSet
	Set objMirroredObject = objDrawingObject.Mirror(varPoint1, varPoint2)
	objMirroredObject.Update
	Next
	objSelectionSet.Delete
	
End Sub

Public Sub MoveObjects()
	Dim varPoint1 As Variant
	Dim varPoint2 As Variant
	Dim objSelectionSet As AcadSelectionSet
	Dim objDrawingObject As AcadEntity
	'choose a selection set name that you only use as temporary storage and
	'ensure that it does not currently exist
	On Error Resume Next
	ThisDrawing.SelectionSets("TempSSet").Delete
	Set objSelectionSet = ThisDrawing.SelectionSets.Add("TempSSet")
	'ask user to pick entities on the screen
	objSelectionSet.SelectOnScreen
	varPoint1 = ThisDrawing.Utility.GetPoint(, vbCrLf _
	& "Base point of displacement: ")
	varPoint2 = ThisDrawing.Utility.GetPoint(varPoint1, vbCrLf _
	& "Second point of displacement: ")
	'move the selection of entities
	For Each objDrawingObject In objSelectionSet
	objDrawingObject.Move varPoint1, varPoint2
	objDrawingObject.Update
	Next
	objSelectionSet.Delete
End Sub

Public Sub CreatePolarArray()
	Dim objDrawingObject As AcadEntity
	Dim varEntityPickedPoint As Variant
	Dim varArrayCenter As Variant
	Dim lngNumberofObjects As Long
	Dim dblAngletoFill As Double
	Dim varPolarArray As Variant
	Dim intCount As Integer
	On Error Resume Next
	ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
	"Please select an entity to form the basis of a polar array"
	If objDrawingObject Is Nothing Then
	MsgBox "You did not choose an object"
	Exit Sub
	End If
	varArrayCenter = ThisDrawing.Utility.GetPoint(, _
	"Pick the center of the array: ")
	lngNumberofObjects = ThisDrawing.Utility.GetInteger( _
	"Enter total number of objects required in the array: ")
	dblAngletoFill = ThisDrawing.Utility.GetReal( _
	"Enter an angle (in degrees less than 360) over which the array should extend: ")
	dblAngletoFill = ThisDrawing.Utility.AngleToReal _
	(CStr(dblAngletoFill), acDegrees)
	If dblAngletoFill > 359 Then
	MsgBox "Angle must be less than 360 degrees", vbCritical
	Exit Sub
	End If
	varPolarArray = objDrawingObject.ArrayPolar(lngNumberofObjects, _
	dblAngletoFill, varArrayCenter)
	For intCount = 0 To UBound(varPolarArray)
	varPolarArray(intCount).color = acRed
	varPolarArray(intCount).Update
	Next
End Sub

Public Sub ToggleVisibility()
	Dim objDrawingObject As AcadEntity
	Dim varEntityPickedPoint As Variant
	On Error Resume Next
	ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
	"Choose an object to toggle visibility: "
	If objDrawingObject Is Nothing Then
	MsgBox "You did not choose an object"
	Exit Sub
	End If
	objDrawingObject.Visible = False
	objDrawingObject.Update
	MsgBox "The object was made invisible!"
	objDrawingObject.Visible = True
	objDrawingObject.Update
	MsgBox "Now it is visible again!"
End Sub

Public Sub OffsetEllipse()
	Dim objEllipse As AcadEllipse
	Dim varObjectArray As Variant
	Dim dblCenter(2) As Double
	Dim dblMajor(2) As Double
	dblMajor(0) = 100#
	Set objEllipse = ThisDrawing.ModelSpace.AddEllipse(dblCenter, dblMajor, _
	0.5)
	varObjectArray = objEllipse.Offset(50)
	MsgBox "The offset object is a " & varObjectArray(0).ObjectName
	varObjectArray = objEllipse.Offset(-25)
	MsgBox "The offset object is a " & varObjectArray(0).ObjectName
End Sub

Public Sub RotateObject()
	Dim objDrawingObject As AcadEntity
	Dim varEntityPickedPoint As Variant
	Dim varBasePoint As Variant
	Dim dblRotationAngle As Double
	On Error Resume Next
	ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
	"Please pick an entity to rotate: "
	If objDrawingObject Is Nothing Then
	MsgBox "You did not choose an object."
	Exit Sub
	End If
	varBasePoint = ThisDrawing.Utility.GetPoint(, _
	"Enter a base point for the rotation.")
	dblRotationAngle = ThisDrawing.Utility.GetReal( _
	"Enter the rotation angle in degrees: ")
	'convert to radians
	dblRotationAngle = ThisDrawing.Utility. _
	AngleToReal(CStr(dblRotationAngle), acDegrees)
	'Rotate the object
	objDrawingObject.Rotate varBasePoint, dblRotationAngle
	objDrawingObject.Update
End Sub

Public Sub ScaleObject()
	Dim objDrawingObject As AcadEntity
	Dim varEntityPickedPoint As Variant
	Dim varBasePoint As Variant
	Dim dblScaleFactor As Double
	On Error Resume Next
	ThisDrawing.Utility.GetEntity objDrawingObject, varEntityPickedPoint, _
	"Please pick an entity to scale: "
	If objDrawingObject Is Nothing Then
	MsgBox "You did not choose an object"
	Exit Sub
	End If
	varBasePoint = ThisDrawing.Utility.GetPoint(, _
	"Pick a base point for the scale:")
	dblScaleFactor = ThisDrawing.Utility.GetReal("Enter the scale factor: ")
	'Scale the object
	objDrawingObject.ScaleEntity varBasePoint, dblScaleFactor
	objDrawingObject.Update
End Sub

'11.Dimensions

Sub Ch5_CopyDimStyles()
	Dim newStyle1 As AcadDimStyle
	Dim newStyle2 As AcadDimStyle
	Dim newStyle3 As AcadDimStyle
	Set newStyle1 = ThisDrawing.DimStyles.Add _
	("Style 1 copied from a dim")
	Call newStyle1.CopyFrom(ThisDrawing.ModelSpace(0))
	Set newStyle2 = ThisDrawing.DimStyles.Add _
	("Style 2 copied from Style 1")
	Call newStyle2.CopyFrom(ThisDrawing.DimStyles.Item _
	("Style 1 copied from a dim"))
	Set newStyle2 = ThisDrawing.DimStyles.Add _
	("Style 3 copied from the running drawing values")
	Call newStyle2.CopyFrom(ThisDrawing)
End Sub

Sub Example_AddDimOrdinate()
	Dim objDim As AcadDimOrdinate
	Dim varDefPt As Variant
	Dim varLdrPt As Variant
	Dim strKeyWord As String
	Dim blnXaxis As Boolean
	' Get the definition point.
	varDefPt = ThisDrawing.Utility.GetPoint(, vbCrLf & "Feature location: ")
	' Get the leader point.
	varLdrPt = ThisDrawing.Utility.GetPoint(varDefPt, vbCrLf & "Leader endpoint: ")
	' Initialize the GetKeyword method with the keywords "X and Y"
	' and disable null input.
	ThisDrawing.Utility.InitializeUserInput 1, "X Y"
	' Determine the axis the user wants to dimension.
	strKeyWord = ThisDrawing.Utility.GetKeyword _
	("X axis or Y axis? [X/Y]: ")
	' Set the Boolean flag appropriately.
	If strKeyWord = "X" Then
	blnXaxis = True
	Else
	blnXaxis = False
	End If
	' Create the ordinate dimension.
	Set objDim = ThisDrawing.ModelSpace.AddDimOrdinate(varDefPt, varLdrPt, _
	blnXaxis)
	' Update the object so it is sure to appear.
	objDim.Update
End Sub

Sub dim_inch()
	' acadDoc.setvariable "DIMSCALE", 1# ' Overall scale factor"
	' acadDoc.setvariable "DIMTXSTY", "Standard" ' Text style"
	acadDoc.SetVariable "DIMLUNIT", 5 ' Linear unit format"
	acadDoc.SetVariable "DIMFRAC", 1 ' Fraction format"
	acadDoc.SetVariable "DIMTXT", 0.125 ' Text height"
	acadDoc.SetVariable "DIMASZ", 0.09375 ' Arrow size"
	acadDoc.SetVariable "DIMCLRT", 7 ' Dimension text color
	acadDoc.SetVariable "DIMTAD", 1 ' Place text above the dimension line"
	acadDoc.SetVariable "DIMTOH", 0 ' Text outside horizontal"
	acadDoc.SetVariable "DIMTIH", 0 ' Text inside extensions is horizontal"
	acadDoc.SetVariable "DIMTOFL", 1 ' Force line inside extension lines"
	acadDoc.SetVariable "DIMTIX", 1 ' Place text inside extensions"
	acadDoc.SetVariable "DIMTMOVE", 2 ' Text movement - dont move the line"
	acadDoc.SetVariable "DIMEXE", 0.0625 ' Extension above dimension line"
End Sub

Sub dim_feet()
	acadDoc.SetVariable "DIMLUNIT", 4 ' Linear unit format"
	acadDoc.SetVariable "DIMFRAC", 2 ' Fraction format"
	acadDoc.SetVariable "DIMZIN", 3 ' Zero suppression"
End Sub

Sub dim_decimal()
	acadDoc.SetVariable "DIMLUNIT", 2 ' Linear unit format"
	acadDoc.SetVariable "DIMADEC", 1 ' Angular decimal places
End Sub

Sub getfont()
	'Connect_Acad
	Dim styles As AcadTextStyles
	Dim style As AcadTextStyle
	Set styles = acadDoc.TextStyles
	Dim strtypeface As String
	Dim bold As Boolean, italic As Boolean
	Dim lngchar As Long, lngpitch As Long
	For Each style In styles
	style.getfont strtypeface, bold, italic, lngchar, lngpitch
	Debug.Print style.Name 'user supplied name in the list box of the style dialog
	Debug.Print style.fontFile 'actual file name not shown in style dialog
	Debug.Print strtypeface 'font name dropdown box in style dialog
	Debug.Print bold
	Debug.Print italic
	Debug.Print lngchar
	Debug.Print lngpitch
	Debug.Print
	Next
End Sub

Sub new_textstyle(str_stylename As String, str_typeface As String)
	Dim bold As Boolean, italic As Boolean
	Dim lngchar As Long, lngpitch As Long
	lngchar = 0
	lngpitch = 34 'i am sure this is not meaningless but this is typ(swiss 32 variable 2)
	bold = False
	-1-D:\_S\Excel workbook\CAD_Col\11.Dimensions\mk_dimensionStyle3.bas Thursday, May 20, 2021 7:27 PM
	italic = False
	Dim TextStyles As AcadTextStyles
	Dim curStyle As AcadTextStyle
	Dim newStyle As AcadTextStyle
	Set curStyle = acadDoc.ActiveTextStyle
	Set TextStyles = acadDoc.TextStyles
	Set newStyle = TextStyles.Add(str_stylename)
	acadDoc.ActiveTextStyle = newStyle
	'new style is added with no font information
	'autocad assigns defaults similar or same as standard
	newStyle.SetFont str_typeface, bold, italic, lngchar, lngpitch
	'sometimes i get a transient new style at this point
	'eliminated by the next new style unless i create actual text
End Sub

Sub new_dimstyle(strname As String, strtype As String, dm As Integer)
	Dim style As AcadDimStyle
	Dim strdimstyle As String
	Dim strtextstyle As String
	strtextstyle = acadDoc.GetVariable("textstyle")
	strdimstyle = strname & "_" & strtype & "_" & dm
	Set style = acadDoc.DimStyles.Add(strdimstyle)
	acadDoc.ActiveDimStyle = style
	Select Case strtype
	Case "Inch"
	Call dim_inch
	Case "Feet"
	Call dim_inch
	Call dim_feet
	Case "Decimal"
	Call dim_inch
	Call dim_feet
	Call dim_decimal
	End Select
	acadDoc.SetVariable "DIMTXSTY", strtextstyle
	acadDoc.SetVariable "DIMSCALE", dm
	style.CopyFrom acadDoc 'the basic method for changing style contents
End Sub

Sub test_dim()
	Call Connect_Acad
	Call new_textstyle("Arial Narrow", "Arial Narrow")
	Call new_dimstyle("ArialN", "Inch", 24)
	Call new_textstyle("Technic", "Technic")
	Call new_dimstyle("Technic", "Feet", 24)
	Call new_textstyle("Courier", "Courier New")
	Call new_dimstyle("Courier", "Decimal", 24)
	-2-D:\_S\Excel workbook\CAD_Col\11.Dimensions\mk_dimensionStyle3.bas Thursday, May 20, 2021 7:27 PM
End Sub

Option Explicit
Sub New_Layer()
	Dim acadApp As AcadApplication
	Dim acadDoc As AcadDocument
	Dim mSp As AcadModelSpace
	Dim dimstyle As AcadDimStyle
	Dim sDim As AcadDimAligned
	Dim point1(0 To 2) As Double
	Dim point2(0 To 2) As Double
	Dim location(0 To 2) As Double
	'Check if AutoCAD is open.
	On Error Resume Next
	Set acadApp = GetObject(, "AutoCAD.Application")
	On Error GoTo 0
	'If AutoCAD is not opened create a new instance and make it visible.
	If acadApp Is Nothing Then
	Set acadApp = New AcadApplication
	acadApp.Visible = True
	End If
	'Check if there is an active drawing.
	On Error Resume Next
	Set acadDoc = acadApp.ActiveDocument
	On Error GoTo 0
	'No active drawing found. Create a new one.
	If acadDoc Is Nothing Then
	Set acadDoc = acadApp.Documents.Add
	acadApp.Visible = True
	End If
	Set mSp = acadDoc.ModelSpace
	'Dimension points
	point1(0) = 0#: point1(1) = 5#: point1(2) = 0#
	point2(0) = 6.1: point2(1) = 5: point2(2) = 0#
	location(0) = 5#: location(1) = 4.4: location(2) = 0#
	'Add dimension
	Set sDim = acadDoc.ModelSpace.AddDimAligned(point1, point2, location)
	'Set dimension properties
	sDim.color = acByLayer
	sDim.ExtensionLineExtend = 0
	sDim.Arrowhead1Type = acArrowOblique
	sDim.Arrowhead2Type = acArrowOblique
	sDim.ArrowheadSize = 0.1
	sDim.TextColor = acGreen
	sDim.TextHeight = 0.2
	sDim.UnitsFormat = acDimLDecimal
	sDim.PrimaryUnitsPrecision = acDimPrecisionOne
	sDim.TextGap = 0.1
	sDim.LinearScaleFactor = 100
	sDim.ExtensionLineOffset = 0.1
	sDim.VerticalTextPosition = acOutside
	'Create a new dimension style
	Set dimstyle = acadDoc.DimStyles.Add("D100")
	'Copy dimension properties from previously added dimension
	dimstyle.CopyFrom (sDim)
	'Delete dimension
	sDim.Delete
End Sub

Public Sub UgDimStyle_CreateNew( _
	Optional ByVal sDimStyName As String = "Standard_Dim")
	Dim CurDimStyle As AcadDimStyle
	Dim NewDimstyle As AcadDimStyle
	Dim iAltUnits As Integer
	Dim dDimScale As Double
	'Save copy of current dimstyle
	Set CurDimStyle = ThisDrawing.ActiveDimStyle
	Create New dimstyle
	Set NewDimstyle = ThisDrawing.DimStyles.Add(sDimStyName)
	'Set newly created dimstyle current
	ThisDrawing.ActiveDimStyle = NewDimstyle
	'Save the target "dimvar" values
	dDimScale = ThisDrawing.GetVariable("Dimscale")
	iAltUnits = ThisDrawing.GetVariable("Dimalt")
	'Alter the target "dimvar" values
	ThisDrawing.SetVariable "DIMSCALE", 1# 'will control size of dim text
	ThisDrawing.SetVariable "DIMASZ", 2.5 'arrowhead size
	ThisDrawing.SetVariable "DIMATFIT", 2 'arrow-text arrangement
	ThisDrawing.SetVariable "DIMAZIN", 3 '0 suppression before/after angular
	ThisDrawing.SetVariable "DIMBLK", "" 'special arrow blk
	ThisDrawing.SetVariable "DIMDLE", 0 'dim line extension past extension
	ThisDrawing.SetVariable "DIMDLI", 10 'dist between baseline dims
	ThisDrawing.SetVariable "DIMDSEP", "." 'decimal separator
	ThisDrawing.SetVariable "DIMEXE", 1 'dim line extension past extension
	ThisDrawing.SetVariable "DIMEXO", 1 'dim offset from origin
	ThisDrawing.SetVariable "DIMFIT", 5 'control fit if not enough space
	ThisDrawing.SetVariable "DIMGAP", 2 'gap around text
	ThisDrawing.SetVariable "DIMJUST", 0 'text placement - above centered
	ThisDrawing.SetVariable "DIMLFAC", 1# 'length scaling
	ThisDrawing.SetVariable "DIMTAD", 1 'text to dim placement - above
	ThisDrawing.SetVariable "DIMTIH", 0 'aligned with dim
	ThisDrawing.SetVariable "DIMTIX", 0 'force inside
	ThisDrawing.SetVariable "DIMTMOVE", 0 'dim moves with text
	ThisDrawing.SetVariable "DIMTSZ", 0 'draw arrowheads
	ThisDrawing.SetVariable "DIMTXT", 3.5 'text height
	ThisDrawing.SetVariable "DIMTZIN", 12 '0 suppression before/after tol
	ThisDrawing.SetVariable "DIMUNIT", 2 'unit format - decimal
	ThisDrawing.SetVariable "DIMZIN", 12 '0 suppression before/after
	'Copy new document dimvar settings into new dimstyle
	NewDimstyle.CopyFrom ThisDrawing
	'Set original dimstyle current
	ThisDrawing.ActiveDimStyle = CurDimStyle
	'Restore the altered "dimvar" values
	ThisDrawing.SetVariable "Dimscale", dDimScale
	ThisDrawing.SetVariable "Dimalt", iAltUnits
	'Copy restored document dimvar settings into original dimstyle
	CurDimStyle.CopyFrom ThisDrawing
	Set CurDimStyle = Nothing
	Set NewDimstyle = Nothing
	-1-D:\_S\Excel workbook\CAD_Col\11.Dimensions\mk_DimStyleCreate.bas Thursday, May 20, 2021 7:26 PM
End Sub

Public Sub UgDimStyle_CreateNew( _
	Optional ByVal sDimStyName As String = "Standard_Dim")
	Dim CurDimStyle As AcadDimStyle
	Dim NewDimstyle As AcadDimStyle
	Dim iAltUnits As Integer
	Dim dDimScale As Double
	'Save copy of current dimstyle
	Set CurDimStyle = ThisDrawing.ActiveDimStyle
	Create New dimstyle
	Set NewDimstyle = ThisDrawing.DimStyles.Add(sDimStyName)
	'Set newly created dimstyle current
	ThisDrawing.ActiveDimStyle = NewDimstyle
	'Save the target "dimvar" values
	dDimScale = ThisDrawing.GetVariable("Dimscale")
	iAltUnits = ThisDrawing.GetVariable("Dimalt")
	'Alter the target "dimvar" values
	ThisDrawing.SetVariable "DIMSCALE", 1# 'will control size of dim text
	ThisDrawing.SetVariable "DIMASZ", 2.5 'arrowhead size
	ThisDrawing.SetVariable "DIMATFIT", 2 'arrow-text arrangement
	ThisDrawing.SetVariable "DIMAZIN", 3 '0 suppression before/after angular
	ThisDrawing.SetVariable "DIMBLK", "" 'special arrow blk
	ThisDrawing.SetVariable "DIMDLE", 0 'dim line extension past extension
	ThisDrawing.SetVariable "DIMDLI", 10 'dist between baseline dims
	ThisDrawing.SetVariable "DIMDSEP", "." 'decimal separator
	ThisDrawing.SetVariable "DIMEXE", 1 'dim line extension past extension
	ThisDrawing.SetVariable "DIMEXO", 1 'dim offset from origin
	ThisDrawing.SetVariable "DIMFIT", 5 'control fit if not enough space
	ThisDrawing.SetVariable "DIMGAP", 2 'gap around text
	ThisDrawing.SetVariable "DIMJUST", 0 'text placement - above centered
	ThisDrawing.SetVariable "DIMLFAC", 1# 'length scaling
	ThisDrawing.SetVariable "DIMTAD", 1 'text to dim placement - above
	ThisDrawing.SetVariable "DIMTIH", 0 'aligned with dim
	ThisDrawing.SetVariable "DIMTIX", 0 'force inside
	ThisDrawing.SetVariable "DIMTMOVE", 0 'dim moves with text
	ThisDrawing.SetVariable "DIMTSZ", 0 'draw arrowheads
	ThisDrawing.SetVariable "DIMTXT", 3.5 'text height
	ThisDrawing.SetVariable "DIMTZIN", 12 '0 suppression before/after tol
	ThisDrawing.SetVariable "DIMUNIT", 2 'unit format - decimal
	ThisDrawing.SetVariable "DIMZIN", 12 '0 suppression before/after
	'Copy new document dimvar settings into new dimstyle
	NewDimstyle.CopyFrom ThisDrawing
	'Set original dimstyle current
	ThisDrawing.ActiveDimStyle = CurDimStyle
	'Restore the altered "dimvar" values
	ThisDrawing.SetVariable "Dimscale", dDimScale
	ThisDrawing.SetVariable "Dimalt", iAltUnits
	'Copy restored document dimvar settings into original dimstyle
	CurDimStyle.CopyFrom ThisDrawing
	Set CurDimStyle = Nothing
	Set NewDimstyle = Nothing
	-1-D:\_S\Excel workbook\CAD_Col\11.Dimensions\mk_DimStylefmExsting.bas Thursday, May 20, 2021 7:27 PM
End Sub

Sub Example_AddDimRotated()
	' This example creates a rotated dimension in model space.
	Dim dimObj As AcadDimRotated
	Dim point1(0 To 2) As Double
	Dim point2(0 To 2) As Double
	Dim location(0 To 2) As Double
	Dim rotAngle As Double
	' Define the dimension
	point1(0) = 0#: point1(1) = 5#: point1(2) = 0#
	point2(0) = 5#: point2(1) = 5#: point2(2) = 0#
	location(0) = 0#: location(1) = 0#: location(2) = 0#
	rotAngle = 120
	rotAngle = rotAngle * 3.141592 / 180# ' covert to Radians
	' Create the rotated dimension in model space
	Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, rotAngle)
	ZoomAll
End Sub

Sub Example_AddDimOrdinate()
	Dim aObj As AcadEntity
	For Each aObj In ThisDrawing.ModelSpace
	If TypeOf aObj Is AcadDimRotated Then
	Dim testDim As AcadDimRotated
	Set testDim = aObj
	If testDim.Measurement = 12# Then
	testDim.TextOverride = "24.0"
	End If
	End If
	Next aObj
End Sub

Sub Example_TextStyle()
	' This example creates an aligned dimension in model space and
	' creates a new system text style. The new text style is then assigned to
	' the new dimension
	Dim dimObj As AcadDimAligned
	Dim newText As AcadTextStyle
	Dim point1(0 To 2) As Double, point2(0 To 2) As Double
	Dim location(0 To 2) As Double
	' Define the dimension
	point1(0) = 5: point1(1) = 5: point1(2) = 0
	point2(0) = 5.5: point2(1) = 5: point2(2) = 0
	location(0) = 5: location(1) = 7: location(2) = 0
	' Create an aligned dimension object in model space
	Set dimObj = ThisDrawing.ModelSpace.AddDimAligned(point1, point2, location)
	' Create new text style
	Set newText = ThisDrawing.TextStyles.Add("MYSTYLE")
	newText.Height = 0.5 ' Just set the height of the new style so we can differentiate
	ThisDrawing.Application.ZoomAll
	' Read and display the current text style for this dimension
	MsgBox "The text style is currently set to: " & dimObj.TextStyle
	' Change the text style to use the new style we created
	dimObj.TextStyle = "MYSTYLE"
	ThisDrawing.Regen acAllViewports
	' Read and display the current text style for this dimension
	MsgBox "The text style is now set to: " & dimObj.TextStyle
End Sub

Public Sub AddDiametricDimension()
	Dim varFirstPoint As Variant
	Dim varSecondPoint As Variant
	Dim dblLeaderLength As Double
	Dim objDimDiametric As AcadDimDiametric
	Dim intOsmode As Integer
	'get original object snap settings
	intOsmode = ThisDrawing.GetVariable("osmode")
	ThisDrawing.SetVariable "osmode", 512 ' Near
	With ThisDrawing.Utility
	varFirstPoint = .GetPoint(, "Select first point on circle: ")
	ThisDrawing.SetVariable "osmode", 128 ' Per
	varSecondPoint = .GetPoint(varFirstPoint, _
	"Select a point opposite the first: ")
	dblLeaderLength = .GetDistance(varFirstPoint, _
	"Enter leader length from first point: ")
	End With
	Set objDimDiametric = ThisDrawing.ModelSpace.AddDimDiametric( _
	varFirstPoint, varSecondPoint, dblLeaderLength)
	objDimDiametric.UnitsFormat = acDimLEngineering
	objDimDiametric.PrimaryUnitsPrecision = acDimPrecisionFive
	objDimDiametric.FractionFormat = acNotStacked
	objDimDiametric.Update
	'reinstate original object snap settings
	ThisDrawing.SetVariable "osmode", intOsmode
End Sub

Public Sub ChangeDimStyle()
	Dim objDimension As AcadDimension
	Dim varPickedPoint As Variant
	Dim objDimStyle As AcadDimStyle
	Dim strDimStyles As String
	Dim strChosenDimStyle As String
	On Error Resume Next
	ThisDrawing.Utility.GetEntity objDimension, varPickedPoint, _
	"Pick a dimension whose style you wish to set"
	If objDimension Is Nothing Then
	MsgBox "You failed to pick a dimension object"
	Exit Sub
	End If
	For Each objDimStyle In ThisDrawing.DimStyles
	strDimStyles = strDimStyles & objDimStyle.Name & vbCrLf
	Next objDimStyle
	strChosenDimStyle = InputBox("Choose one of the following " & _
	"Dimension styles to apply" & vbCrLf & strDimStyles)
	If strChosenDimStyle = "" Then Exit Sub
	objDimension.StyleName = strChosenDimStyle
End Sub

Public Sub SetActiveDimStyle()
	Dim strDimStyles As String
	Dim strChosenDimStyle As String
	Dim objDimStyle As AcadDimStyle
	For Each objDimStyle In ThisDrawing.DimStyles
	strDimStyles = strDimStyles & objDimStyle.Name & vbCrLf
	Next
	strChosenDimStyle = InputBox("Choose one of the following Dimension " & _
	"styles:" & vbCr & vbCr & strDimStyles, "Existing Dimension style is: " & _
	ThisDrawing.ActiveDimStyle.Name, ThisDrawing.ActiveDimStyle.Name)
	If strChosenDimStyle = "" Then Exit Sub
	On Error Resume Next
	ThisDrawing.ActiveDimStyle = ThisDrawing.DimStyles(strChosenDimStyle)
	If Err Then MsgBox "Dimension style was not recognized"
End Sub

Public Sub ChangeTextStyle()
	Dim strTextStyles As String
	Dim objTextStyle As AcadTextStyle
	Dim objLayer As AcadLayer
	Dim strLayerName As String
	Dim strStyleName As String
	Dim objAcadObject As AcadObject
	On Error Resume Next
	For Each objTextStyle In ThisDrawing.TextStyles
	strTextStyles = strTextStyles & vbCr & objTextStyle.Name
	Next
	strStyleName = InputBox("Enter name of style to apply:" & vbCr & _
	strTextStyles, "TextStyles", ThisDrawing.ActiveTextStyle.Name)
	Set objTextStyle = ThisDrawing.TextStyles(strStyleName)
	If objTextStyle Is Nothing Then
	MsgBox "Style does not exist"
	Exit Sub
	End If
	For Each objAcadObject In ThisDrawing.ModelSpace
	If objAcadObject.ObjectName = "AcDbMText" Or _
	objAcadObject.ObjectName = "AcDbText" Then
	objAcadObject.StyleName = strStyleName
	objAcadObject.Update
	End If
	Next
End Sub

Public Sub Add3PointAngularDimension()
	Dim varAngularVertex As Variant
	Dim varFirstPoint As Variant
	Dim varSecondPoint As Variant
	Dim varTextLocation As Variant
	Dim objDim3PointAngular As AcadDim3PointAngular
	'Define the dimension
	varAngularVertex = ThisDrawing.Utility.GetPoint(, _
	"Enter the center point: ")
	varFirstPoint = ThisDrawing.Utility.GetPoint(varAngularVertex, _
	"Select first point: ")
	varSecondPoint = ThisDrawing.Utility.GetPoint(varAngularVertex, _
	"Select second point: ")
	varTextLocation = ThisDrawing.Utility.GetPoint(varAngularVertex, _
	"Pick dimension text location: ")
	Set objDim3PointAngular = ThisDrawing.ModelSpace.AddDim3PointAngular( _
	varAngularVertex, varFirstPoint, varSecondPoint, varTextLocation)
	objDim3PointAngular.Update
End Sub

Public Sub AddAlignedDimension()
	Dim varFirstPoint As Variant
	Dim varSecondPoint As Variant
	Dim varTextLocation As Variant
	Dim objDimAligned As AcadDimAligned
	'Define the dimension
	varFirstPoint = ThisDrawing.Utility.GetPoint(, "Select first point: ")
	varSecondPoint = ThisDrawing.Utility.GetPoint(varFirstPoint, _
	"Select second point: ")
	varTextLocation = ThisDrawing.Utility.GetPoint(, _
	"Pick dimension text location: ")
	'Create an aligned dimension
	Set objDimAligned = ThisDrawing.ModelSpace.AddDimAligned(varFirstPoint, _
	varSecondPoint, varTextLocation)
	objDimAligned.Update
	MsgBox "Now we will change to Engineering units format"""
	objDimAligned.UnitsFormat = acDimLEngineering
	objDimAligned.Update
End Sub

Public Sub AddAngularDimension()
	Dim varAngularVertex As Variant
	Dim varFirstPoint As Variant
	Dim varSecondPoint As Variant
	Dim varTextLocation As Variant
	Dim objDimAngular As AcadDimAngular
	'Define the dimension
	varAngularVertex = ThisDrawing.Utility.GetPoint(, _
	"Enter the center point: ")
	varFirstPoint = ThisDrawing.Utility.GetPoint(varAngularVertex, _
	"Select first point: ")
	varSecondPoint = ThisDrawing.Utility.GetPoint(varAngularVertex, _
	"Select second point: ")
	varTextLocation = ThisDrawing.Utility.GetPoint(varAngularVertex, _
	"Pick dimension text location: ")
	'Create an angular dimension
	Set objDimAngular = ThisDrawing.ModelSpace.AddDimAngular( _
	varAngularVertex, varFirstPoint, varSecondPoint, varTextLocation)
	objDimAngular.AngleFormat = acGrads
	objDimAngular.Update
	MsgBox "Angle measured in GRADS"
	objDimAngular.AngleFormat = acDegreeMinuteSeconds
	objDimAngular.TextPrecision = acDimPrecisionFour
	objDimAngular.Update
	MsgBox "Angle measured in Degrees Minutes Seconds"
End Sub

Public Sub AddOrdinateDimension()
	Dim varBasePoint As Variant
	Dim varLeaderEndPoint As Variant
	Dim blnUseXAxis As Boolean
	Dim strKeywordList As String
	Dim strAnswer As String
	Dim objDimOrdinate As AcadDimOrdinate
	strKeywordList = "X Y"
	'Define the dimension
	varBasePoint = ThisDrawing.Utility.GetPoint(, _
	"Select ordinate dimension position: ")
	ThisDrawing.Utility.InitializeUserInput 1, strKeywordList
	strAnswer = ThisDrawing.Utility.GetKeyword("Along Which Axis? <X/Y>: ")
	If strAnswer = "X" Then
	varLeaderEndPoint = ThisDrawing.Utility.GetPoint(varBasePoint, _
	"Select X point for dimension text: ")
	blnUseXAxis = True
	Else
	varLeaderEndPoint = ThisDrawing.Utility.GetPoint(varBasePoint, _
	"Select Y point for dimension text: ")
	blnUseXAxis = False
	End If
	'Create an ordinate dimension
	Set objDimOrdinate = ThisDrawing.ModelSpace.AddDimOrdinate( _
	varBasePoint, varLeaderEndPoint, blnUseXAxis)
	objDimOrdinate.TextSuffix = "units"
	objDimOrdinate.Update
End Sub

Public Sub AddRadialDimension()
	Dim objUserPickedEntity As Object
	Dim varEntityPickedPoint As Variant
	Dim varEdgePoint As Variant
	Dim dblLeaderLength As Double
	Dim objDimRadial As AcadDimRadial
	Dim intOsmode As Integer
	intOsmode = ThisDrawing.GetVariable("osmode")
	ThisDrawing.SetVariable "osmode", 512 ' Near
	'Define the dimension
	On Error Resume Next
	With ThisDrawing.Utility
	.GetEntity objUserPickedEntity, varEntityPickedPoint, _
	"Pick Arc or Circle:"
	If objUserPickedEntity Is Nothing Then
	MsgBox "You did not pick an entity"
	Exit Sub
	End If
	varEdgePoint = .GetPoint(objUserPickedEntity.Center, _
	"Pick edge point")
	dblLeaderLength = .GetReal("Enter leader length from this point: ")
	End With
	'Create the radial dimension
	Set objDimRadial = ThisDrawing.ModelSpace.AddDimRadial( _
	objUserPickedEntity.Center, varEdgePoint, dblLeaderLength)
	objDimRadial.ArrowheadType = acArrowArchTick
	objDimRadial.Update
	'reinstate original setting
	ThisDrawing.SetVariable "osmode", intOsmode
End Sub

Public Sub AddRotatedDimension()
	Dim varFirstPoint As Variant
	Dim varSecondPoint As Variant
	Dim varTextLocation As Variant
	Dim strRotationAngle As String
	Dim objDimRotated As AcadDimRotated
	'Define the dimension
	With ThisDrawing.Utility
	varFirstPoint = .GetPoint(, "Select first point: ")
	varSecondPoint = .GetPoint(varFirstPoint, "Select second point: ")
	varTextLocation = .GetPoint(, "Pick dimension text location: ")
	strRotationAngle = .GetString(False, "Enter rotation angle in degrees")
	End With
	'Create a rotated dimension
	Set objDimRotated = ThisDrawing.ModelSpace.AddDimRotated(varFirstPoint, _
	varSecondPoint, varTextLocation, _
	ThisDrawing.Utility.AngleToReal(strRotationAngle, acDegrees))
	objDimRotated.DecimalSeparator = ","
	objDimRotated.Update
End Sub

Public Sub GetTextSettings()
	Dim objTextStyle As AcadTextStyle
	Dim strTextStyleName As String
	Dim strTextStyles As String
	Dim strTypeFace As String
	Dim blnBold As Boolean
	Dim blnItalic As Boolean
	Dim lngCharacterSet As Long
	Dim lngPitchandFamily As Long
	Dim strText As String
	' Get the name of each text style in the drawing
	For Each objTextStyle In ThisDrawing.TextStyles
	strTextStyles = strTextStyles & vbCr & objTextStyle.Name
	Next
	' Ask the user to select the Text Style to look at
	strTextStyleName = InputBox("Please enter the name of the TextStyle " & _
	"whose setting you would like to see" & vbCr & _
	strTextStyles, "TextStyles", ThisDrawing.ActiveTextStyle.Name)
	' Exit the program if the user input was cancelled or empty
	If strTextStyleName = "" Then Exit Sub
	On Error Resume Next
	Set objTextStyle = ThisDrawing.TextStyles(strTextStyleName)
	' Check for existence the text style
	If objTextStyle Is Nothing Then
	MsgBox "This text style does not exist"
	Exit Sub
	End If
	' Get the Font properties
	objTextStyle.GetFont strTypeFace, blnBold, blnItalic, lngCharacterSet, _
	lngPitchandFamily
	' Check for Type face
	If strTypeFace = "" Then ' No True type
	MsgBox "Text Style: " & objTextStyle.Name & vbCr & _
	"Using file font: " & objTextStyle.fontFile, _
	vbInformation, "Text Style: " & objTextStyle.Name
	Else
	' True Type font info
	strText = "The text style: " & strTextStyleName & " has " & vbCrLf & _
	"a " & strTypeFace & " type face"
	If blnBold Then strText = strText & vbCrLf & " and is bold"
	If blnItalic Then strText = strText & vbCrLf & " and is italicized"
	MsgBox strText & vbCr & "Using file font: " & objTextStyle.fontFile, _
	vbInformation, "Text Style: " & objTextStyle.Name
	End If
End Sub

Public Sub CreateStraightLeaderWithNote()
	Dim dblPoints(5) As Double
	Dim varStartPoint As Variant
	Dim varEndPoint As Variant
	Dim intLeaderType As Integer
	Dim objAcadLeader As AcadLeader
	Dim objAcadMtext As AcadMText
	Dim strMtext As String
	Dim intI As Integer
	intLeaderType = acLineWithArrow
	varStartPoint = ThisDrawing.Utility.GetPoint(, _
	"Select leader start point: ")
	varEndPoint = ThisDrawing.Utility.GetPoint(varStartPoint, _
	"Select leader end point: ")
	For intI = 0 To 2
	dblPoints(intI) = varStartPoint(intI)
	dblPoints(intI + 3) = varEndPoint(intI)
	Next
	strMtext = InputBox("Notes:", "Leader Notes")
	If strMtext = "" Then Exit Sub
	' Create the text for the leader
	Set objAcadMtext = ThisDrawing.ModelSpace.AddMText(varEndPoint, _
	Len(strMtext) * ThisDrawing.GetVariable("dimscale"), strMtext)
	' Flip the alignment direction of the text
	If varEndPoint(0) > varStartPoint(0) Then
	objAcadMtext.AttachmentPoint = acAttachmentPointMiddleLeft
	Else
	objAcadMtext.AttachmentPoint = acAttachmentPointMiddleRight
	End If
	objAcadMtext.InsertionPoint = varEndPoint
	'Create the leader object
	Set objAcadLeader = ThisDrawing.ModelSpace.AddLeader(dblPoints, _
	objAcadMtext, intLeaderType)
	objAcadLeader.Update
End Sub

Public Sub SetDefaultTextStyle()
	Dim strTextStyles As String
	Dim objTextStyle As AcadTextStyle
	Dim strTextStyleName As String
	For Each objTextStyle In ThisDrawing.TextStyles
	strTextStyles = strTextStyles & vbCr & objTextStyle.Name
	Next
	strTextStyleName = InputBox("Enter name of style to apply:" & vbCr & _
	strTextStyles, "TextStyles", ThisDrawing.ActiveTextStyle.Name)
	If strTextStyleName = "" Then Exit Sub
	On Error Resume Next
	Set objTextStyle = ThisDrawing.TextStyles(strTextStyleName)
	If objTextStyle Is Nothing Then
	MsgBox "This text style does not exist"
	Exit Sub
	End If
	ThisDrawing.ActiveTextStyle = objTextStyle
End Sub

Public Sub CreateTolerance()
	Dim strToleranceText As String
	Dim varInsertionPoint As Variant
	Dim varTextDirection As Variant
	Dim intI As Integer
	Dim objTolerance As AcadTolerance
	strToleranceText = InputBox("Please enter the text for the tolerance")
	varInsertionPoint = ThisDrawing.Utility.GetPoint(, _
	"Please enter the insertion point for the tolerance")
	varTextDirection = ThisDrawing.Utility.GetPoint(varInsertionPoint, _
	"Please enter a direction for the tolerance")
	For intI = 0 To 2
	varTextDirection(intI) = varTextDirection(intI) - varInsertionPoint(intI)
	Next
	Set objTolerance = ThisDrawing.ModelSpace.AddTolerance(strToleranceText, _
	varInsertionPoint, varTextDirection)
End Sub


'12.SelectionSets

Sub selEntByPline()
	On Error Resume Next
	Dim objCadEnt As AcadEntity
	ThisDrawing.Utility.GetEntity objCadEnt, vrRetPnt
	Dim vrRetPnt As Variant
	If objCadEnt.ObjectName = "AcDbPolyline" Then '|-- Checking for 2D Polylines --|
	Dim objLWPline As AcadLWPolyline
	Dim objSSet As AcadSelectionSetlv
	Dim dblCurCords() As Double
	Dim dblNewCords() As Double
	Dim iMaxCurArr, iMaxNewArr As Integer
	Dim iCurArrIdx, iNewArrIdx, iCnt As Integer
	Set objLWPline = objCadEnt
	dblCurCords = objLWPline.Coordinates '|-- The returned coordinates are 2D only --|
	iMaxCurArr = UBound(dblCurCords)
	If iMaxCurArr = 3 Then
	ThisDrawing.Utility.Prompt "The selected polyline should have minimum 2 segments..."
	Exit Sub
	Else
	'|-- The 2D Coordinates are insufficient to use in SelectByPolygon method --|
	'|-- So convert those into 3D coordinates --|
	iMaxNewArr = ((iMaxCurArr + 1) * 1.5) - 1 '|-- New array dimension
	ReDim dblNewCords(iMaxNewArr) As Double
	iCurArrIdx = 0: iCnt = 1
	For iNewArrIdx = 0 To iMaxNewArr
	If iCnt = 3 Then '|-- The z coordinate is set to 0 --|
	dblNewCords(iNewArrIdx) = 0
	iCnt = 1
	Else
	dblNewCords(iNewArrIdx) = dblCurCords(iCurArrIdx)
	iCurArrIdx = iCurArrIdx + 1
	iCnt = iCnt + 1
	End If
	Next
	Set objSSet = ThisDrawing.SelectionSets.Add("SEL_ENT")
	objSSet.SelectByPolygon acSelectionSetWindowPolygon, dblNewCords
	objSSet.Highlight True
	objSSet.Delete
	End If
	Else
	ThisDrawing.Utility.Prompt "The selected object is not a 2D Polyline...."
	End If
	If Err.Number <> 0 Then
	MsgBox Err.Description
	Err.Clear
	End If
End Sub

Option Explicit
' request check on 'Break on Unhadled Errors' option button
' in Tools->Options->General
Public Function IsGroupExist(groupName As String) As Boolean
	' Frank Oquendo's technic
	Dim oGroup As AcadGroup
	On Error Resume Next
	Set oGroup = ThisDrawing.Groups.Item(groupName)
	IsGroupExist = (Err.Number = 0)
End Function
''=========================''
Public Sub GetGroupItems()
	Dim oGroup As AcadGroup
	Dim gpName As String
	Dim objEnt As AcadObject
	On Error GoTo Err_Control
	gpName = InputBox("Enter group name to discover items: ", "Group Manipulation")
	' if emtpy string entered
	If gpName = vbNullString Then Exit Sub
	' check for existing group
	If Not IsGroupExist(gpName) Then
	MsgBox "Group named " & gpName & vbCr & _
	"does not exist"
	Exit Sub
	End If
	' get the group from Groups collection
	Set oGroup = ThisDrawing.Groups.Item(gpName)
	' Check fro count items in the group
	MsgBox "There are " & oGroup.count & " items in this group"
	' Test
	' loop trough the group items, say change color to red
	For Each objEnt In oGroup
	objEnt.color = acRed
	Debug.Print objEnt.ObjectName
	Next
	Exit_Here:
	' Clean up
	If Not oGroup Is Nothing Then
	Set oGroup = Nothing
	End If
	Exit Sub
	' error handler
	Err_Control:
	If Err.Number <> 0 Then
	MsgBox Err.Description
	Err.Clear
	End If
	Resume Exit_Here
End Sub

Public Sub TestAddGroup()
	Dim objGroup As AcadGroup
	Dim strName As String
	On Error Resume Next
	'' get a name from user
	strName = InputBox("Enter a new group name: ")
	If "" = strName Then Exit Sub
	Set objGroup = ThisDrawing.Groups.Item(strName)
	'' create it
	If Not objGroup Is Nothing Then
	MsgBox "Group already exists"
	Exit Sub
	End If
	Set objGroup = ThisDrawing.Groups.Add(strName)
	'' check if it was created
	If objGroup Is Nothing Then
	MsgBox "Unable to Add '" & strName & "'"
	Else
	MsgBox "Added group '" & objGroup.Name & "'"
	End If
End Sub

Public Sub TestGroupAppendRemove()
	Dim objSS As AcadSelectionSet
	Dim objGroup As AcadGroup
	Dim objEnts() As AcadEntity
	Dim strName As String
	Dim strOpt As String
	Dim intI As Integer
	On Error Resume Next
	'' set pickstyle to NOT select groups
	ThisDrawing.SetVariable "Pickstyle", 2
	With ThisDrawing.Utility
	'' get group name from user
	strName = .GetString(True, vbCr & "Group name: ")
	If Err Or "" = strName Then GoTo Done
	'' get the existing group or add new one
	Set objGroup = ThisDrawing.Groups.Add(strName)
	'' pause for the user
	.Prompt vbCr & "Group contains: " & objGroup.Count & " entities" & _
	vbCrLf
	'' get input for mode
	.InitializeUserInput 1, "Append Remove"
	strOpt = .GetKeyword(vbCr & "Option [Append/Remove]: ")
	If Err Then GoTo Done
	'' create a new selectionset
	Set objSS = ThisDrawing.SelectionSets.Add("TestGroupAppendRemove")
	If Err Then GoTo Done
	'' get a selection set from user
	objSS.SelectOnScreen
	'' convert selection set to array
	'' resize the entity array to the selection size
	ReDim objEnts(objSS.Count - 1)
	'' copy entities from the selection to entity array
	For intI = 0 To objSS.Count - 1
	Set objEnts(intI) = objSS(intI)
	Next
	'' append or remove entities based on input
	If "Append" = strOpt Then
	objGroup.AppendItems objEnts
	Else
	objGroup.RemoveItems objEnts
	End If
	'' pause for the user
	.Prompt vbCr & "Group contains: " & objGroup.Count & " entities"
	'' unhighlight the entities
	objSS.Highlight False
	End With
	Done:
	If Err Then MsgBox "Error occurred: " & Err.Description
	'' if the selection was created, delete it
	If Not objSS Is Nothing Then
	objSS.Delete
	End If
End Sub

Public Sub TestAddSelectionSet()
	Dim objSS As AcadSelectionSet
	Dim strName As String
	On Error Resume Next
	'' get a name from user
	strName = InputBox("Enter a new selection set name: ")
	If "" = strName Then Exit Sub
	'' create it
	Set objSS = ThisDrawing.SelectionSets.Add(strName)
	'' check if it was created
	If objSS Is Nothing Then
	MsgBox "Unable to Add '" & strName & "'"
	Else
	MsgBox "Added selection set '" & objSS.Name & "'"
	End If
End Sub

Public Sub TestSelectErase()
	Dim objSS As AcadSelectionSet
	On Error GoTo Done
	With ThisDrawing.Utility
	'' create a new selectionset
	Set objSS = ThisDrawing.SelectionSets.Add("TestSelectErase")
	'' let user select entities interactively
	objSS.SelectOnScreen
	'' highlight the selected entities
	objSS.Highlight True
	'' erase the selected entities
	objSS.Erase
	'' prove that the selection is empty (but still viable)
	.Prompt vbCr & objSS.Count & " entities selected"
	End With
	Done:
	'' if the selection was created, delete it
	If Not objSS Is Nothing Then
	objSS.Delete
	End If
End Sub

Public Sub ListSelectionSets()
	Dim objSS As AcadSelectionSet
	Dim strSSList As String
	For Each objSS In ThisDrawing.SelectionSets
	strSSList = strSSList & vbCr & objSS.Name
	Next
	MsgBox strSSList, , "List of Selection Sets"
End Sub

Public Sub TestSelectByPolygon()
	Dim objSS As AcadSelectionSet
	Dim strOpt As String
	Dim lngMode As Long
	Dim varPoints As Variant
	On Error GoTo Done
	With ThisDrawing.Utility
	'' create a new selectionset
	Set objSS = ThisDrawing.SelectionSets.Add("TestSelectByPolygon1")
	'' get the mode from the user
	.InitializeUserInput 1, "Fence Window Crossing"
	strOpt = .GetKeyword(vbCr & "Select by [Fence/Window/Crossing]: ")
	'' convert keyword into mode
	Select Case strOpt
	Case "Fence": lngMode = acSelectionSetFence
	Case "Window": lngMode = acSelectionSetWindowPolygon
	Case "Crossing": lngMode = acSelectionSetCrossingPolygon
	End Select
	'' let user digitize points
	varPoints = InputPoints()
	'' select entities using mode and points specified
	objSS.SelectByPolygon lngMode, varPoints
	'' highlight the selected entities
	objSS.Highlight True
	'' pause for the user
	.Prompt vbCr & objSS.Count & " entities selected"
	.GetString False, vbLf & "Enter to continue "
	'' unhighlight the entities
	objSS.Highlight False
	End With
	Done:
	'' if the selection was created, delete it
	If Not objSS Is Nothing Then
	objSS.Delete
	End If
End Sub
Function InputPoints() As Variant
	Dim varStartPoint As Variant
	Dim varNextPoint As Variant
	Dim varWCSPoint As Variant
	Dim lngLast As Long
	Dim dblPoints() As Double
	On Error Resume Next
	'' get first points from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varStartPoint = .GetPoint(, vbLf & "Pick the start point: ")
	'' setup initial point
	ReDim dblPoints(2)
	dblPoints(0) = varStartPoint(0)
	dblPoints(1) = varStartPoint(1)
	dblPoints(2) = varStartPoint(2)
	varNextPoint = varStartPoint
	'' append vertexes in a loop
	Do
	'' translate picked point to UCS for basepoint below
	varWCSPoint = .TranslateCoordinates(varNextPoint, acWorld, _
	acUCS, True)
	'' get user point for new vertex, use last pick as basepoint
	varNextPoint = .GetPoint(varWCSPoint, vbCr & _
	"Pick another point <exit>: ")
	'' exit loop if no point picked
	If Err Then Exit Do
	'' get the upper bound
	lngLast = UBound(dblPoints)
	-1-D:\_S\Excel workbook\CAD_Col\12.SelectionSets\mk12_SelectByPolygonMethod.bas Thursday, May 20, 2021 7:32 PM
	'' expand the array
	ReDim Preserve dblPoints(lngLast + 3)
	'' add the new point
	dblPoints(lngLast + 1) = varNextPoint(0)
	dblPoints(lngLast + 2) = varNextPoint(1)
	dblPoints(lngLast + 3) = varNextPoint(2)
	Loop
	End With
	'' return the points
	InputPoints = dblPoints
End Function

Private Sub AcadDocument_SelectionChanged()
	Dim objSS As AcadSelectionSet
	Dim dblStart As Double
	'' get the pickfirst selection from drawing
	Set objSS = ThisDrawing.PickfirstSelectionSet
	'' highlight the selected entities
	objSS.Highlight True
	MsgBox "There are " & objSS.Count & " objects in selection set: " & objSS.Name
	'' delay for 1/2 second
	dblStart = Timer
	Do While Timer < dblStart + 0.5
	Loop
	'' unhighlight the selected entities
	objSS.Highlight False
End Sub

Public Sub TestSelectionSetFilter()
	Dim objSS As AcadSelectionSet
	Dim intCodes(0) As Integer
	Dim varCodeValues(0) As Variant
	Dim strName As String
	On Error GoTo Done
	With ThisDrawing.Utility
	strName = .GetString(True, vbCr & "Layer name to filter: ")
	If "" = strName Then Exit Sub
	'' create a new selectionset
	Set objSS = ThisDrawing.SelectionSets.Add("TestSelectionSetFilter")
	'' set the code for layer
	intCodes(0) = 8
	'' set the value specified by user
	varCodeValues(0) = strName
	'' filter the objects
	objSS.Select acSelectionSetAll, , , intCodes, varCodeValues
	'' highlight the selected entities
	objSS.Highlight True
	'' pause for the user
	.Prompt vbCr & objSS.Count & " entities selected"
	.GetString False, vbLf & "Enter to continue "
	'' unhighlight the entities
	objSS.Highlight False
	End With
	Done:
	'' if the selection was created, delete it
	If Not objSS Is Nothing Then
	objSS.Delete
	End If
End Sub

Public Sub TestSelectionSetOperator()
	Dim objSS As AcadSelectionSet
	Dim intCodes() As Integer
	Dim varCodeValues As Variant
	Dim strName As String
	On Error GoTo Done
	With ThisDrawing.Utility
	strName = .GetString(True, vbCr & "Layer name to exclude: ")
	If "" = strName Then Exit Sub
	'' create a new selectionset
	Set objSS = ThisDrawing.SelectionSets.Add("TestSelectionSetOperator")
	'' using 9 filters
	ReDim intCodes(9): ReDim varCodeValues(9)
	'' set codes and values - indented for clarity
	intCodes(0) = -4: varCodeValues(0) = "<and"
	intCodes(1) = -4: varCodeValues(1) = "<or"
	intCodes(2) = 0: varCodeValues(2) = "line"
	intCodes(3) = 0: varCodeValues(3) = "arc"
	intCodes(4) = 0: varCodeValues(4) = "circle"
	intCodes(5) = -4: varCodeValues(5) = "or>"
	intCodes(6) = -4: varCodeValues(6) = "<not"
	intCodes(7) = 8: varCodeValues(7) = strName
	intCodes(8) = -4: varCodeValues(8) = "not>"
	intCodes(9) = -4: varCodeValues(9) = "and>"
	'' filter the objects
	objSS.Select acSelectionSetAll, , , intCodes, varCodeValues
	'' highlight the selected entities
	objSS.Highlight True
	'' pause for the user
	.Prompt vbCr & objSS.Count & " entities selected"
	.GetString False, vbLf & "Enter to continue "
	'' unhighlight the entities
	objSS.Highlight False
	End With
	Done:
	'' if the selection was created, delete it
	If Not objSS Is Nothing Then
	objSS.Delete
	End If
End Sub

Public Sub TestSelectAddRemoveClear()
	Dim objSS As AcadSelectionSet
	Dim objSStmp As AcadSelectionSet
	Dim strType As String
	Dim objEnts() As AcadEntity
	Dim intI As Integer
	On Error Resume Next
	With ThisDrawing.Utility
	'' create a new selectionset
	Set objSS = ThisDrawing.SelectionSets.Add("ssAddRemoveClear")
	If Err Then GoTo Done
	'' create a new temporary selection
	Set objSStmp = ThisDrawing.SelectionSets.Add("ssAddRemoveClearTmp")
	If Err Then GoTo Done
	'' loop until the user has finished
	Do
	'' clear any pending errors
	Err.Clear
	'' get input for type
	.InitializeUserInput 1, "Add Remove Clear Exit"
	strType = .GetKeyword(vbCr & "Select [Add/Remove/Clear/Exit]: ")
	'' branch based on input
	If "Exit" = strType Then
	'' exit if requested
	Exit Do
	ElseIf "Clear" = strType Then
	'' unhighlight the main selection
	objSS.Highlight False
	'' clear the main set
	objSS.Clear
	'' otherwise, we're adding/removing
	Else
	'' clear the temporary selection
	objSStmp.Clear
	objSStmp.SelectOnScreen
	'' highlight the temporary selection
	objSStmp.Highlight True
	'' convert temporary selection to array
	'' resize the entity array to the selection size
	ReDim objEnts(objSStmp.Count - 1)
	'' copy entities from the selection to entity array
	For intI = 0 To objSStmp.Count - 1
	Set objEnts(intI) = objSStmp(intI)
	Next
	'' add/remove items from main selection using entity array
	If "Add" = strType Then
	objSS.AddItems objEnts
	Else
	objSS.RemoveItems objEnts
	End If
	'' unhighlight the temporary selection
	objSStmp.Highlight False
	'' highlight the main selection
	objSS.Highlight True
	End If
	Loop
	End With
	Done:
	'' if the selections were created, delete them
	If Not objSS Is Nothing Then
	'' unhighlight the entities
	objSS.Highlight False
	'' delete the main selection
	objSS.Delete
	-1-D:\_S\Excel workbook\CAD_Col\12.SelectionSets\mk12_TestSelectAddRemoveClear.bas Thursday, May 20, 2021 7:31 PM
	End If
	If Not objSStmp Is Nothing Then
	'' delete the temporary selection
	objSStmp.Delete
	End If
End Sub

Public Sub TestSelectAtPoint()
	Dim varPick As Variant
	Dim objSS As AcadSelectionSet
	On Error GoTo Done
	With ThisDrawing.Utility
	'' create a new selectionset
	Set objSS = ThisDrawing.SelectionSets.Add("TestSelectAtPoint")
	'' get a point of selection from the user
	varPick = .GetPoint(, vbCr & "Select entities at a point: ")
	'' let user select entities interactively
	objSS.SelectAtPoint varPick
	'' highlight the selected entities
	objSS.Highlight True
	'' pause for the user
	.Prompt vbCr & objSS.Count & " entities selected"
	.GetString False, vbLf & "Enter to continue "
	'' unhighlight the entities
	objSS.Highlight False
	End With
	Done:
	'' if the selection was created, delete it
	If Not objSS Is Nothing Then
	objSS.Delete
	End If
End Sub

Public Sub TestSelectOnScreen()
	Dim objSS As AcadSelectionSet
	On Error GoTo Done
	With ThisDrawing.Utility
	'' create a new selectionset
	Set objSS = ThisDrawing.SelectionSets.Add("TestSelectOnScreen")
	'' let user select entities interactively
	objSS.SelectOnScreen
	'' highlight the selected entities
	objSS.Highlight True
	'' pause for the user
	.Prompt vbCr & objSS.Count & " entities selected"
	.GetString False, vbLf & "Enter to continue "
	'' unhighlight the entities
	objSS.Highlight False
	End With
	Done:
	'' if the selection was created, delete it
	If Not objSS Is Nothing Then
	objSS.Delete
	End If
End Sub

Public Sub TestGroupDelete()
	Dim objGroup As AcadGroup
	Dim strName As String
	On Error Resume Next
	With ThisDrawing.Utility
	strName = .GetString(True, vbCr & "Group name: ")
	If Err Or "" = strName Then Exit Sub
	'' get the existing group
	Set objGroup = ThisDrawing.Groups.Item(strName)
	If Err Then
	.Prompt vbCr & "Group does not exist "
	Exit Sub
	End If
	'' delete the group
	objGroup.Delete
	If Err Then
	.Prompt vbCr & "Error deleting group "
	Exit Sub
	End If
	'' pause for the user
	.Prompt vbCr & "Group deleted"
	End With
End Sub

Public Sub TestSelect()
	Dim objSS As AcadSelectionSet
	Dim varPnt1 As Variant
	Dim varPnt2 As Variant
	Dim strOpt As String
	Dim lngMode As Long
	On Error GoTo Done
	With ThisDrawing.Utility
	'' get input for mode
	.InitializeUserInput 1, "Window Crossing Previous Last All"
	strOpt = .GetKeyword(vbCr & _
	"Select [Window/Crossing/Previous/Last/All]: ")
	'' convert keyword into mode
	Select Case strOpt
	Case "Window": lngMode = acSelectionSetWindow
	Case "Crossing": lngMode = acSelectionSetCrossing
	Case "Previous": lngMode = acSelectionSetPrevious
	Case "Last": lngMode = acSelectionSetLast
	Case "All": lngMode = acSelectionSetAll
	End Select
	'' create a new selectionset
	Set objSS = ThisDrawing.SelectionSets.Add("TestSelectSS")
	'' if it's window or crossing, get the points
	If "Window" = strOpt Or "Crossing" = strOpt Then
	'' get first point
	.InitializeUserInput 1
	varPnt1 = .GetPoint(, vbCr & "Pick the first corner: ")
	'' get corner, using dashed lines if crossing
	.InitializeUserInput 1 + IIf("Crossing" = strOpt, 32, 0)
	varPnt2 = .GetCorner(varPnt1, vbCr & "Pick other corner: ")
	'' select entities using points
	objSS.Select lngMode, varPnt1, varPnt2
	Else
	'' select entities using mode
	objSS.Select lngMode
	End If
	'' highlight the selected entities
	objSS.Highlight True
	'' pause for the user
	.GetString False, vbCr & "Enter to continue"
	'' unhighlight the entities
	objSS.Highlight False
	End With
	Done:
	'' if the selectionset was created, delete it
	If Not objSS Is Nothing Then
	objSS.Delete
	End If
End Sub

'13.BlocksAttributes



Sub ExtractAtts()
	Dim Excel As Excel.Application
	Dim ExcelSheet As Object
	Dim ExcelWorkbook As Object
	Dim RowNum As Integer
	Dim Header As Boolean
	Dim elem As AcadEntity
	Dim Array1 As Variant
	Dim Count As Integer
	' Launch Excel.
	Set Excel = New Excel.Application
	' Create a new workbook and find the active sheet.
	Set ExcelWorkbook = Excel.Workbooks.Add
	Set ExcelSheet = Excel.ActiveSheet
	ExcelWorkbook.SaveAs "Attribute.xls"
	RowNum = 1
	Header = False
	' Iterate through model space finding
	' all block references.
	For Each elem In ThisDrawing.ModelSpace
	With elem
	' When a block reference has been found,
	' check it for attributes
	If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
	If .HasAttributes Then
	' Get the attributes
	Array1 = .GetAttributes
	' Copy the Tagstrings for the
	' Attributes into Excel
	For Count = LBound(Array1) To UBound(Array1)
	If Header = False Then
	If StrComp(Array1(Count).EntityName, "AcDbAttribute", 1) = 0 Then
	ExcelSheet.Cells(RowNum, Count + 1).Value = Array1(Count).TagString
	End If
	End If
	Next Count
	RowNum = RowNum + 1
	For Count = LBound(Array1) To UBound(Array1)
	ExcelSheet.Cells(RowNum, Count + 1).Value = Array1(Count).TextString
	Next Count
	Header = True
	End If
	End If
	End With
	Next elem
	Excel.Application.Quit
End Sub

sub myBlock()
	' Get the position of the Solid inside the block definition
	Dim blk As AcadBlock
	Set blk = ThisDrawing.Blocks("myBlock")
	Dim ent As AcadEntity
	Dim position(0 To 2) As Double
	For Each ent In blk
	If TypeOf ent Is AcadSolid And [some other condition to make sure it is the target, if
	necessary] Then
	position = [Get a set of coordinate from the Solid, such as one of its corner...]
	Exit For
	End If
	Next
	'' Get the positions related to the drawing's 0, 0,0 (such as ModelSpace's World Origin:
	Dim bref As AcadBlockReference
	Dim pos() As Variant
	Dim i As Ineter
	For Each ent In ThisDrawing.ModelSpace
	If TypeOf ent Is AcadBlockrefernece Then
	Set bref = ent
	If UCase(bref.EffectiveName) = UCase("myBlock") Then
	ReDim pos(i)
	pos(i)(0) = bref.InsertionPoint(0) + position(0)
	pos(i)(1) = bref.InsertionPoint(1) + position(1)
	pos(i)(2) = bref.InsertionPoint(2) + position(2)
	i = i + 1
	End If
	End If
	Next

end sub

Sub Ch10_AttachingExternalReference()
	On Error GoTo ERRORHANDLER
	Dim InsertPoint(0 To 2) As Double
	Dim insertedBlock As AcadExternalReference
	Dim tempBlock As AcadBlock
	Dim msg As String, PathName As String
	' Define external reference to be inserted
	InsertPoint(0) = 1
	InsertPoint(1) = 1
	InsertPoint(2) = 0
	PathName = "C:/Program Files/Autodesk/AutoCAD release/sample/3D House.dwg"
	' Display current Block information for this drawing
	GoSub ListBlocks
	' Add the external reference to the drawing
	Set insertedBlock = ThisDrawing.ModelSpace. _
	AttachExternalReference(PathName, "XREF_IMAGE", _
	InsertPoint, 1, 1, 1, 0, False)
	ZoomAll
	' Display new Block information for this drawing
	GoSub ListBlocks
	Exit Sub
	ListBlocks:
	msg = vbCrLf ' Reset message
	For Each tempBlock In ThisDrawing.Blocks
	msg = msg & tempBlock.Name & vbCrLf
	Next
	MsgBox "The current blocks in this drawing are: " & msg
	Return
	ERRORHANDLER:
	MsgBox Err.Description
End Sub

Option Explicit
' S:\Everyone\ENGR_REQUESTS\AutoREV\Data\GateWay1.xls
'Global vars
Private excelApp As Excel.Application 'points to excel application
Private wbkObj As Workbook 'points to excel workbook
Private rSearch As Range 'Range where the search is performed
Private rFound As Range 'Range where the data is found
Private dwginfo As Collection 'holds the "found" info

Public Function AddSelectionSet(SetName As String) As AcadSelectionSet
	' This routine does the error trapping neccessary for when you want to create a
	' selectin set. It takes the proposed name and either adds it to the selectionsets
	' collection or sets it.
	On Error Resume Next
	Set AddSelectionSet = ThisDrawing.SelectionSets.Add(SetName)
	If Err.Number <> 0 Then
	Set AddSelectionSet = ThisDrawing.SelectionSets.item(SetName)
	AddSelectionSet.Clear
	End If
End Function
Public Sub GetTitleBlockInfo(PrjNo As String)
	On Error GoTo Err_Control
	Set dwginfo = New Collection
	With rSearch
	Set rFound = .Find(What:=PrjNo, _
	LookIn:=xlValues, _
	LookAt:=xlWhole, _
	SearchOrder:=xlByRows, _
	SearchDirection:=xlNext, _
	MatchCase:=False)
	If Not rFound Is Nothing Then
	dwginfo.Add rFound.Offset(, 1).Value, "SALES_ORDER"
	dwginfo.Add rFound.Offset(, 3).Value, "CUSTOMER"
	dwginfo.Add rFound.Offset(, 4).Value, "CITY"
	dwginfo.Add rFound.Offset(, 5).Value, "STATE"
	dwginfo.Add rFound.Offset(, 12).Value, "STORE_NAME"
	Else
	Err.Raise vbObjectError + 101
	End If
	End With
	Exit_Here:
	Exit Sub
	Err_Control:
	Select Case Err.Number
	Case Is = 101
	'Search Item not found.
	'Pass them up to calling sub.
	Err.Raise vbObjectError + 101, "Module1.GetTitleBlockInfo", "Search Item not found."
	Resume Exit_Here
	Case Else
	'Handle unforseen errors.
	'Pass them up to calling sub.
	Err.Raise vbObjectError + 100, "Module1.GetTitleBlockInfo"
	Resume Exit_Here
	End Select
End Sub

Public Function GetExcel() As Excel.Application
	On Error GoTo Err_Control
	Dim m_app As Excel.Application
	Set m_app = GetObject(, "Excel.Application")
	Return_App:
	Set GetExcel = m_app
	Exit_Here:
	Exit Function
	Err_Control:
	Select Case Err.Number
	Case Is = 429
	'Excel is not running. Start it.
	Set m_app = CreateObject("Excel.Application")
	Resume Return_App
	Case Else
	'Handle unforseen errors.
	MsgBox Err.Number & ", " & Err.Description, , "GetExcel"
	Err.Clear
	Resume Exit_Here
	End Select
End Function

Public Function GetSS_BlockName(BlockName As String) As AcadSelectionSet
	'creates a ss of blocks with the name supplied in the argument
	Dim s2 As AcadSelectionSet
	Set s2 = AddSelectionSet("ssBlocks") ' create ss with a name
	s2.Clear ' clear the set
	Dim intFtyp(3) As Integer ' setup for the filter
	Dim varFval(3) As Variant
	Dim varFilter1, varFilter2 As Variant
	intFtyp(0) = -4: varFval(0) = "<AND"
	intFtyp(1) = 0: varFval(1) = "INSERT" ' get only blocks
	intFtyp(2) = 2: varFval(2) = BlockName ' whose name is specified in argument
	intFtyp(3) = -4: varFval(3) = "AND>"
	varFilter1 = intFtyp: varFilter2 = varFval
	s2.Select acSelectionSetAll, , , varFilter1, varFilter2 ' do it
	Set GetSS_BlockName = s2
End Function

Public Sub UpdateTitleblock()
	Dim ent As Object
	On Error GoTo Err_Control
	'Open excel
	Set excelApp = GetExcel()
	Set wbkObj =
	excelApp.Workbooks.Open("S:\Everyone\ENGR_REQUESTS\AutoREV\Data\GateWay1.xls")
	Set rSearch = wbkObj.Worksheets(1).Range("A:A")
	GetTitleBlockInfo CLng(Left(ThisDrawing.Name, 7))
	'Update DWG
	Dim ss As AcadSelectionSet
	Dim blk As AcadBlockReference
	Set ss = GetSS_BlockName("TitleBlock")
	Set blk = ss(0)
	If blk.HasAttributes = True Then
	-2-D:\_S\Excel workbook\CAD_Col\13.BlocksAttributes\mk_AttachTitleBlock.bas Thursday, May 20, 2021 8:20 PM
	Dim x As Long
	Dim attArr As Variant
	Dim att As AcadAttributeReference
	attArr = blk.GetAttributes
	For x = 0 To UBound(attArr)
	Set att = attArr(x)
	Select Case att.TagString
	Case Is = "SALES_ORDER"
	att.TextString = dwginfo("SALES_ORDER")
	Case Is = "CUSTOMER"
	att.TextString = dwginfo("CUSTOMER")
	Case Is = "CITY"
	att.TextString = dwginfo("CITY")
	Case Is = "STATE"
	att.TextString = dwginfo("STATE")
	Case Is = "STORE_NAME"
	att.TextString = dwginfo("STORE_NAME")
	End Select
	Next
	End If
	Cleanup:
	'Cleanup out-of-process object, in reverse order of creation.
	excelApp.Quit
	Set rFound = Nothing
	Set rSearch = Nothing
	Set wbkObj = Nothing
	Set excelApp = Nothing
	Exit_Here:
	Exit Sub
	Err_Control:
	Select Case Err.Number
	Case Is = 1004
	'File not found.
	MsgBox "File not found." & vbCrLf & Err.Number & ", " & Err.Description, , Err.Source
	Err.Clear
	Resume Cleanup
	Case Is = vbObjectError + 100
	'Unhandled error in GetTitleBlockInfo
	MsgBox "Unhandled Error in GetTitleBlockInfo(): " & Err.Number & ", " &
	Err.Description, , Err.Source
	Err.Clear
	Resume Cleanup
	Case Is = vbObjectError + 101
	'File not found.
	MsgBox "Project Number was not found in Excel spreadsheet.", , Err.Source
	Err.Clear
	Resume Cleanup
	Case Else
	'Handle unforseen errors.
	MsgBox Err.Number & ", " & Err.Description, , "UpdateTitleblock"
	Err.Clear
	Resume Cleanup
	End Select
End Sub

Sub GetInsertpoint()
	Dim oEnt As AcadEntity
	Dim varPick As Variant
	Dim brBref As AcadBlockReference
	Dim arAttR As AcadAttributeReference
	Dim varAt As Variant
	Dim i As Double
	ThisDrawing.Utility.GetEntity oEnt, varPick, vbCr & "Get the block"
	If TypeOf oEnt Is AcadBlockReference Then
	MsgBox "Thank you, very nice!"
	Set brBref = oEnt
	MsgBox brBref.InsertionPoint(0) & brBref.InsertionPoint(1) & brBref.InsertionPoint(2)
	Else
	MsgBox "Not a block reference!"
	Exit Sub
	End If
End Sub

Option Explicit

Private Sub ScanBlock()
	Dim strString As String
	Dim obj As Object
	Dim obj2 As Object
	'On Error Resume Next
	For Each obj In ThisDrawing.PaperSpace
	If obj.ObjectName = "AcDbBlockReference" Then
	If obj.Name = "tbD2" Then
	For Each obj2 In obj
	If obj2.ObjectName = ?AcDbMtext? Then
	strString = obj2.Value
	MsgBox strString
	'Return strString
	End If
	Next obj2
	End If
	Next Object
End Sub

	''' try this from other guy..
	Dim objtext As AcadText
	Dim objBlock As AcadBlockReference
	For Each obj In ThisDrawing.PaperSpace
	If TypeOf obj Is AcDbBlockReference And obj.Name = "tbD2" Then
	Set objBlock = obj
	For Each obj2 In objBlock
	If TypeOf obj2 Is AcadText Then
	Set objtext = obj2
	strString = objtext.TextString
	MsgBox strString
	'''' here is the final or...


Private Sub ScanBlock()
	Dim strString As String
	Dim obj As AcadBlock
	Dim obj2 As AcadEntity
	For Each obj In ThisDrawing.Blocks
	If obj.Name = "tbD2" Then
	For Each obj2 In obj
	If obj2.ObjectName = "AcDbMText" Then 'this was the problem, notice the capital T
	strString = obj2.TextString
	MsgBox strString
	'Return strString
	End If
	Next obj2
	End If
	Next obj
End Sub

Sub PlaceInExcel(iStrInfo As String)
	Dim xL As Excel.Application
	Dim xLWrkBk As Workbook
	Dim xLSheet As Worksheet
	' this will create a new instance every time you may want to check for an existing instance
	Set xL = CreateObject("Excel.Application")
	xL.Visible = True
	Set xLWrkBk = xL.Workbooks.Add
	Set xLSheet = xLWrkBk.Sheets("Sheet1")
	xLSheet.Cells(1, 1).Value = iStrInfo
	-1-D:\_S\Excel workbook\CAD_Col\13.BlocksAttributes\mk_BlockScan.bas Thursday, May 20, 2021 8:20 PM
	' this leaves excel open and active if you need to do something else let me know
	Set xLSheet = Nothing
	Set xLWrkBk = Nothing
	Set xL = Nothing
End Sub


Sub FIRST_TEST()
	Dim MyACAD As String
	Dim A2K As AcadApplication
	Dim A2Kdwg As AcadDocument
	Dim lBlockInsertionPoint(0 To 2) As Double
	Dim TxtObj As AcadText
	Dim TxtStr As String
	Dim BlockObj, BlockObj1, BlockObj2 As AcadBlock
	MyACAD = "AutoCAD.Application.16" '"AutoCAD.Application.16" = AutoCAD 2004
	On Error Resume Next
	Err.Clear
	With ACAD ' DEFINED IN MODULE
	Set .Application = CreateObject(MyACAD)
	If Err.Number = 429 Then
	Dim AskToActivateACAD
	AskToActivateACAD = MsgBox("Want to activate AutoCAD?", vbQuestion + vbYesNo,
	"AppCaption")
	If AskToActivateACAD <> vbYes Then
	GoTo JumpSpot
	End If
	End If
	If Err <> 0 And Err.Number = 429 Then
	Err.Clear
	Set ACAD.Application = CreateObject(MyACAD)
	If Err <> 0 Then
	MsgBox "AutoCAD is not installed." & vbCr & vbCr & _
	"or" & vbCr & vbCr & _
	"you are using AutoCAD LT", vbInformation + vbOKOnly, "AppCaption"
	GoTo JumpSpot
	End If
	.Application.Visible = True
	.Application.WindowState = 3
	End If
	.DocumentCount = .Application.Documents.Count
	If .DocumentCount < 1 Then
	Set .ActiveDoc = .Application.Documents.Add
	.DocumentCount = 1
	Else
	Set .ActiveDoc = .Application.ActiveDocument
	End If
	.LayoutCount = .ActiveDoc.Layouts.Count
	Set .ActiveLayout = .ActiveDoc.ActiveLayout
	.LayerCount = .ActiveDoc.Layers.Count
	Set .ActiveLayer = .ActiveDoc.ActiveLayer
	JumpSpot:
	'Add a new document.
	Set A2Kdwg = .Application.Documents.Add
	'Add a new document. Gather its number
	iNewDocument = .Application.Documents.Count - 1 ' 1, 2, 3 in watch gives 0, 1, 2 in
	array.
	'A2Kdwg shows as 'nothing' in the watch window!
	Set A2Kdwg = .Application.Documents.Item(iNewDocument)
	MsgBox ACAD.Application.Name & " version " & ACAD.Application.Version & " is running."
	'The file holds the block that I want to add
	strBlockFile = "C:\COM\RAW\AI-2W-1H-SB.DWG"
	strBlockName = FileNameNoExt(CStr(strBlockFile))
	lBlockInsertionPoint(0) = 1: lBlockInsertionPoint(1) = 1: lBlockInsertionPoint(2) = 0
	' INSERT BLOCK
	Set BlockObj2 =
	.Application.Documents.Item(iNewDocument).ModelSpace.InsertBlock(lBlockInsertionPoint,
	-1-D:\_S\Excel workbook\CAD_Col\13.BlocksAttributes\mk_BlockTextManipulation.bas Thursday, May 20, 2021 8:20 PM
	strBlockFile, 1#, 1#, 1#, 0)
	'Count the blocks in the drawing...
	iMyBlockItemCount = .Application.Documents.Item(iNewDocument).Blocks.Count - 1
	'Search through the blocks for the block you just added.
	For iBlockHunt = 0 To iMyBlockItemCount
	strSourceBlockName =
	.Application.Documents.Item(iNewDocument).Blocks.Item(iBlockHunt).Name
	If strSourceBlockName = strBlockName Then
	'MsgBox "Found the block we just added"
	'Count the items in this block.
	iItemCount =
	.Application.Documents.Item(iNewDocument).Blocks.Item(iBlockHunt).Count - 1
	'For every item in the block you just added...
	For iThisItem = 0 To iItemCount
	strThisLayer =
	.Application.Documents.Item(iNewDocument).Blocks.Item(iBlockHunt).Item(iThisI
	tem).Layer
	strTest = .Application.Documents.Item(1).Blocks.Item(8).Item(0).Layer
	fTest = .Application.Documents.Item(1).Blocks.Item(8).Item(0).Length
	'Layer "0" holds the items we want to change.
	If "0" = strThisLayer Then
	'Stop
	'Assign local variables to be the values we want to change.
	'This works well, the changed values from the ACAD block show up in the
	watch window.
	MyPromptString =
	.Application.Documents.Item(iNewDocument).Blocks.Item(iBlockHunt).Item(iT
	hisItem).PromptString
	MyTagString =
	.Application.Documents.Item(iNewDocument).Blocks.Item(iBlockHunt).Item(iT
	hisItem).TagString
	MyTextString =
	.Application.Documents.Item(iNewDocument).Blocks.Item(iBlockHunt).Item(iT
	hisItem).TextString
	'This works strangely!
	'I CAN assign a PROMPTSTRING. The Watch window shows the change, AND
	the resulting ACAD file shows the change.
	.Application.Documents.Item(iNewDocument).Blocks.Item(iBlockHunt).Item(iT
	hisItem).PromptString = "[PROMPT_" & iThisItem & "]"
	'I CANNOT assign a TAG STRING... The Watch window shows the change, but
	the resulting ACAD file does not show the change.
	.Application.Documents.Item(iNewDocument).Blocks.Item(iBlockHunt).Item(iT
	hisItem).TagString = "[TAG_" & iThisItem & "]"
	'I CANNOT assign a TEXT STRING... The Watch window shows the change,
	but the resulting ACAD file does not show the change.
	.Application.Documents.Item(iNewDocument).Blocks.Item(iBlockHunt).Item(iT
	hisItem).TextString = "[TEXT_" & iThisItem & "]"
	.Application.Documents.Item(iNewDocument).Regen acAllViewports
	'SAVE A COPY AFTER EACH CHANGE... TO SEE HOW THE CHANGE PROGRESSES
	.Application.Documents.Item(iNewDocument).SaveAs "C:\COM\___RSTest" &
	iThisItem & ".dwg"
	End If
	Next iThisItem
	End If
	Next iBlockHunt
	.Application.Documents.Item(iNewDocument).SaveAs "C:\COM\___RSTest.dwg"
	ACAD.Application.Documents.Close
	ACAD.Application.Quit
	Set ACAD.Application = Nothing
	End With
	-2-D:\_S\Excel workbook\CAD_Col\13.BlocksAttributes\mk_BlockTextManipulation.bas Thursday, May 20, 2021 8:20 PM
End Sub


Sub ChangeBlockNames(Dwg As AcadDocument)
	Dim BlockDef As AcadBlock
	For Each BlockDef In Dwg.Blocks
	' BlockDef.Name = New block name here...
	Next
End Sub

Private Sub AcadDocument_BeginCommand(ByVal CommandName As String)
	NoRun = True
End Sub

Private Sub AcadDocument_EndCommand(ByVal CommandName As String)
	NoRun = False
End Sub

Private Sub AcadDocument_BeginSave(ByVal FileName As String)
	NoRun = True
End Sub

Private Sub AcadDocument_EndSave(ByVal FileName As String)
	NoRun = False
End Sub

Private Sub AcadDocument_ObjectModified(ByVal Object As Object)
	If NoRun Then Exit Sub ' if there are othere activities do not call Sub
	' If the Modified Object is a dynamic block then call the Sub
	If Object.ObjectName = "AcDbAttribute" And LastID <> Object.OwnerID32 Then
	' Save the Owner ID so the Sub is only called once for the Block Modification
	LastID = Object.OwnerID32
	Call ModifiedObject(Object, LastID)
	End If
End Sub

Option Explicit

Public Sub ModifiedObject(ByVal Object As Object, LastID As Variant)
	Dim ParObj As AcadObject
	Dim obj As AcadObject
	Dim BlkRef As AcadBlockReference
	Dim varAttributes As Variant
	Dim DBRP As Variant
	Set obj = Object
	Set ParObj = ThisDrawing.ObjectIdToObject32(obj.OwnerID32)
	' Check the object's parent type (Dynamic Block Reference)
	If ParObj.EntityName = "AcDbBlockReference" Then
	' change object type to get Block Reference Attributes
	' and test for the correct Dynamic Block Name
	Set BlkRef = ParObj
	If BlkRef.EffectiveName = "SW_INPUT" Then
	' Get the attributes for the block reference
	varAttributes = BlkRef.GetAttributes
	' If the Instrument Tag attribute hasn't changed leave the sub
	If varAttributes(0).TextString = "INSTRUMENT_TAG" Then
	Set obj = Nothing
	Set ParObj = Nothing
	Set BlkRef = Nothing
	LastID = ""
	Exit Sub
	End If
	varAttributes(4).TextString = ""
	' Set the value of of the dynamic block attribute to
	' the code for the switch type
	DBRP = BlkRef.GetDynamicBlockProperties
	-1-D:\_S\Excel workbook\CAD_Col\13.BlocksAttributes\mk_DynamicBlocksEdit.bas Thursday, May 20, 2021 8:20 PM
	DBRP(0).Value = GetSwitchType(varAttributes(0).TextString)
	End If
	End If
	' Clean up before exiting the Sub
	Set obj = Nothing
	Set ParObj = Nothing
	Set BlkRef = Nothing
End Sub

Private Function GetSwitchType(ByVal TagName As String) As String
	Dim RetVal, Fst, Trd, WkStr As String
	WkStr = Left(TagName, 3)
	Fst = UCase(Left(WkStr, 1))
	Trd = UCase(Right(WkStr, 1))
	Debug.Print TagName & " = " & Fst & ", " & Trd
	Select Case Fst
	Case "P" ' Test for a Pressure Switch
	If Trd = "L" Then
	GetSwitchType = "A1"
	Else
	GetSwitchType = "A2"
	End If
	Case "T" ' Test for a Temperature Switch
	If Trd = "L" Then
	GetSwitchType = "B1"
	Else
	GetSwitchType = "B2"
	End If
	Case "L" ' Test for a Level Switch
	If Trd = "L" Then
	GetSwitchType = "C1"
	Else
	GetSwitchType = "C2"
	End If
	Case "F" ' Test for a Flow Switch
	If Trd = "L" Then
	GetSwitchType = "D1"
	Else
	GetSwitchType = "D2"
	End If
	Case "Z" ' Test for a Position Switch
	If Trd = "C" Then
	GetSwitchType = "E1"
	Else
	GetSwitchType = "E2"
	End If
	Case Else ' If the type doesn't match anything else use contacts
	If Trd = "L" Then
	GetSwitchType = "F1"
	Else
	GetSwitchType = "F2"
	End If
	End Select
	-2-D:\_S\Excel workbook\CAD_Col\13.BlocksAttributes\mk_DynamicBlocksEdit.bas Thursday, May 20, 2021 8:20 PM
End Function

Public Sub ReplaceSomeBlocks()
	Dim objBlock As AcadBlock
	Dim objBRef As AcadBlockReference
	Dim objNewBRef As AcadBlockReference
	Dim objAttTag As AcadAttributeReference
	Dim intTagCnt As Integer
	Dim varAttList As Variant
	Dim strRmNo As String
	Dim strNewName As String
	Dim intcodes(0) As Integer
	Dim varCodeValues(0) As Variant
	Dim varInPt As Variant
	Dim dblX As Double
	Dim dblY As Double
	Dim dblZ As Double
	Dim dblRotation As Double
	Dim lngMode As Long
	Dim objSS As AcadSelectionSet
	intcodes(0) = 2
	varCodeValues(0) = "VanillaBlock"
	strNewName = "MyDynamicBlock"
	lngMode = acSelectionSetAll
	intTagCnt = 2
	On Error Resume Next
	ThisDrawing.SelectionSets.Item("TempSSet").Delete
	Set objSS = ThisDrawing.SelectionSets.Add("TempSSet")
	objSS.Select acSelectionSetAll, , , intcodes, varCodeValues
	MsgBox objSS.Count
	For Each objBRef In objSS
	If objBRef.HasAttributes Then
	varAttList = objBRef.GetAttributes
	strRmNo = varAttList(intTagCnt).TextString
	End If
	dblRotation = objBRef.Rotation
	dblX = objBRef.XScaleFactor
	dblY = objBRef.YScaleFactor
	dblZ = objBRef.ZScaleFactor
	varInPt = objBRef.InsertionPoint
	ThisDrawing.ModelSpace.InsertBlock(varInPt, strNewName, dblX, dblY, dblZ, dblRotation)
	objBRef.Delete
	Next objBRef
	ThisDrawing.SelectionSets.Item("TempSSet").Delete
End Sub

Public Sub ReplaceSomeBlocks()
	Dim objBlock As AcadBlock
	Dim objBRef As AcadBlockReference
	Dim objNewBRef As AcadBlockReference
	Dim objAttTag As AcadAttributeReference
	Dim intTagCnt As Integer
	Dim varAttList As Variant
	Dim strRmNo As String
	Dim strNewName As String
	Dim intcodes(0) As Integer
	Dim varCodeValues(0) As Variant
	Dim varInPt As Variant
	Dim dblX As Double
	Dim dblY As Double
	Dim dblZ As Double
	Dim dblRotation As Double
	Dim lngMode As Long
	Dim objSS As AcadSelectionSet
	intcodes(0) = 2
	varCodeValues(0) = "VanillaBlock"
	strNewName = "MyDynamicBlock"
	lngMode = acSelectionSetAll
	intTagCnt = 2
	On Error Resume Next
	ThisDrawing.SelectionSets.Item("TempSSet").Delete
	Set objSS = ThisDrawing.SelectionSets.Add("TempSSet")
	objSS.Select acSelectionSetAll, , , intcodes, varCodeValues
	MsgBox objSS.Count
	For Each objBRef In objSS
	If objBRef.HasAttributes Then
	varAttList = objBRef.GetAttributes
	strRmNo = varAttList(intTagCnt).TextString
	End If
	dblRotation = objBRef.Rotation
	dblX = objBRef.XScaleFactor
	dblY = objBRef.YScaleFactor
	dblZ = objBRef.ZScaleFactor
	varInPt = objBRef.InsertionPoint
	ThisDrawing.ModelSpace.InsertBlock(varInPt, strNewName, dblX, dblY, dblZ, dblRotation)
	objBRef.Delete
	Next objBRef
	ThisDrawing.SelectionSets.Item("TempSSet").Delete
End Sub

Sub Ch10_ExplodingABlock()
	' Define the block
	Dim blockObj As AcadBlock
	Dim insertionPnt(0 To 2) As Double
	insertionPnt(0) = 0
	insertionPnt(1) = 0
	insertionPnt(2) = 0
	Set blockObj = ThisDrawing.Blocks.Add _
	(insertionPnt, "CircleBlock")
	' Add a circle to the block
	Dim circleObj As AcadCircle
	Dim center(0 To 2) As Double
	Dim radius As Double
	center(0) = 0
	center(1) = 0
	center(2) = 0
	radius = 1
	Set circleObj = blockObj.AddCircle(center, radius)
	' Insert the block
	Dim blockRefObj As AcadBlockReference
	insertionPnt(0) = 2
	insertionPnt(1) = 2
	insertionPnt(2) = 0
	Set blockRefObj = ThisDrawing.ModelSpace.InsertBlock _
	(insertionPnt, "CircleBlock", 1#, 1#, 1#, 0)
	ZoomAll
	MsgBox "The circle belongs to " & blockRefObj.ObjectName
	' Explode the block reference
	Dim explodedObjects As Variant
	explodedObjects = blockRefObj.Explode
	' Loop through the exploded objects
	Dim I As Integer
	For I = 0 To UBound(explodedObjects)
	explodedObjects(I).color = acRed
	explodedObjects(I).Update
	MsgBox "Exploded Object " & I & ": " _
	& explodedObjects(I).ObjectName
	explodedObjects(I).color = acByLayer
	explodedObjects(I).Update
	Next
End Sub

Sub Change_All_existing_To_Layer()
	Dim ssAll As AcadSelectionSet
	Dim mEntity As AcadEntity, oBlkRef As AcadEntity
	Dim blkAtt As String
	Dim ssets As AcadSelectionSets
	'clear selection set
	Set ssets = ThisDrawing.SelectionSets
	On Error Resume Next
	Set ssAll = ssets.Item("AU2013_SS")
	If Err.Number <> 0 Then
	Set ssAll = ssets.Add("AU2013_SS")
	Else
	ssAll.Clear
	End If
	'selects all entries in drawing
	Set ssAll = ThisDrawing.SelectionSets.Add("AllEntities")
	ssAll.Select acSelectionSetAll
	For Each mEntity In ssAll 'loops thru entry list
	If TypeOf mEntity Is AcadBlockReference Then 'picks blocks
	Set oBlkRef = mEntity
	If oBlkRef.HasAttributes Then 'picks blocks with attributes
	blkAtt = Get_Attribute(oBlkRef, "BLKSET.FURNISH")
	If blkAtt = "\W0.7500;E" Then 'picks "existing" devices
	mEntity.Layer = "STAGE.NOTES.SD"
	End If
	End If
	End If
	Next
	ssAll.Clear 'cleans the selection set
	ssAll.Delete
	Set ssAll = Nothing
	ThisDrawing.Regen acActiveViewport 'regenerates the drawing
End Sub

Sub blocktoanother()
	For i = 1 To AddHole
	dwgInsert = gsGetPath(objVBE.ActiveVBProject.FileName) &
	"\Blocks\FrameProtector\HoleOS43.dwg"
	InsertionPoint(0) = SP1 + (HoleGap * i): InsertionPoint(1) = SP2: InsertionPoint(2) = 0#
	Set blockObj = ThisDrawing.ModelSpace.InsertBlock(InsertionPoint, dwgInsert, 1#, 1#,
	1#, 0)
	Next
	dwgInsert = gsGetPath(objVBE.ActiveVBProject.FileName) &
	"\Blocks\FrameProtector\FPDOUBLEXXXX.dwg"
	InsertionPoint(0) = SP1: InsertionPoint(1) = SP2: InsertionPoint(2) = 0#
	Set blockObj = ThisDrawing.ModelSpace.InsertBlock(InsertionPoint, dwgInsert, 1#, 1#, 1#, 0)
	GetAtt = blockObj.GetAttributes
	For vari = 0 To UBound(GetAtt)
	Select Case GetAtt(vari).TagString
	Case "ITEMCODE"
	GetAtt(vari).TextString = CodeHeight & CodeDirection & CodeLength
	Case "QTY"
	GetAtt(vari).TextString = "1"
	Case "VIEW"
	GetAtt(vari).TextString = "PV"
	Case "CREATEDBY"
	GetAtt(vari).TextString = "SYSTEM"
	Case "PRODUCTTYPE"
	GetAtt(vari).TextString = "FRAME PROTECTOR"
	Case "PRODUCTCATEGORY"
	GetAtt(vari).TextString = "i600"
	End Select
	Next
	dynVar = blockObj.GetDynamicBlockProperties
	For counter = LBound(dynVar) To UBound(dynVar)
	If dynVar(counter).PropertyName = "length" Then
	dynVar(counter).Value = CDbl(Length) - 500
	End If
	Next counter
	Quick reply to this message
End Sub

Public Function pfFastenerLabel()
	Dim SS As AcadSelectionSet
	Dim BlkRef As AcadBlockReference
	Dim FType(0) As Integer
	Dim FData(0) As Variant
	FType(0) = 0
	FData(0) = "INSERT"
	On Error Resume Next
	Set SS = ThisDrawing.SelectionSets.Add("FSSS")
	Set SS = ThisDrawing.SelectionSets.Item("FSSS")
	On Error GoTo 0
	SS.Clear
	SS.SelectOnScreen FType, FData
	For Each BlkRef In SS
	If BlkRef.EffectiveName Like "db*" Then
	MsgBox BlkRef.Name
	End If
	Next BlkRef
	SS.Delete
End Function

Sub temp()
	Dim obj As AcadEntity
	For Each obj In ThisDrawing.ModelSpace
	If obj.ObjectName = "AcDbBlockReference" Then
	tmp = obj.EffectiveName
	If InStr(1, tmp, "*U", vbTextCompare) Then
	Debug.Print tmp
	End If
	End If
	Next obj
End Sub

Public Sub TestExplode()
	Dim objBRef As AcadBlockReference
	Dim varPick As Variant
	Dim varNew As Variant
	Dim varEnts As Variant
	Dim intI As Integer
	On Error Resume Next
	'' get an entity and new point from user
	With ThisDrawing.Utility
	.GetEntity objBRef, varPick, vbCr & "Pick a block reference: "
	If Err Then Exit Sub
	varNew = .GetPoint(varPick, vbCr & "Pick a new location: ")
	If Err Then Exit Sub
	End With
	'' explode the blockref
	varEnts = objBRef.Explode
	If Err Then
	MsgBox "Error has occurred: " & Err.Description
	Exit Sub
	End If
	'' move resulting entities to new location
	For intI = 0 To UBound(varEnts)
	varEnts(intI).Move varPick, varNew
	Next
End Sub

Private Sub CommandButton1_Click()
	Dim objBlockRef As AcadBlockReference
	Dim varInsertionPoint As Variant
	Dim dblX As Double
	Dim dblY As Double
	Dim dblZ As Double
	Dim dblRotation As Double
	'' get input from user
	dlgOpenFile.Filter = "AutoCAD Blocks (*.DWG) | *.dwg"
	dlgOpenFile.InitDir = Application.Path
	dlgOpenFile.ShowOpen
	If dlgOpenFile.FileName = "" Then Exit Sub
	Me.Hide
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varInsertionPoint = .GetPoint(, vbCr & "Pick the insert point: ")
	.InitializeUserInput 1 + 2
	dblX = .GetDistance(varInsertionPoint, vbCr & "X scale: ")
	.InitializeUserInput 1 + 2
	dblY = .GetDistance(varInsertionPoint, vbCr & "Y scale: ")
	.InitializeUserInput 1 + 2
	dblZ = .GetDistance(varInsertionPoint, vbCr & "Z scale: ")
	.InitializeUserInput 1
	dblRotation = .GetAngle(varInsertionPoint, vbCr & "Rotation angle: ")
	End With
	'' create the object
	On Error Resume Next
	Set objBlockRef = ThisDrawing.ModelSpace.InsertBlock _
	(varInsertionPoint, dlgOpenFile.FileName, dblX, _
	dblY, dblZ, dblRotation)
	If Err Then
	MsgBox "Unable to insert this block"
	Exit Sub
	End If
	objBlockRef.Update
	Me.Show
End Sub

Public Sub TestWBlock()
	Dim objSS As AcadSelectionSet
	Dim varBase As Variant
	Dim dblOrigin(2) As Double
	Dim objEnt As AcadEntity
	Dim strFilename As String
	'choose a selection set name that you use only as temporary storage and
	'ensure that it does not currently exist
	On Error Resume Next
	ThisDrawing.SelectionSets("TempSSet").Delete
	Set objSS = ThisDrawing.SelectionSets.Add("TempSSet")
	objSS.SelectOnScreen
	With ThisDrawing.Utility
	.InitializeUserInput 1
	strFilename = .GetString(True, vbCr & "Enter a filename: ")
	.InitializeUserInput 1
	varBase = .GetPoint(, vbCr & "Pick a base point: ")
	End With
	'' WCS origin
	dblOrigin(0) = 0: dblOrigin(1) = 0: dblOrigin(2) = 0
	'' move selection to the origin
	For Each objEnt In objSS
	objEnt.Move varBase, dblOrigin
	Next
	'' wblock selection to file name
	ThisDrawing.Wblock strFilename, objSS
	'' move selection back
	For Each objEnt In objSS
	objEnt.Move dblOrigin, varBase
	Next
	'' clean up selection set
	objSS.Delete
End Sub

Public Sub TestAddAttribute()
	Dim dblOrigin(2) As Double
	Dim dblEnt(2) As Double
	Dim dblHeight As Double
	Dim lngMode As Long
	Dim strTag As String
	Dim strPrompt As String
	Dim strValue As String
	Dim objBlock As AcadBlock
	Dim objEnt As AcadEntity
	'' create the block
	dblOrigin(0) = 0: dblOrigin(1) = 0: dblOrigin(2) = 0
	Set objBlock = ThisDrawing.Blocks.Add(dblOrigin, "Affirmations")
	'' delete existing entities (in case we've run before)
	For Each objEnt In objBlock
	objEnt.Delete
	Next
	'' create an ellipse in the block
	dblEnt(0) = 4: dblEnt(1) = 0: dblEnt(2) = 0
	objBlock.AddEllipse dblOrigin, dblEnt, 0.5
	'' set the height for all attributes
	dblHeight = 0.25
	dblEnt(0) = -1.5: dblEnt(1) = 0: dblEnt(2) = 0
	'' create a regular attribute
	lngMode = acAttributeModeNormal
	strTag = "Regular"
	strPrompt = "Enter a value"
	strValue = "I'm regular"
	dblEnt(1) = 1
	objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
	strValue
	'' create an invisible attribute
	lngMode = acAttributeModeInvisible
	strTag = "Invisible"
	strPrompt = "Enter a hidden value"
	strValue = "I'm invisible"
	dblEnt(1) = 0.5
	objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
	strValue
	'' create a constant attribute
	lngMode = acAttributeModeConstant
	strTag = "Constant"
	strPrompt = "Don't bother"
	strValue = "I'm set"
	dblEnt(1) = 0
	objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
	strValue
	'' create a verify attribute
	lngMode = acAttributeModeVerify
	strTag = "Verify"
	strPrompt = "Enter an important value"
	strValue = "I'm important"
	dblEnt(1) = -0.5
	objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
	strValue
	'' create a preset attribute
	lngMode = acAttributeModePreset
	strTag = "Preset"
	strPrompt = "No question"
	strValue = "I've got values"
	dblEnt(1) = -1
	objBlock.AddAttribute dblHeight, lngMode, strPrompt, dblEnt, strTag, _
	strValue
	'' now insert block interactively using sendcommand
	ThisDrawing.SendCommand "._-insert" & vbCr & "Affirmations" & vbCr
	-1-D:\_S\Excel workbook\CAD_Col\13.BlocksAttributes\mk13_AddAttributeMethod.bas Thursday, May 20, 2021 8:33 PM
End Sub

Public Sub AddBlock()
	Dim dblOrigin(2) As Double
	Dim objBlock As AcadBlock
	Dim strName As String
	'' get a name from user
	strName = InputBox("Enter a new block name: ")
	If "" = strName Then Exit Sub ' exit if no old name
	'' set the origin point
	dblOrigin(0) = 0: dblOrigin(1) = 0: dblOrigin(2) = 0
	''check if block already exists
	On Error Resume Next
	Set objBlock = ThisDrawing.Blocks.Item(strName)
	If Not objBlock Is Nothing Then
	MsgBox "Block already exists"
	Exit Sub
	End If
	'' create the block
	Set objBlock = ThisDrawing.Blocks.Add(dblOrigin, strName)
	'' then add entities (circle)
	objBlock.AddCircle dblOrigin, 10
End Sub

Public Sub TestAttachExternalReference()
	Dim strPath As String
	Dim strName As String
	Dim varInsertionPoint As Variant
	Dim dblX As Double
	Dim dblY As Double
	Dim dblZ As Double
	Dim dblRotation As Double
	Dim strInput As String
	Dim blnOver As Boolean
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	strPath = .GetString(True, vbCr & "External file name: ")
	.InitializeUserInput 1
	strName = .GetString(True, vbCr & "Block name to create: ")
	.InitializeUserInput 1
	varInsertionPoint = .GetPoint(, vbCr & "Pick the insert point: ")
	.InitializeUserInput 1 + 2
	dblX = .GetDistance(varInsertionPoint, vbCr & "X scale: ")
	.InitializeUserInput 1 + 2
	dblY = .GetDistance(varInsertionPoint, vbCr & "Y scale: ")
	.InitializeUserInput 1 + 2
	dblZ = .GetDistance(varInsertionPoint, vbCr & "Z scale: ")
	.InitializeUserInput 1
	dblRotation = .GetAngle(varInsertionPoint, vbCr & "Rotation angle: ")
	.InitializeUserInput 1, "Attach Overlay"
	strInput = .GetKeyword(vbCr & "Type [Attach/Overlay]: ")
	blnOver = IIf("Overlay" = strInput, True, False)
	End With
	'' create the object
	ThisDrawing.ModelSpace.AttachExternalReference strPath, strName, _
	varInsertionPoint, dblX, dblY, dblZ, dblRotation, blnOver
End Sub

Public Sub TestBind()
	Dim strName As String
	Dim strOpt As String
	Dim objBlock As AcadBlock
	On Error Resume Next
	'' get input from user
	With ThisDrawing.Utility
	'' get the block name
	.InitializeUserInput 1
	strName = .GetString(True, vbCr & "External reference name: ")
	If Err Then Exit Sub
	'' get the block definition
	Set objBlock = ThisDrawing.Blocks.Item(strName)
	'' exit if not found
	If Err Then
	MsgBox "Unable to get block " & strName
	Exit Sub
	End If
	'' exit if not an xref
	If Not objBlock.IsXRef Then
	MsgBox "That is not an external reference"
	Exit Sub
	End If
	'' get the option
	.InitializeUserInput 1, "Prefix Merge"
	strOpt = .GetKeyword(vbCr & "Dependent entries [Prefix/Merge]: ")
	If Err Then Exit Sub
	'' perform the bind, using option entered
	objBlock.Bind ("Merge" = strOpt)
	End With
End Sub

Public Sub TestCopyObjects()
	Dim objSS As AcadSelectionSet
	Dim varBase As Variant
	Dim objBlock As AcadBlock
	Dim strName As String
	Dim strErase As String
	Dim varEnt As Variant
	Dim objSourceEnts() As Object
	Dim varDestEnts As Variant
	Dim dblOrigin(2) As Double
	Dim intI As Integer
	'choose a selection set name that you use only as temporary storage and
	'ensure that it does not currently exist
	On Error Resume Next
	ThisDrawing.SelectionSets.Item("TempSSet").Delete
	Set objSS = ThisDrawing.SelectionSets.Add("TempSSet")
	objSS.SelectOnScreen
	'' get the other user input
	With ThisDrawing.Utility
	.InitializeUserInput 1
	strName = .GetString(True, vbCr & "Enter a block name: ")
	.InitializeUserInput 1
	varBase = .GetPoint(, vbCr & "Pick a base point: ")
	.InitializeUserInput 1, "Yes No"
	strErase = .GetKeyword(vbCr & "Erase originals [Yes/No]? ")
	End With
	'' set WCS origin
	dblOrigin(0) = 0: dblOrigin(1) = 0: dblOrigin(2) = 0
	'' create the block
	Set objBlock = ThisDrawing.Blocks.Add(dblOrigin, strName)
	'' put selected entities into an array for CopyObjects
	ReDim objSourceEnts(objSS.Count - 1)
	For intI = 0 To objSS.Count - 1
	Set objSourceEnts(intI) = objSS(intI)
	Next
	'' copy the entities into block
	varDestEnts = ThisDrawing.CopyObjects(objSourceEnts, objBlock)
	'' move copied entities so that base point becomes origin
	For Each varEnt In varDestEnts
	varEnt.Move varBase, dblOrigin
	Next
	'' if requested, erase the originals
	If strErase = "Yes" Then
	objSS.Erase
	End If
	'' we're done - prove that we did it
	ThisDrawing.SendCommand "._-insert" & vbCr & strName & vbCr
	'' clean up selection set
	objSS.Delete
End Sub

Public Sub DeleteBlock()
	Dim strName As String
	Dim objBlock As AcadBlock
	On Error Resume Next ' handle exceptions inline
	strName = InputBox("Block name to delete: ")
	If "" = strName Then Exit Sub ' exit if no old name
	Set objBlock = ThisDrawing.Blocks.Item(strName)
	If objBlock Is Nothing Then ' exit if not found
	MsgBox "Block '" & strName & "' not found"
	Exit Sub
	End If
	objBlock.Delete ' try to delete it
	If Err Then ' check if it worked
	MsgBox "Unable to delete Block: " & vbCr & Err.Description
	Else
	MsgBox "Block '" & strName & "' deleted"
	End If
End Sub

Function GetAttributes(objBlock As AcadBlock) As Collection
	Dim objEnt As AcadEntity
	Dim objAttribute As AcadAttribute
	Dim coll As New Collection
	'' iterate the block
	For Each objEnt In objBlock
	'' if it's an attribute
	If objEnt.ObjectName = "AcDbAttributeDefinition" Then
	'' cast to an attribute
	Set objAttribute = objEnt
	'' add attribute to the collection
	coll.Add objAttribute, objAttribute.TagString
	End If
	Next
	'' return collection
	Set GetAttributes = coll
End Function


Public Sub DemoGetAttributes()
	Dim objAttribs As Collection
	Dim objAttrib As AcadAttribute
	Dim objBlock As AcadBlock
	Dim strAttribs As String
	'' get the block
	Set objBlock = ThisDrawing.Blocks.Item("Affirmations")
	'' get the attributes
	Set objAttribs = GetAttributes(objBlock)
	'' show some information about each
	For Each objAttrib In objAttribs
	strAttribs = objAttrib.TagString & vbCrLf
	strAttribs = strAttribs & "Tag: " & objAttrib.TagString & vbCrLf & _
	"Prompt: " & objAttrib.PromptString & vbCrLf & " Value: " & _
	objAttrib.TextString & vbCrLf & " Mode: " & _
	objAttrib.Mode
	MsgBox strAttribs
	Next
	'' find specific attribute by TagString
	Set objAttrib = objAttribs.Item("PRESET")
	'' prove that we have the right one
	strAttribs = "Tag: " & objAttrib.TagString & vbCrLf & "Prompt: " & _
	objAttrib.PromptString & vbCrLf & "Value: " & objAttrib.TextString & _
	vbCrLf & "Mode: " & objAttrib.Mode
	MsgBox strAttribs
End Sub

Public Sub TestGetAttributes()
	Dim varPick As Variant
	Dim objEnt As AcadEntity
	Dim objBRef As AcadBlockReference
	Dim varAttribs As Variant
	Dim strAttribs As String
	Dim intI As Integer
	On Error Resume Next
	With ThisDrawing.Utility
	'' get an entity from user
	.GetEntity objEnt, varPick, vbCr & "Pick a block with attributes: "
	If Err Then Exit Sub
	'' cast it to a blockref
	Set objBRef = objEnt
	'' exit if not a block
	If objBRef Is Nothing Then
	.Prompt vbCr & "That wasn't a block."
	Exit Sub
	End If
	'' exit if it has no attributes
	If Not objBRef.HasAttributes Then
	.Prompt vbCr & "That block doesn't have attributes."
	Exit Sub
	End If
	'' get the attributerefs
	varAttribs = objBRef.GetAttributes
	'' show some information about each
	strAttribs = "Block Name: " & objBRef.Name & vbCrLf
	For intI = LBound(varAttribs) To UBound(varAttribs)
	strAttribs = strAttribs & " Tag(" & intI & "): " & _
	varAttribs(intI).TagString & vbTab & " Value(" & intI & "): " & _
	varAttribs(intI).TextString & vbCrLf
	Next
	End With
	MsgBox strAttribs
End Sub

Public Sub TestGetConstantAttributes()
	Dim varPick As Variant
	Dim objEnt As AcadEntity
	Dim objBRef As AcadBlockReference
	Dim varAttribs As Variant
	Dim strAttribs As String
	Dim intI As Integer
	On Error Resume Next
	With ThisDrawing.Utility
	'' get an entity from user
	.GetEntity objEnt, varPick, vbCr & _
	"Pick a block with constant attributes: "
	If Err Then Exit Sub
	'' cast it to a blockref
	Set objBRef = objEnt
	'' exit if not a block
	If objBRef Is Nothing Then
	.Prompt vbCr & "That wasn't a block."
	Exit Sub
	End If
	'' exit if it has no attributes
	If Not objBRef.HasAttributes Then
	.Prompt vbCr & "That block doesn't have attributes."
	Exit Sub
	End If
	'' get the constant attributes
	varAttribs = objBRef.GetConstantAttributes
	'' show some information about each
	strAttribs = "Block Name: " & objBRef.Name & vbCrLf
	For intI = LBound(varAttribs) To UBound(varAttribs)
	strAttribs = strAttribs & " Tag(" & intI & "): " & _
	varAttribs(intI).TagString & vbTab & "Value(" & intI & "): " & _
	varAttribs(intI).TextString
	Next
	End With
	MsgBox strAttribs
End Sub

Public Sub TestInsertAndSetAttributes()
	Dim objBRef As AcadBlockReference
	Dim varAttribRef As Variant
	Dim varInsertionPoint As Variant
	Dim dblX As Double
	Dim dblY As Double
	Dim dblZ As Double
	Dim dblRotation As Double
	'' get block input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	varInsertionPoint = .GetPoint(, vbCr & "Pick the insert point: ")
	.InitializeUserInput 1 + 2
	dblX = .GetDistance(varInsertionPoint, vbCr & "X scale: ")
	.InitializeUserInput 1 + 2
	dblY = .GetDistance(varInsertionPoint, vbCr & "Y scale: ")
	.InitializeUserInput 1 + 2
	dblZ = .GetDistance(varInsertionPoint, vbCr & "Z scale: ")
	.InitializeUserInput 1
	dblRotation = .GetAngle(varInsertionPoint, vbCr & "Rotation angle: ")
	End With
	'' insert the block
	Set objBRef = ThisDrawing.ModelSpace.InsertBlock(varInsertionPoint, _
	"Affirmations", dblX, dblY, dblZ, dblRotation)
	'' iterate the attributerefs
	For Each varAttribRef In objBRef.GetAttributes
	'' change specific values based on Tag
	Select Case varAttribRef.TagString
	Case "Regular":
	varAttribRef.TextString = "I have new values"
	Case "Invisible":
	varAttribRef.TextString = "I'm still invisible"
	Case "Verify":
	varAttribRef.TextString = "No verification needed"
	Case "Preset":
	varAttribRef.TextString = "I can be changed"
	End Select
	Next
End Sub

Function GetAttributes(objBlock As AcadBlock) As Collection
	Dim objEnt As AcadEntity
	Dim objAttribute As AcadAttribute
	Dim coll As New Collection
	'' iterate the block
	For Each objEnt In objBlock
	'' if it's an attribute
	If objEnt.ObjectName = "AcDbAttributeDefinition" Then
	'' cast to an attribute
	Set objAttribute = objEnt
	'' add attribute to the collection
	coll.Add objAttribute, objAttribute.TagString
	End If
	Next
	'' return collection
	Set GetAttributes = coll
End Function


Public Sub DemoGetAttributes()
	Dim objAttribs As Collection
	Dim objAttrib As AcadAttribute
	Dim objBlock As AcadBlock
	Dim strAttribs As String
	'' get the block
	Set objBlock = ThisDrawing.Blocks.Item("Affirmations")
	'' get the attributes
	Set objAttribs = GetAttributes(objBlock)
	'' show some information about each
	For Each objAttrib In objAttribs
	strAttribs = objAttrib.TagString & vbCrLf
	strAttribs = strAttribs & "Tag: " & objAttrib.TagString & vbCrLf & _
	"Prompt: " & objAttrib.PromptString & vbCrLf & " Value: " & _
	objAttrib.TextString & vbCrLf & " Mode: " & _
	objAttrib.Mode
	MsgBox strAttribs
	Next
	'' find specific attribute by TagString
	Set objAttrib = objAttribs.Item("PRESET")
	'' prove that we have the right one
	strAttribs = "Tag: " & objAttrib.TagString & vbCrLf & "Prompt: " & _
	objAttrib.PromptString & vbCrLf & "Value: " & objAttrib.TextString & _
	vbCrLf & "Mode: " & objAttrib.Mode
	MsgBox strAttribs
End Sub

Public Sub ListBlocks()
	Dim objBlock As AcadBlock
	Dim strBlockList As String
	strBlockList = "List of blocks: "
	For Each objBlock In ThisDrawing.Blocks
	strBlockList = strBlockList & vbCr & objBlock.Name
	Next
	MsgBox strBlockList
End Sub

Public Sub TestAddMInsertBlock()
	Dim strName As String
	Dim varInsertionPoint As Variant
	Dim dblX As Double
	Dim dblY As Double
	Dim dblZ As Double
	Dim dR As Double
	Dim lngNRows As Long
	Dim lngNCols As Long
	Dim dblSRows As Double
	Dim dblSCols As Double
	'' get input from user
	With ThisDrawing.Utility
	.InitializeUserInput 1
	strName = .GetString(True, vbCr & "Block or file name: ")
	.InitializeUserInput 1
	varInsertionPoint = .GetPoint(, vbCr & "Pick the insert point: ")
	.InitializeUserInput 1 + 2
	dblX = .GetDistance(varInsertionPoint, vbCr & "X scale: ")
	.InitializeUserInput 1 + 2
	dblY = .GetDistance(varInsertionPoint, vbCr & "Y scale: ")
	.InitializeUserInput 1 + 2
	dblZ = .GetDistance(varInsertionPoint, vbCr & "Z scale: ")
	.InitializeUserInput 1
	dR = .GetAngle(varInsertionPoint, vbCr & "Rotation angle: ")
	.InitializeUserInput 1 + 2 + 4
	lngNRows = .GetInteger(vbCr & "Number of rows: ")
	.InitializeUserInput 1 + 2 + 4
	lngNCols = .GetInteger(vbCr & "Number of columns: ")
	.InitializeUserInput 1 + 2
	dblSRows = .GetDistance(varInsertionPoint, vbCr & "Row spacing: ")
	.InitializeUserInput 1 + 2
	dblSCols = .GetDistance(varInsertionPoint, vbCr & "Column spacing: ")
	End With
	'' create the object
	ThisDrawing.ModelSpace.AddMInsertBlock varInsertionPoint, strName, _
	dblX, dblY, dblZ, dR, lngNRows, lngNCols, dblSRows, dblSCols
End Sub

Public Sub TestExternalReference()
	Dim strName As String
	Dim strOpt As String
	Dim objBlock As AcadBlock
	On Error Resume Next '' get input from user
	With ThisDrawing.Utility
	'' get the block name
	.InitializeUserInput 1
	strName = .GetString(True, vbCr & "External reference name: ")
	If Err Then Exit Sub
	'' get the block definition
	Set objBlock = ThisDrawing.Blocks.Item(strName)
	'' exit if not found
	If Err Then
	MsgBox "Unable to get block " & strName
	Exit Sub
	End If
	'' exit if not an xref
	If Not objBlock.IsXRef Then
	MsgBox "That is not an external reference"
	Exit Sub
	End If
	'' get the operation
	.InitializeUserInput 1, "Detach Reload Unload"
	strOpt = .GetKeyword(vbCr & "Option [Detach/Reload/Unload]: ")
	If Err Then Exit Sub
	'' perform operation requested
	If strOpt = "Detach" Then
	objBlock.Detach
	ElseIf strOpt = "Reload" Then
	objBlock.Reload
	Else
	objBlock.Unload
	End If
	End With
End Sub

Public Sub RenameBlock()
	Dim strName As String
	Dim objBlock As AcadBlock
	On Error Resume Next ' handle exceptions inline
	strName = InputBox("Original Block name: ")
	If "" = strName Then Exit Sub ' exit if no old name
	Set objBlock = ThisDrawing.Blocks.Item(strName)
	If objBlock Is Nothing Then ' exit if not found
	MsgBox "Block '" & strName & "' not found"
	Exit Sub
	End If
	strName = InputBox("New Block name: ")
	If "" = strName Then Exit Sub ' exit if no new name
	objBlock.Name = strName ' try and change name
	If Err Then ' check if it worked
	MsgBox "Unable to rename block: " & vbCr & Err.Description
	Else
	MsgBox "Block renamed to '" & strName & "'"
	End If
End Sub

Public Sub TestEditobjMInsertBlock()
	Dim objMInsert As AcadMInsertBlock
	Dim varPick As Variant
	Dim lngNRows As Long
	Dim lngNCols As Long
	Dim dblSRows As Double
	Dim dblSCols As Double
	On Error Resume Next
	'' get an entity and input from user
	With ThisDrawing.Utility
	.GetEntity objMInsert, varPick, vbCr & "Pick an MInsert: "
	If objMInsert Is Nothing Then
	MsgBox "You did not choose an MInsertBlock object"
	Exit Sub
	End If
	.InitializeUserInput 1 + 2 + 4
	lngNRows = .GetInteger(vbCr & "Number of rows: ")
	.InitializeUserInput 1 + 2 + 4
	lngNCols = .GetInteger(vbCr & "Number of columns: ")
	.InitializeUserInput 1 + 2
	dblSRows = .GetDistance(varPick, vbCr & "Row spacing: ")
	.InitializeUserInput 1 + 2
	dblSCols = .GetDistance(varPick, vbCr & "Column spacing: ")
	End With
	'' update the objMInsert
	With objMInsert
	.Rows = lngNRows
	.Columns = lngNCols
	.RowSpacing = dblSRows
	.ColumnSpacing = dblSCols
	.Update
	End With
End Sub


'14.View&ViewPorts


Public Sub MakeLineType()
	Dim LineText As String
	Dim LineSpace As Double
	Dim LinFile As String
	Dim Line1 As String
	Dim Line2 As String
	Dim myDir As String
	myDir = CurDir()
	If Right(myDir, 1) <> "" Then myDir = myDir & ""
	'change next line so LinFile = your lin file
	'if you do not what to use this filename
	LinFile = myDir & "TextInline.lin"
	LineText = InputBox("Enter Text for This LineType", "Make Text Line By John W. Anstaett")
	LineSpace = CDbl(InputBox("Enter Text Spaceing in Drawing Units", "Make Text Line By John W.
	Anstaett"))
	Line1 = "*" & LineText & ",----" & LineText & "----"
	Line2 = "A, " & CStr(LineSpace) & ",[""" & LineText
	Line2 = Line2 & """,Standard,y=-.5], "
	Line2 = Line2 & CStr(-Len(LineText))
	Open LinFile For Append As 1
	Print #1, Line1
	Print #1, Line2
	Close 1
	ThisDrawing.Linetypes.Load LineText, LinFile
	'you cant delete next line if you do not what the msgbox
	MsgBox "Line Type " & LineText & "Loaded ", , "Make Text Line By John W. Anstaett"
End Sub

Public Sub CreatePViewports()
	Dim objTopVPort As AcadPViewport
	Dim objFrontVPort As AcadPViewport
	Dim objRightVPort As AcadPViewport
	Dim objIsoMetricVPort As AcadPViewport
	Dim objLayout As AcadLayout
	Dim objAcadObject As AcadObject
	Dim dblPoint(2) As Double
	Dim dblViewDirection(2) As Double
	Dim dblOrigin(1) As Double
	Dim dblHeight As Double
	Dim dblWidth As Double
	Dim varMarginLL As Variant
	Dim varMarginUR As Variant
	ThisDrawing.ActiveSpace = acPaperSpace
	Set objLayout = ThisDrawing.ActiveLayout
	dblOrigin(0) = 0: dblOrigin(1) = 0
	objLayout.PlotOrigin = dblOrigin
	If objLayout.PlotRotation = ac0degrees Or objLayout.PlotRotation = _
	ac180degrees Then
	objLayout.GetPaperSize dblWidth, dblHeight
	Else
	objLayout.GetPaperSize dblHeight, dblWidth
	End If
	objLayout.GetPaperMargins varMarginLL, varMarginUR
	dblWidth = dblWidth - (varMarginUR(0) + varMarginLL(0))
	dblHeight = dblHeight - (varMarginUR(1) + varMarginLL(1))
	dblWidth = dblWidth / 2#
	dblHeight = dblHeight / 2#
	'Clear the layout of old PViewports
	For Each objAcadObject In ThisDrawing.PaperSpace
	If TypeName(objAcadObject) = "IAcadPViewport" Then
	objAcadObject.Delete
	End If
	Next
	'create Top Viewport
	dblPoint(0) = dblWidth - dblWidth * 0.5 '25
	dblPoint(1) = dblHeight - dblHeight * 0.5 '75
	dblPoint(2) = 0#
	Set objTopVPort = ThisDrawing.PaperSpace.AddPViewport(dblPoint, _
	dblWidth, dblHeight)
	'need to set view direction
	dblViewDirection(0) = 0
	dblViewDirection(1) = 0
	dblViewDirection(2) = 1
	objTopVPort.Direction = dblViewDirection
	objTopVPort.Display acOn
	ThisDrawing.MSpace = True
	ThisDrawing.ActivePViewport = objTopVPort
	ThisDrawing.Application.ZoomExtents
	ThisDrawing.Application.ZoomScaled 0.5, acZoomScaledRelativePSpace
	'create Front Viewport
	dblPoint(0) = dblWidth - dblWidth * 0.5
	dblPoint(1) = dblHeight + dblHeight * 0.5
	dblPoint(2) = 0
	Set objFrontVPort = ThisDrawing.PaperSpace.AddPViewport(dblPoint, _
	dblWidth, dblHeight)
	'need to set view direction
	dblViewDirection(0) = 0
	dblViewDirection(1) = -1
	dblViewDirection(2) = 0
	objFrontVPort.Direction = dblViewDirection
	objFrontVPort.Display acOn
	ThisDrawing.MSpace = True
	ThisDrawing.ActivePViewport = objFrontVPort
	-1-D:\_S\Excel workbook\CAD_Col\14.View&ViewPorts\mk14_ PaperSpaceViewport.bas Thursday, May 20, 2021 8:35 PM
	ThisDrawing.Application.ZoomExtents
	ThisDrawing.Application.ZoomScaled 0.5, acZoomScaledRelativePSpace
	'create Right Viewport
	dblPoint(0) = dblWidth + dblWidth * 0.5
	dblPoint(1) = dblHeight - dblHeight * 0.5
	dblPoint(2) = 0
	Set objRightVPort = ThisDrawing.PaperSpace.AddPViewport(dblPoint, _
	dblWidth, dblHeight)
	'need to set view direction
	dblViewDirection(0) = 1
	dblViewDirection(1) = 0
	dblViewDirection(2) = 0
	objRightVPort.Direction = dblViewDirection
	objRightVPort.Display acOn
	ThisDrawing.MSpace = True
	ThisDrawing.ActivePViewport = objRightVPort
	ThisDrawing.Application.ZoomExtents
	ThisDrawing.Application.ZoomScaled 0.5, acZoomScaledRelativePSpace
	'create Isometric Viewport
	dblPoint(0) = dblWidth + dblWidth * 0.5
	dblPoint(1) = dblHeight + dblHeight * 0.5
	dblPoint(2) = 0
	Set objIsoMetricVPort = ThisDrawing.PaperSpace.AddPViewport(dblPoint, _
	dblWidth, dblHeight)
	'need to set view direction
	dblViewDirection(0) = 1
	dblViewDirection(1) = -1
	dblViewDirection(2) = 1
	objIsoMetricVPort.Direction = dblViewDirection
	objIsoMetricVPort.Display acOn
	ThisDrawing.MSpace = True
	ThisDrawing.ActivePViewport = objIsoMetricVPort
	ThisDrawing.Application.ZoomExtents
	ThisDrawing.Application.ZoomScaled 0.5, acZoomScaledRelativePSpace
	'make paper space active again and we're almost done
	ThisDrawing.ActiveSpace = acPaperSpace
	ThisDrawing.Application.ZoomExtents
	'regen in all viewports
	ThisDrawing.Regen acAllViewports
End Sub

Public Sub AddView()
	Dim objView As AcadView
	Dim objActViewPort As AcadViewport
	Dim strNewViewName As String
	Dim varCenterPoint As Variant
	Dim dblPoint(1) As Double
	strNewViewName = InputBox("Enter name for new view: ")
	If strNewViewName = "" Then Exit Sub
	On Error Resume Next
	Set objView = ThisDrawing.Views.Item(strNewViewName)
	If objView Is Nothing Then
	Set objView = ThisDrawing.Views.Add(strNewViewName)
	ThisDrawing.ActiveSpace = acModelSpace
	varCenterPoint = ThisDrawing.GetVariable("VIEWCTR")
	dblPoint(0) = varCenterPoint(0): dblPoint(1) = varCenterPoint(1)
	Set objActViewPort = ThisDrawing.ActiveViewport
	'Current view info. is stored in the new AcadView object
	With objView
	.Center = dblPoint
	.Direction = ThisDrawing.GetVariable("VIEWDIR")
	.Height = objActViewPort.Height
	.Target = objActViewPort.Target
	.Width = objActViewPort.Width
	End With
	MsgBox "A new view called " & objView.Name & _
	" has been added to the Views collection."
	Else
	MsgBox "This view already exists."
	End If
End Sub

Public Sub DeleteView()
	Dim objView As AcadView
	Dim strViewName As String
	Dim strExistingViewNames As String
	For Each objView In ThisDrawing.Views
	strExistingViewNames = strExistingViewNames & objView.Name & vbCrLf
	Next
	strViewName = InputBox("Existing Views: " & vbCrLf & _
	strExistingViewNames & vbCrLf & _
	"Enter the view you wish to delete from the list.")
	If strViewName = "" Then Exit Sub
	On Error Resume Next
	Set objView = ThisDrawing.Views.Item(strViewName)
	If Not objView Is Nothing Then
	objView.Delete
	Else
	MsgBox "View was not recognized."
	End If
End Sub

Public Sub DisplayViews()
	Dim objView As AcadView
	Dim strViewNames As String
	If ThisDrawing.Views.Count > 0 Then
	For Each objView In ThisDrawing.Views
	strViewNames = strViewNames & objView.Name & vbCrLf
	Next
	MsgBox "The following views are saved for this drawing:" & vbCrLf _
	& strViewNames
	Else
	MsgBox "There are no saved View objects in the Views collection."
	End If
End Sub

Public Sub CreateViewport()
	Dim objViewPort As AcadViewport
	Dim objCurrentViewport As AcadViewport
	Dim varLowerLeft As Variant
	Dim dblViewDirection(2) As Double
	Dim strViewPortName As String
	strViewPortName = InputBox("Enter a name for the new viewport.")
	'user cancelled
	If strViewPortName = "" Then Exit Sub
	'check if viewport already exists
	On Error Resume Next
	Set objViewPort = ThisDrawing.Viewports.Item(strViewPortName)
	If Not objViewPort Is Nothing Then
	MsgBox "Viewport already exists"
	Exit Sub
	End If
	'Create a new viewport
	Set objViewPort = ThisDrawing.Viewports.Add(strViewPortName)
	'Split the screen viewport into 4 windows
	objViewPort.Split acViewport4
	For Each objCurrentViewport In ThisDrawing.Viewports
	If objCurrentViewport.LowerLeftCorner(0) = 0 Then
	If objCurrentViewport.LowerLeftCorner(1) = 0 Then
	'this takes care of Top view
	dblViewDirection(0) = 0
	dblViewDirection(1) = 0
	dblViewDirection(2) = 1
	objCurrentViewport.Direction = dblViewDirection
	Else
	'this takes care of Front view
	dblViewDirection(0) = 0
	dblViewDirection(1) = -1
	dblViewDirection(2) = 0
	objCurrentViewport.Direction = dblViewDirection
	End If
	End If
	If objCurrentViewport.LowerLeftCorner(0) = 0.5 Then
	If objCurrentViewport.LowerLeftCorner(1) = 0 Then
	'this takes care of the Right view
	dblViewDirection(0) = 1
	dblViewDirection(1) = 0
	dblViewDirection(2) = 0
	objCurrentViewport.Direction = dblViewDirection
	Else
	'this takes care of the Isometric view
	dblViewDirection(0) = 1
	dblViewDirection(1) = -1
	dblViewDirection(2) = 1
	objCurrentViewport.Direction = dblViewDirection
	End If
	End If
	Next
	'make viewport active to see effects of changes
	ThisDrawing.ActiveViewport = objViewPort
End Sub

Public Sub SetView()
	Dim objView As AcadView
	Dim objActViewPort As AcadViewport
	Dim strViewName As String
	ThisDrawing.ActiveSpace = acModelSpace
	Set objActViewPort = ThisDrawing.ActiveViewport
	'Redefine the current ViewPort with the View info
	strViewName = InputBox("Enter the view you require.")
	If strViewName = "" Then Exit Sub
	On Error Resume Next
	Set objView = ThisDrawing.Views.Item(strViewName)
	If Not objView Is Nothing Then
	objActViewPort.SetView objView
	ThisDrawing.ActiveViewport = objActViewPort
	Else
	MsgBox "View was not recognized."
	End If
End Sub

Public Sub PlotLayouts()
	Dim objLayout As AcadLayout
	Dim strLayoutList() As String
	Dim intCount As Integer
	Dim objPlot As AcadPlot
	intCount = -1
	For Each objLayout In ThisDrawing.Layouts
	If MsgBox("Do you wish to plot the layout: " _
	& objLayout.Name, vbYesNo) = vbYes Then
	intCount = intCount + 1
	ReDim Preserve strLayoutList(intCount)
	strLayoutList(intCount) = objLayout.Name
	End If
	Next objLayout
	Set objPlot = ThisDrawing.Plot
	objPlot.SetLayoutsToPlot strLayoutList
	objPlot.PlotToDevice
End Sub

'15.Text


Sub Ch4_UpdateTextFont()
	MsgBox ("Look at the text now...")
	Dim typeFace As String
	Dim SavetypeFace As String
	Dim Bold As Boolean
	Dim Italic As Boolean
	Dim charSet As Long
	Dim PitchandFamily As Long
	' Get the current settings to fill in the
	' default values for the SetFont method
	ThisDrawing.ActiveTextStyle.GetFont typeFace, Bold, Italic, charSet, PitchandFamily
	' ThisDrawing.ActiveTextStyle.height = 85
	' Change the typeface for the font
	SavetypeFace = typeFace
	typeFace = "PlayBill"
	ThisDrawing.ActiveTextStyle.SetFont typeFace, Bold, Italic, charSet, PitchandFamily
	ThisDrawing.Regen acActiveViewport
	MsgBox ("Now see how it looks after changing the font...")
	'Restore the original typeface
	ThisDrawing.ActiveTextStyle.SetFont SavetypeFace, Bold, Italic, charSet, PitchandFamily
	ThisDrawing.Regen acActiveViewport
End Sub


Sub Ch4_CreateMText()
	Dim mtextObj As AcadMText
	Dim insertPoint(0 To 2) As Double
	Dim width As Double
	Dim textString As String
	insertPoint(0) = 2
	insertPoint(1) = 2
	insertPoint(2) = 0
	width = 4
	textString = "This is a text string for the mtext object."
	' Create a text Object in model space
	Set mtextObj = ThisDrawing.ModelSpace.AddMText(insertPoint, width, textString)
	ZoomAll
End Sub

Sub Ch4_FormatMText()
	Dim mtextObj As AcadMText
	Dim insertPoint(0 To 2) As Double
	Dim width As Double
	Dim textString As String
	insertPoint(0) = 2
	insertPoint(1) = 2
	insertPoint(2) = 0
	width = 4
	' Define the ASCII characters for the control characters
	Dim OB As Long ' Open Bracket {
	Dim CB As Long ' Close Bracket }
	Dim BS As Long ' Back Slash \
	Dim FS As Long ' Forward Slash /
	Dim SC As Long ' Semicolon ;
	OB = Asc("{")
	CB = Asc("}")
	BS = Asc("\")
	FS = Asc("/")
	SC = Asc(";")
	' Assign the text string the following line of control
	' characters and text characters:
	' {{\H1.5x; Big text}\A2; over text\A1;/\A0; under text}
	textString = Chr(OB) + Chr(OB) + Chr(BS) + "H1.5x" _
	+ Chr(SC) + "Big text" + Chr(CB) + Chr(BS) + "A2" _
	+ Chr(SC) + "over text" + Chr(BS) + "A1" + Chr(SC) _
	+ Chr(FS) + Chr(BS) + "A0" + Chr(SC) + "under text" _
	+ Chr(CB)
	' Create a text Object in model space
	Set mtextObj = ThisDrawing.ModelSpace.AddMText(insertPoint, width, textString)
	ZoomAll
End Sub

Sub test512()
	Dim mymtext As AcadMText
	Dim MyMTextString As String
	Dim Mypoint As Variant
	Mypoint = ThisDrawing.Utility.GetPoint(, "specify location:")
	MyMTextString = "test {\C1;test} test"
	Set mymtext = ThisDrawing.ActiveLayout.Block.AddMText(Mypoint, 20, MyMTextString)
	mymtext.Height = 5 ' set referenced object height
End Sub

Sub TestString()
	Dim MyString As String
	Dim MyPoint As Variant
	MyPoint = ThisDrawing.Utility.GetPoint(, "Pick a point: ")
	MyString = MyPoint(1)
	MyString = Replace(MyString, ",", ".")
	Debug.Print MyString
End Sub
' formating
' MyStr = Format(7.999E-4, "##,##0.0000") ' Returns "0.0008".
' MyStr = Format(5459.4, "##,##0.00") ' Returns "5,459.40".
' MyStr = Format(334.9, "###0.00") ' Returns "334.90".
' MyStr = Format(5, "0.00%") ' Returns "500.00%".
' MyStr = Format("HELLO", "<") ' Returns "hello".


'Area_calcu

Sub testaddssObJ()
	Dim ss As AcadSelectionSet
	Dim ent As AcadPolyline
	Dim ents() As AcadPolyline
	Dim i As Integer
	' Dim plineObj As AcadPolyline
	'create selection set
	Set ss = ThisDrawing.SelectionSets.Add("ioyyyt")
	pto = ThisDrawing.Utility.GetPoint(, vbCr & "pick the first point ")
	xo = pto(0): yo = pto(1)
	Dim ppb(0 To 14) As Double
	ppb(0) = xo: ppb(1) = yo
	ppb(3) = xo + 30: ppb(4) = yo
	ppb(6) = xo + 30: ppb(7) = yo - 30
	ppb(9) = xo + 30 - 30: ppb(10) = yo - 30
	ppb(12) = xo + 30 - 30: ppb(13) = yo - 30 + 30
	Set ent = ThisDrawing.ModelSpace.AddPolyline(ppb)
	Stop
	' For Each ent In ThisDrawing.ModelSpace
	' ReDim Preserve ents(i)
	' Set ents(i) = ent
	' i = i + 1
	' Next
	i = 0
	ReDim Preserve ents(i)
	Set ents(i) = ent
	Stop
	ss.AddItems ents
	'' Do whatever with the selection set
	' ss.Select
	Stop
	' ss.Highlight True
	Stop
	yu = ents(0).Area
	''''''''''
	ss.Delete
End Sub

Sub testaddssObJ()
	Dim ss As AcadSelectionSet
	Dim ent As AcadEntity
	Dim ents() As AcadEntity
	Dim i As Integer
	' Dim plineObj As AcadPolyline
	'create selection set
	Set ss = ThisDrawing.SelectionSets.Add("ioioit")
	'
	pto = ThisDrawing.Utility.GetPoint(, vbCr & "pick the first point ")
	xo = pto(0): yo = pto(1)
	Dim ppb(0 To 14) As Double
	ppb(0) = xo: ppb(1) = yo
	ppb(3) = xo + 30: ppb(4) = yo
	ppb(6) = xo + 30: ppb(7) = yo - 30
	ppb(9) = xo + 30 - 30: ppb(10) = yo - 30
	ppb(12) = xo + 30 - 30: ppb(13) = yo - 30 + 30
	Set ent = ThisDrawing.ModelSpace.AddPolyline(ppb)
	Stop
	For Each ent In ThisDrawing.ModelSpace
	ReDim Preserve ents(i)
	Set ents(i) = ent
	i = i + 1
	Next
	Stop
	ss.AddItems ents
	'' Do whatever with the selection set
	' ss.Select
	' ss.Delete
	Stop
	ss.Highlight True
	yu = ents.Area
End Sub


Sub testaddssObJ()
	Dim ss As AcadSelectionSet
	Dim ent As AcadPolyline
	Dim ents() As AcadPolyline
	Dim i As Integer
	' Dim plineObj As AcadPolyline
	'create selection set
	Set ss = ThisDrawing.SelectionSets.Add("ioyyyt")
	pto = ThisDrawing.Utility.GetPoint(, vbCr & "pick the first point ")
	xo = pto(0): yo = pto(1)
	Dim ppb(0 To 14) As Double
	ppb(0) = xo: ppb(1) = yo
	ppb(3) = xo + 30: ppb(4) = yo
	ppb(6) = xo + 30: ppb(7) = yo - 30
	ppb(9) = xo + 30 - 30: ppb(10) = yo - 30
	ppb(12) = xo + 30 - 30: ppb(13) = yo - 30 + 30
	Set ent = ThisDrawing.ModelSpace.AddPolyline(ppb)
	Stop
	' For Each ent In ThisDrawing.ModelSpace
	' ReDim Preserve ents(i)
	' Set ents(i) = ent
	' i = i + 1
	' Next
	i = 0
	ReDim Preserve ents(i)
	Set ents(i) = ent
	Stop
	ss.AddItems ents
	'' Do whatever with the selection set
	' ss.Select
	Stop
	' ss.Highlight True
	Stop
	yu = ents(0).Area
	''''''''''
	ss.Delete
End Sub


Sub experiment()
	Dim blockobj As AcadBlockReference
	Dim LL As Variant
	Dim UR As Variant
	Dim area as double
	For Each blockobj In ThisDrawing.ModelSpace
	blockobj.GetBoundingBox LL, UR
	area = (UR(0)-LL(0))*(UR(1)-LL(1))
	msgbox "'Area' of current block is: " & area
End Sub
'script for calculating area

Sub cccc()
	Dim x As AcadEntity
	Dim pt As Variant
	ThisDrawing.Utility.GetEntity x, pt
	Area = x.Area
End Sub
''''''''''''''''''''''''''''''''''''''''''''''_________'''''
'by selection sets....


Sub SelectionLWPolyline()
	Dim filterType(0) As Integer
	Dim filterData(0) As Variant
	For Each MySelection In ThisDrawing.SelectionSets
	If MySelection.Name = "PP1" Then
	MySelection.Delete
	Exit For
	End If
	Next
	Set MySelection = ThisDrawing.SelectionSets.Add("PP1")
	filterType(0) = 0
	filterData(0) = "LWPOLYLINE"
	MySelection.Select acSelectionSetAll, , , filterType, filterData
	If MySelection.Count >= 1 Then
	MyArea = MySelection.Item(0).Area
	End If
End Sub

'command_activate

Sub Line_to_Polyline_Filtered_SelSet()
	Dim cmd As String
	Dim curEnt As AcadEntity
	Dim ssType(0) As Integer: Dim ssData(0)
	ssType(0) = 0: ssData(0) = "LINE"
	Dim ssTest As AcadSelectionSet
	Set ssTest = ThisDrawing.PickfirstSelectionSet
	ssTest.Clear
	ssTest.Select acSelectionSetAll, , , ssType, ssData
	For Each curEnt In ssTest 'ThisDrawing.ModelSpace
	'If TypeOf curEnt Is AcadLine Then
	ThisDrawing.SendCommand "pedit " & "(handent """ & _
	curEnt.Handle & """) "
	'End If
	Next
End Sub


'coordinates


'Coordinate Example
Sub Example_Coordinate()
	' This example creates a polyline in model space and
	' queries and changes the coordinate in the first index position.
	Dim plineObj As AcadPolyline
	Dim points(0 To 14) As Double
	' Define the 2D polyline points
	points(0) = 1: points(1) = 1: points(2) = 0
	points(3) = 1: points(4) = 2: points(5) = 0
	points(6) = 2: points(7) = 2: points(8) = 0
	points(9) = 3: points(10) = 2: points(11) = 0
	points(12) = 4: points(13) = 4: points(14) = 0
	' Create a lightweight Polyline object in model space
	Set plineObj = ThisDrawing.ModelSpace.AddPolyline(points)
	ZoomAll
	' Find the coordinate in the first index position
	Dim coord As Variant
	coord = plineObj.Coordinate(0)
	MsgBox "The coordinate in the first index position of the polyline is: " & coord(0) & ", " _
	& coord(1) & ", " & coord(2)
	' Change the coordinate
	coord(0) = coord(0) + 1
	plineObj.Coordinate(0) = coord
	plineObj.Update
	' Query the new coordinate
	coord = plineObj.Coordinate(0)
	MsgBox "The coordinate in the first index position of the polyline is now: " & coord(0) &
	", " _
	& coord(1) & ", " & coord(2)
End Sub

'Coordinate

sub cooordinatesCAD()
	'some how declare the worksheet working on..........
	dh=thisworkbook.worksheet."sheet1"
	'some how play with ...active AUTOCAD file
	'active.autocad......soem...s.s.s.s.
	'activate the drawing... here..
	'select drawing here....
	Set NewDC = GetObject(, "AutoCAD.Application")
	Set A2Kdwg = NewDC.ActiveDocument
	Dim Selection As AcadSelectionSet
	Dim poly As AcadLWPolyline
	Dim Obj As AcadEntity
	Dim Bound As Double
	Dim x, y As Double
	Dim rows, i, scount As Integer
	'---Search Object from SelectionSet and Delete If Found ----''
	For i = 0 To A2Kdwg.SelectionSets.count - 1
	If A2Kdwg.SelectionSets.Item(i).Name = "AcDbPolyline" Then
	''-- Delete Object Name from AutoCAD SelectionSet ---''
	A2Kdwg.SelectionSets.Item(i).Delete
	Exit For
	End If
	Next i
	''-- Add Object to AutoCad SelectionSet ----''
	Set Selection = A2Kdwg.SelectionSets.Add("AcDbPolyline")
	''-- Select Object from AutoCad Screen ---'''
	Selection.SelectOnScreen
	''-- Get Coordinates of Object if Object name is ACadPolyline--''
	rows = cx
	For Each Obj In Selection
	If Obj.ObjectName = "AcDbPolyline" Then
	''- Set Obj as Polyline--''
	Set poly = Obj
	On Error Resume Next
	''-- Set Size of Coordinates Like array Size--''
	Bound = UBound(poly.Coordinates)
	'' Starting Index of Excel Row to insert Coordinates --''
	rows = rows
	''-- Display Coordinates one by one to Excel Columns --'''
	For i = 0 To Bound
	''-- Set Coordinates into Variables--'''
	x = Round(poly.Coordinates(i), 3)
	y = Round(poly.Coordinates(i + 1), 3)
	''-- Set Coordinates into Excel Columns --
	'' some how manipulaete here...
	'' with excel worksheet selected... activate...it.. then...locate place to copy cooordinates...
	to...
	Worksheets(sname).Cells(rows, cy) = Round(x, 3)
	Worksheets(sname).Cells(rows, cy + 1) = Round(y, 4)
	''- Increment variable for Excel Rows ---''
	rows = rows + 1
	''--- Increment Counter variable to get Next point of Polyline --'''
	i = i + 1
	Next
	Else
	MsgBox "--- This is not a Polyline --- ", vbInformation, "Please Select a
	Polyline"
	-1-D:\_S\Excel workbook\CAD_Col\coordinates\coordinates.txt Thursday, May 20, 2021 8:51 PM
	End If
	rows = rows + 1
	Next Obj
end sub

'File


Attribute VB_Name = "Module1"
Option Explicit
'' based on listing written by Ken Puls (www.exelguru.ca)

Sub TextStreamTest()
	Const ForReading = 1, ForWriting = 2, ForAppending = 3
	Dim fso As Object
	Dim fdw As Object, sdw As Object
	Dim fdr As Object, sdr As Object
	Dim strText As String
	'create file sysytem object
	Set fso = CreateObject("Scripting.FileSystemObject")
	'get file to read and open text stream
	Set fdr = fso.GetFile("C:\Wat\Lookdb.ldb")
	Set sdr = fdr.OpenAsTextStream(ForReading, False)
	'create new file to write to
	fso.CreateTextFile ("C:\Wat\Lookdb.txt")
	Set fdw = fso.GetFile("C:\Wat\Lookdb.txt")
	Set sdw = fdw.OpenAsTextStream(ForAppending, False)
	'iterate through the end of first file
	Do Until sdr.AtEndOfStream
	'read string from the first file
	strText = sdr.ReadLine & vbNewLine '<--added carriage return to jump on the next line
	'write to the second one
	sdw.Write strText
	Loop
	'close both files and clean up
	sdr.Close
	sdw.Close
	Set sdr = Nothing
	Set sdw = Nothing
	Set fdr = Nothing
	Set fdw = Nothing
	Set fso = Nothing
End Sub

Private Sub SubmitButton_Click()
	'deletes unselected items to leave only the drawing needed
	Dim i As Long, pth, file, xlFile, MyArray() As Variant, Cnt As Long
	Dim excel As Object, xlWkSht As Object
	file = ThisDrawing.name
	pth = ThisDrawing.path & "\"
	file = Left(ThisDrawing.name, (Len(file) - 4)) & ".xls"
	'On Error GoTo 0
	On Error Resume Next
	Set excel = GetObject(, "Excel.Application")
	If Err <> 0 Then
	Err.Clear
	Set excel = CreateObject("Excel.Application")
	If Err <> 0 Then
	MsgBox "Could Not Load Excel!", vbExclamation
	End If
	End If
	excel.Visible = True
	Set xlFile = excel.Workbooks.Open(pth & file)
	'*resume error handling
	'On Error GoTo 0
	'*DELETE XL SHEETS
	With Me.ListBox1
	Cnt = 0
	For i = 0 To .ListCount - 1
	If .Selected(i) = False Then
	Cnt = Cnt + 1
	ReDim Preserve MyArray(1 To Cnt)
	MyArray(Cnt) = .List(i)
	End If
	Next i
	If Cnt > 0 Then
	If excel.Workbooks(file).Worksheets.count > UBound(MyArray) Then
	excel.DisplayAlerts = False
	'On Error Resume Next
	On Error GoTo 0
	excel.Workbooks(file).Worksheets(MyArray).Delete
	excel.DisplayAlerts = True
	'Call UpdateSheetList
	Else
	MsgBox "A workbook must contain at least one visible sheet.", vbExclamation
	End If
	Else
	MsgBox "Please select one or more sheets for deletion...", vbExclamation
	End If
	End With
	'*DELETE AUTOCAD LAYOUTS
	With Me.ListBox1
	For i = 0 To .ListCount - 1
	If .Selected(i) = False Then
	ThisDrawing.Layouts.Item(ListBox1.List(i, 0)).Delete
	End If
	Next i
	End With
	Call Format
	Unload Me
End Sub

Option Explicit
'' based on listing written by Ken Puls (www.exelguru.ca)
Sub TextStreamTest()
	Const ForReading = 1, ForWriting = 2, ForAppending = 3
	Dim fso As Object
	Dim fdw As Object, sdw As Object
	Dim fdr As Object, sdr As Object
	Dim strText As String
	'create file sysytem object
	Set fso = CreateObject("Scripting.FileSystemObject")
	'get file to read and open text stream
	Set fdr = fso.GetFile("C:\Wat\Lookdb.ldb")
	Set sdr = fdr.OpenAsTextStream(ForReading, False)
	'create new file to write to
	fso.CreateTextFile ("C:\Wat\Lookdb.txt")
	Set fdw = fso.GetFile("C:\Wat\Lookdb.txt")
	Set sdw = fdw.OpenAsTextStream(ForAppending, False)
	'iterate through the end of first file
	Do Until sdr.AtEndOfStream
	'read string from the first file
	strText = sdr.ReadLine & vbNewLine '<--added carriage return to jump on the next line
	'write to the second one
	sdw.Write strText
	Loop
	'close both files and clean up
	sdr.Close
	sdw.Close
	Set sdr = Nothing
	Set sdw = Nothing
	Set fdr = Nothing
	Set fdw = Nothing
	Set fso = Nothing
End Sub

Private Declare Function GetOpenFileName _
	Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" _
	(pOpenfilename As OPENFILENAME) As Long
	Private Type OPENFILENAME
	lStructSize As Long
	hwndOwner As Long
	hInstance As Long
	lpstrFilter As String
	lpstrCustomFilter As String
	nMaxCustFilter As Long
	nFilterIndex As Long
	lpstrFile As String
	nMaxFile As Long
	lpstrFileTitle As String
	nMaxFileTitle As Long
	lpstrInitialDir As String
	lpstrTitle As String
	Flags As Long
	nFileOffset As Integer
	nFileExtension As Integer
	lpstrDefExt As String
	lCustData As Long
	lpfnHook As Long
	lpTemplateName As String
	End Type
	Const OFN_ALLOWMULTISELECT As Long = &H200
	Const OFN_EXPLORER As Long = &H80000
	Const OFN_NOCHANGEDIR As Long = &H8
	Sub GetFileFromAPI()
	Dim OFN As OPENFILENAME
	With OFN
	.lStructSize = Len(OFN) ' Size of structure.
	.lpstrInitialDir = "C:\PlotItIn\"
	.nMaxFile = 260 ' Size of buffer.
	' Create buffer.
	.lpstrFile = String(.nMaxFile - 1, 0)
	.Flags = OFN.Flags Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER
	.nMaxFile = 10000
	Ret = GetOpenFileName(OFN) ' Call function.
	If Ret <> 0 Then ' Non-zero is success.
	' Find first null char.
	n = InStr(.lpstrFile, vbNullChar)
	' Return what's before it.
	MsgBox Left(.lpstrFile, n - 1)
	End If
	End With
End Sub


Public Sub Server_Backup()
	'Option Explicit
	On Error Resume Next
	Dim ServerPath As String
	Dim OriginalPath As String
	Dim FilePath As String
	Dim PathLength As Integer
	Dim Search1 As Integer
	Dim Search2 As Integer
	Dim SearchLen As Integer
	Dim OldSearch As String
	Dim Set1 As String
	Dim Set2 As String
	Dim CompletePath As String
	Dim BkFile As String
	Dim BkLen As Integer
	OriginalPath = ThisDrawing.FullName
	' Check to see if file is saved
	If ThisDrawing.FullName = "" Then
	MsgBox "You must save this drawing before it can be backed up to the server", vbCritical,
	"Backup Error"
	Exit Sub
	End If
	' Check to see if the file exists on the server, and if not, create it
	ServerPath = "G" + Mid(ThisDrawing.FullName, 2, Len(ThisDrawing.FullName))
	If Dir(ServerPath) = "" Then
	For i = 3 To Len(ThisDrawing.FullName)
	If i > Len(ThisDrawing.FullName) Then Exit For
	Search1 = InStr(i, ThisDrawing.FullName, "\")
	Search2 = InStr(Search1 + 1, ThisDrawing.FullName, "\")
	SearchLen = Search2 - Search1
	FilePath = Mid(ThisDrawing.FullName, Search1, SearchLen)
	If FilePath = OldSearch Then GoTo Loop1
	If Search2 = Search1 Then Search2 = i + 1
	OldSearch = FilePath
	Set1 = Set1 + FilePath
	Set2 = "G:" + Set1
	MkDir Set2
	Loop1:
	Next i
	End If
	' save drawing to the server
	ThisDrawing.SaveAs (ServerPath)
	BkLen = Len(ServerPath) - 4
	BkFile = Mid(ServerPath, 1, BkLen) + ".BAK"
	Kill (BkFile)
	' save drawing back to local machine
	ThisDrawing.SaveAs (OriginalPath)
End Sub

Public Sub WriteXRec()
	Dim oDict As AcadDictionary
	Dim oXRec As AcadXRecord
	Dim dxfCode(0 To 1) As Integer
	Dim dxfData(0 To 1)
	Set oDict = ThisDrawing.Dictionaries.Add("SampleTest")
	Set oXRec = oDict.AddXRecord("Record1")
	dxfCode(0) = 1: dxfData(0) = "First Value"
	dxfCode(1) = 2: dxfData(1) = "Second Value"
	oXRec.SetXRecordData dxfCode, dxfData
End Sub

Public Sub ReadXRec()
	Dim oDict As AcadDictionary
	Dim oXRec As AcadXRecord
	Dim dxfCode, dxfData
	Set oDict = ThisDrawing.Dictionaries.Item("SampleTest")
	Set oXRec = oDict.Item("Record1")
	oXRec.GetXRecordData dxfCode, dxfData
	Debug.Print dxfData(0)
	Debug.Print dxfData(1)
End Sub
'colorconversions.dvb sample code for AutoCAD

Sub main()
	Dim RGB
	RGB = lookUpRGB(112)
	Debug.Print RGB(0)
	Debug.Print RGB(1)
	Debug.Print RGB(2)
End Sub

Private Function lookUpRGB(ByVal ACI As Integer) As Integer()
	Dim ACItoRGB(0 To 255, 0 To 2) As Integer
	ACItoRGB(0, 0) = 0: ACItoRGB(0, 1) = 0: ACItoRGB(0, 2) = 0
	ACItoRGB(1, 0) = 255: ACItoRGB(1, 1) = 0: ACItoRGB(1, 2) = 0
	ACItoRGB(2, 0) = 255: ACItoRGB(2, 1) = 255: ACItoRGB(2, 2) = 0
	ACItoRGB(3, 0) = 0: ACItoRGB(3, 1) = 255: ACItoRGB(3, 2) = 0
	ACItoRGB(4, 0) = 0: ACItoRGB(4, 1) = 255: ACItoRGB(4, 2) = 255
	ACItoRGB(5, 0) = 0: ACItoRGB(5, 1) = 0: ACItoRGB(5, 2) = 255
	ACItoRGB(6, 0) = 255: ACItoRGB(6, 1) = 0: ACItoRGB(6, 2) = 255
	ACItoRGB(7, 0) = 255: ACItoRGB(7, 1) = 255: ACItoRGB(7, 2) = 255
	ACItoRGB(8, 0) = 128: ACItoRGB(8, 1) = 128: ACItoRGB(8, 2) = 128
	ACItoRGB(9, 0) = 192: ACItoRGB(9, 1) = 192: ACItoRGB(9, 2) = 192
	ACItoRGB(10, 0) = 255: ACItoRGB(10, 1) = 1: ACItoRGB(10, 2) = 1
	ACItoRGB(11, 0) = 255: ACItoRGB(11, 1) = 127: ACItoRGB(11, 2) = 127
	ACItoRGB(12, 0) = 165: ACItoRGB(12, 1) = 0: ACItoRGB(12, 2) = 0
	ACItoRGB(13, 0) = 165: ACItoRGB(13, 1) = 82: ACItoRGB(13, 2) = 82
	ACItoRGB(14, 0) = 127: ACItoRGB(14, 1) = 0: ACItoRGB(14, 2) = 0
	ACItoRGB(15, 0) = 127: ACItoRGB(15, 1) = 63: ACItoRGB(15, 2) = 63
	ACItoRGB(16, 0) = 76: ACItoRGB(16, 1) = 0: ACItoRGB(16, 2) = 0
	ACItoRGB(17, 0) = 76: ACItoRGB(17, 1) = 38: ACItoRGB(17, 2) = 38
	ACItoRGB(18, 0) = 38: ACItoRGB(18, 1) = 0: ACItoRGB(18, 2) = 0
	ACItoRGB(19, 0) = 38: ACItoRGB(19, 1) = 19: ACItoRGB(19, 2) = 19
	ACItoRGB(20, 0) = 255: ACItoRGB(20, 1) = 63: ACItoRGB(20, 2) = 0
	ACItoRGB(21, 0) = 255: ACItoRGB(21, 1) = 159: ACItoRGB(21, 2) = 127
	ACItoRGB(22, 0) = 165: ACItoRGB(22, 1) = 41: ACItoRGB(22, 2) = 0
	ACItoRGB(23, 0) = 165: ACItoRGB(23, 1) = 103: ACItoRGB(23, 2) = 82
	ACItoRGB(24, 0) = 127: ACItoRGB(24, 1) = 31: ACItoRGB(24, 2) = 0
	ACItoRGB(25, 0) = 127: ACItoRGB(25, 1) = 79: ACItoRGB(25, 2) = 63
	ACItoRGB(26, 0) = 76: ACItoRGB(26, 1) = 19: ACItoRGB(26, 2) = 0
	ACItoRGB(27, 0) = 76: ACItoRGB(27, 1) = 47: ACItoRGB(27, 2) = 38
	ACItoRGB(28, 0) = 38: ACItoRGB(28, 1) = 9: ACItoRGB(28, 2) = 0
	ACItoRGB(29, 0) = 38: ACItoRGB(29, 1) = 23: ACItoRGB(29, 2) = 19
	ACItoRGB(30, 0) = 255: ACItoRGB(30, 1) = 127: ACItoRGB(30, 2) = 0
	-1-D:\_S\Excel workbook\CAD_Col\File\CAD vba files.txt Thursday, May 20, 2021 8:52 PM
	ACItoRGB(31, 0) = 255: ACItoRGB(31, 1) = 191: ACItoRGB(31, 2) = 127
	ACItoRGB(32, 0) = 165: ACItoRGB(32, 1) = 82: ACItoRGB(32, 2) = 0
	ACItoRGB(33, 0) = 165: ACItoRGB(33, 1) = 124: ACItoRGB(33, 2) = 82
	ACItoRGB(34, 0) = 127: ACItoRGB(34, 1) = 63: ACItoRGB(34, 2) = 0
	ACItoRGB(35, 0) = 127: ACItoRGB(35, 1) = 95: ACItoRGB(35, 2) = 63
	ACItoRGB(36, 0) = 76: ACItoRGB(36, 1) = 38: ACItoRGB(36, 2) = 0
	ACItoRGB(37, 0) = 76: ACItoRGB(37, 1) = 57: ACItoRGB(37, 2) = 38
	ACItoRGB(38, 0) = 38: ACItoRGB(38, 1) = 19: ACItoRGB(38, 2) = 0
	ACItoRGB(39, 0) = 38: ACItoRGB(39, 1) = 28: ACItoRGB(39, 2) = 19
	ACItoRGB(40, 0) = 255: ACItoRGB(40, 1) = 191: ACItoRGB(40, 2) = 0
	ACItoRGB(41, 0) = 255: ACItoRGB(41, 1) = 223: ACItoRGB(41, 2) = 127
	ACItoRGB(42, 0) = 165: ACItoRGB(42, 1) = 124: ACItoRGB(42, 2) = 0
	ACItoRGB(43, 0) = 165: ACItoRGB(43, 1) = 145: ACItoRGB(43, 2) = 82
	ACItoRGB(44, 0) = 127: ACItoRGB(44, 1) = 95: ACItoRGB(44, 2) = 0
	ACItoRGB(45, 0) = 127: ACItoRGB(45, 1) = 111: ACItoRGB(45, 2) = 63
	ACItoRGB(46, 0) = 76: ACItoRGB(46, 1) = 57: ACItoRGB(46, 2) = 0
	ACItoRGB(47, 0) = 76: ACItoRGB(47, 1) = 66: ACItoRGB(47, 2) = 38
	ACItoRGB(48, 0) = 38: ACItoRGB(48, 1) = 28: ACItoRGB(48, 2) = 0
	ACItoRGB(49, 0) = 38: ACItoRGB(49, 1) = 33: ACItoRGB(49, 2) = 19
	ACItoRGB(50, 0) = 255: ACItoRGB(50, 1) = 255: ACItoRGB(50, 2) = 1
	ACItoRGB(51, 0) = 255: ACItoRGB(51, 1) = 255: ACItoRGB(51, 2) = 127
	ACItoRGB(52, 0) = 165: ACItoRGB(52, 1) = 165: ACItoRGB(52, 2) = 0
	ACItoRGB(53, 0) = 165: ACItoRGB(53, 1) = 165: ACItoRGB(53, 2) = 82
	ACItoRGB(54, 0) = 127: ACItoRGB(54, 1) = 127: ACItoRGB(54, 2) = 0
	ACItoRGB(55, 0) = 127: ACItoRGB(55, 1) = 127: ACItoRGB(55, 2) = 63
	ACItoRGB(56, 0) = 76: ACItoRGB(56, 1) = 76: ACItoRGB(56, 2) = 0
	ACItoRGB(57, 0) = 76: ACItoRGB(57, 1) = 76: ACItoRGB(57, 2) = 38
	ACItoRGB(58, 0) = 38: ACItoRGB(58, 1) = 38: ACItoRGB(58, 2) = 0
	ACItoRGB(59, 0) = 38: ACItoRGB(59, 1) = 38: ACItoRGB(59, 2) = 19
	ACItoRGB(60, 0) = 191: ACItoRGB(60, 1) = 255: ACItoRGB(60, 2) = 0
	ACItoRGB(61, 0) = 223: ACItoRGB(61, 1) = 255: ACItoRGB(61, 2) = 127
	ACItoRGB(62, 0) = 124: ACItoRGB(62, 1) = 165: ACItoRGB(62, 2) = 0
	ACItoRGB(63, 0) = 145: ACItoRGB(63, 1) = 165: ACItoRGB(63, 2) = 82
	ACItoRGB(64, 0) = 95: ACItoRGB(64, 1) = 127: ACItoRGB(64, 2) = 0
	ACItoRGB(65, 0) = 111: ACItoRGB(65, 1) = 127: ACItoRGB(65, 2) = 63
	ACItoRGB(66, 0) = 57: ACItoRGB(66, 1) = 76: ACItoRGB(66, 2) = 0
	ACItoRGB(67, 0) = 66: ACItoRGB(67, 1) = 76: ACItoRGB(67, 2) = 38
	ACItoRGB(68, 0) = 28: ACItoRGB(68, 1) = 38: ACItoRGB(68, 2) = 0
	ACItoRGB(69, 0) = 33: ACItoRGB(69, 1) = 38: ACItoRGB(69, 2) = 19
	ACItoRGB(70, 0) = 127: ACItoRGB(70, 1) = 255: ACItoRGB(70, 2) = 0
	ACItoRGB(71, 0) = 191: ACItoRGB(71, 1) = 255: ACItoRGB(71, 2) = 127
	ACItoRGB(72, 0) = 82: ACItoRGB(72, 1) = 165: ACItoRGB(72, 2) = 0
	ACItoRGB(73, 0) = 124: ACItoRGB(73, 1) = 165: ACItoRGB(73, 2) = 82
	ACItoRGB(74, 0) = 63: ACItoRGB(74, 1) = 127: ACItoRGB(74, 2) = 0
	ACItoRGB(75, 0) = 95: ACItoRGB(75, 1) = 127: ACItoRGB(75, 2) = 63
	ACItoRGB(76, 0) = 38: ACItoRGB(76, 1) = 76: ACItoRGB(76, 2) = 0
	ACItoRGB(77, 0) = 57: ACItoRGB(77, 1) = 76: ACItoRGB(77, 2) = 38
	ACItoRGB(78, 0) = 19: ACItoRGB(78, 1) = 38: ACItoRGB(78, 2) = 0
	ACItoRGB(79, 0) = 28: ACItoRGB(79, 1) = 38: ACItoRGB(79, 2) = 19
	ACItoRGB(80, 0) = 63: ACItoRGB(80, 1) = 255: ACItoRGB(80, 2) = 0
	ACItoRGB(81, 0) = 159: ACItoRGB(81, 1) = 255: ACItoRGB(81, 2) = 127
	ACItoRGB(82, 0) = 41: ACItoRGB(82, 1) = 165: ACItoRGB(82, 2) = 0
	ACItoRGB(83, 0) = 103: ACItoRGB(83, 1) = 165: ACItoRGB(83, 2) = 82
	ACItoRGB(84, 0) = 31: ACItoRGB(84, 1) = 127: ACItoRGB(84, 2) = 0
	ACItoRGB(85, 0) = 79: ACItoRGB(85, 1) = 127: ACItoRGB(85, 2) = 63
	ACItoRGB(86, 0) = 19: ACItoRGB(86, 1) = 76: ACItoRGB(86, 2) = 0
	ACItoRGB(87, 0) = 47: ACItoRGB(87, 1) = 76: ACItoRGB(87, 2) = 38
	ACItoRGB(88, 0) = 9: ACItoRGB(88, 1) = 38: ACItoRGB(88, 2) = 0
	ACItoRGB(89, 0) = 23: ACItoRGB(89, 1) = 38: ACItoRGB(89, 2) = 19
	ACItoRGB(90, 0) = 1: ACItoRGB(90, 1) = 255: ACItoRGB(90, 2) = 1
	ACItoRGB(91, 0) = 127: ACItoRGB(91, 1) = 255: ACItoRGB(91, 2) = 127
	ACItoRGB(92, 0) = 0: ACItoRGB(92, 1) = 165: ACItoRGB(92, 2) = 0
	ACItoRGB(93, 0) = 82: ACItoRGB(93, 1) = 165: ACItoRGB(93, 2) = 82
	ACItoRGB(94, 0) = 0: ACItoRGB(94, 1) = 127: ACItoRGB(94, 2) = 0
	ACItoRGB(95, 0) = 63: ACItoRGB(95, 1) = 127: ACItoRGB(95, 2) = 63
	ACItoRGB(96, 0) = 0: ACItoRGB(96, 1) = 76: ACItoRGB(96, 2) = 0
	-2-D:\_S\Excel workbook\CAD_Col\File\CAD vba files.txt Thursday, May 20, 2021 8:52 PM
	ACItoRGB(97, 0) = 38: ACItoRGB(97, 1) = 76: ACItoRGB(97, 2) = 38
	ACItoRGB(98, 0) = 0: ACItoRGB(98, 1) = 38: ACItoRGB(98, 2) = 0
	ACItoRGB(99, 0) = 19: ACItoRGB(99, 1) = 38: ACItoRGB(99, 2) = 19
	ACItoRGB(100, 0) = 0: ACItoRGB(100, 1) = 255: ACItoRGB(100, 2) = 63
	ACItoRGB(101, 0) = 127: ACItoRGB(101, 1) = 255: ACItoRGB(101, 2) = 159
	ACItoRGB(102, 0) = 0: ACItoRGB(102, 1) = 165: ACItoRGB(102, 2) = 41
	ACItoRGB(103, 0) = 82: ACItoRGB(103, 1) = 165: ACItoRGB(103, 2) = 103
	ACItoRGB(104, 0) = 0: ACItoRGB(104, 1) = 127: ACItoRGB(104, 2) = 31
	ACItoRGB(105, 0) = 63: ACItoRGB(105, 1) = 127: ACItoRGB(105, 2) = 79
	ACItoRGB(106, 0) = 0: ACItoRGB(106, 1) = 76: ACItoRGB(106, 2) = 19
	ACItoRGB(107, 0) = 38: ACItoRGB(107, 1) = 76: ACItoRGB(107, 2) = 47
	ACItoRGB(108, 0) = 0: ACItoRGB(108, 1) = 38: ACItoRGB(108, 2) = 9
	ACItoRGB(109, 0) = 19: ACItoRGB(109, 1) = 38: ACItoRGB(109, 2) = 23
	ACItoRGB(110, 0) = 0: ACItoRGB(110, 1) = 255: ACItoRGB(110, 2) = 127
	ACItoRGB(111, 0) = 127: ACItoRGB(111, 1) = 255: ACItoRGB(111, 2) = 191
	ACItoRGB(112, 0) = 0: ACItoRGB(112, 1) = 165: ACItoRGB(112, 2) = 82
	ACItoRGB(113, 0) = 82: ACItoRGB(113, 1) = 165: ACItoRGB(113, 2) = 124
	ACItoRGB(114, 0) = 0: ACItoRGB(114, 1) = 127: ACItoRGB(114, 2) = 63
	ACItoRGB(115, 0) = 63: ACItoRGB(115, 1) = 127: ACItoRGB(115, 2) = 95
	ACItoRGB(116, 0) = 0: ACItoRGB(116, 1) = 76: ACItoRGB(116, 2) = 38
	ACItoRGB(117, 0) = 38: ACItoRGB(117, 1) = 76: ACItoRGB(117, 2) = 57
	ACItoRGB(118, 0) = 0: ACItoRGB(118, 1) = 38: ACItoRGB(118, 2) = 19
	ACItoRGB(119, 0) = 19: ACItoRGB(119, 1) = 38: ACItoRGB(119, 2) = 28
	ACItoRGB(120, 0) = 0: ACItoRGB(120, 1) = 255: ACItoRGB(120, 2) = 191
	ACItoRGB(121, 0) = 127: ACItoRGB(121, 1) = 255: ACItoRGB(121, 2) = 223
	ACItoRGB(122, 0) = 0: ACItoRGB(122, 1) = 165: ACItoRGB(122, 2) = 124
	ACItoRGB(123, 0) = 82: ACItoRGB(123, 1) = 165: ACItoRGB(123, 2) = 145
	ACItoRGB(124, 0) = 0: ACItoRGB(124, 1) = 127: ACItoRGB(124, 2) = 95
	ACItoRGB(125, 0) = 63: ACItoRGB(125, 1) = 127: ACItoRGB(125, 2) = 111
	ACItoRGB(126, 0) = 0: ACItoRGB(126, 1) = 76: ACItoRGB(126, 2) = 57
	ACItoRGB(127, 0) = 38: ACItoRGB(127, 1) = 76: ACItoRGB(127, 2) = 66
	ACItoRGB(128, 0) = 0: ACItoRGB(128, 1) = 38: ACItoRGB(128, 2) = 28
	ACItoRGB(129, 0) = 19: ACItoRGB(129, 1) = 38: ACItoRGB(129, 2) = 33
	ACItoRGB(130, 0) = 1: ACItoRGB(130, 1) = 255: ACItoRGB(130, 2) = 255
	ACItoRGB(131, 0) = 127: ACItoRGB(131, 1) = 255: ACItoRGB(131, 2) = 255
	ACItoRGB(132, 0) = 0: ACItoRGB(132, 1) = 165: ACItoRGB(132, 2) = 165
	ACItoRGB(133, 0) = 82: ACItoRGB(133, 1) = 165: ACItoRGB(133, 2) = 165
	ACItoRGB(134, 0) = 0: ACItoRGB(134, 1) = 127: ACItoRGB(134, 2) = 127
	ACItoRGB(135, 0) = 63: ACItoRGB(135, 1) = 127: ACItoRGB(135, 2) = 127
	ACItoRGB(136, 0) = 0: ACItoRGB(136, 1) = 76: ACItoRGB(136, 2) = 76
	ACItoRGB(137, 0) = 38: ACItoRGB(137, 1) = 76: ACItoRGB(137, 2) = 76
	ACItoRGB(138, 0) = 0: ACItoRGB(138, 1) = 38: ACItoRGB(138, 2) = 38
	ACItoRGB(139, 0) = 19: ACItoRGB(139, 1) = 38: ACItoRGB(139, 2) = 38
	ACItoRGB(140, 0) = 0: ACItoRGB(140, 1) = 191: ACItoRGB(140, 2) = 255
	ACItoRGB(141, 0) = 127: ACItoRGB(141, 1) = 223: ACItoRGB(141, 2) = 255
	ACItoRGB(142, 0) = 0: ACItoRGB(142, 1) = 124: ACItoRGB(142, 2) = 165
	ACItoRGB(143, 0) = 82: ACItoRGB(143, 1) = 145: ACItoRGB(143, 2) = 165
	ACItoRGB(144, 0) = 0: ACItoRGB(144, 1) = 95: ACItoRGB(144, 2) = 127
	ACItoRGB(145, 0) = 63: ACItoRGB(145, 1) = 111: ACItoRGB(145, 2) = 127
	ACItoRGB(146, 0) = 0: ACItoRGB(146, 1) = 57: ACItoRGB(146, 2) = 76
	ACItoRGB(147, 0) = 38: ACItoRGB(147, 1) = 66: ACItoRGB(147, 2) = 76
	ACItoRGB(148, 0) = 0: ACItoRGB(148, 1) = 28: ACItoRGB(148, 2) = 38
	ACItoRGB(149, 0) = 19: ACItoRGB(149, 1) = 33: ACItoRGB(149, 2) = 38
	ACItoRGB(150, 0) = 0: ACItoRGB(150, 1) = 127: ACItoRGB(150, 2) = 255
	ACItoRGB(151, 0) = 127: ACItoRGB(151, 1) = 191: ACItoRGB(151, 2) = 255
	ACItoRGB(152, 0) = 0: ACItoRGB(152, 1) = 82: ACItoRGB(152, 2) = 165
	ACItoRGB(153, 0) = 82: ACItoRGB(153, 1) = 124: ACItoRGB(153, 2) = 165
	ACItoRGB(154, 0) = 0: ACItoRGB(154, 1) = 63: ACItoRGB(154, 2) = 127
	ACItoRGB(155, 0) = 63: ACItoRGB(155, 1) = 95: ACItoRGB(155, 2) = 127
	ACItoRGB(156, 0) = 0: ACItoRGB(156, 1) = 38: ACItoRGB(156, 2) = 76
	ACItoRGB(157, 0) = 38: ACItoRGB(157, 1) = 57: ACItoRGB(157, 2) = 76
	ACItoRGB(158, 0) = 0: ACItoRGB(158, 1) = 19: ACItoRGB(158, 2) = 38
	ACItoRGB(159, 0) = 19: ACItoRGB(159, 1) = 28: ACItoRGB(159, 2) = 38
	ACItoRGB(160, 0) = 0: ACItoRGB(160, 1) = 63: ACItoRGB(160, 2) = 255
	ACItoRGB(161, 0) = 127: ACItoRGB(161, 1) = 159: ACItoRGB(161, 2) = 255
	ACItoRGB(162, 0) = 0: ACItoRGB(162, 1) = 41: ACItoRGB(162, 2) = 165
	-3-D:\_S\Excel workbook\CAD_Col\File\CAD vba files.txt Thursday, May 20, 2021 8:52 PM
	ACItoRGB(163, 0) = 82: ACItoRGB(163, 1) = 103: ACItoRGB(163, 2) = 165
	ACItoRGB(164, 0) = 0: ACItoRGB(164, 1) = 31: ACItoRGB(164, 2) = 127
	ACItoRGB(165, 0) = 63: ACItoRGB(165, 1) = 79: ACItoRGB(165, 2) = 127
	ACItoRGB(166, 0) = 0: ACItoRGB(166, 1) = 19: ACItoRGB(166, 2) = 76
	ACItoRGB(167, 0) = 38: ACItoRGB(167, 1) = 47: ACItoRGB(167, 2) = 76
	ACItoRGB(168, 0) = 0: ACItoRGB(168, 1) = 9: ACItoRGB(168, 2) = 38
	ACItoRGB(169, 0) = 19: ACItoRGB(169, 1) = 23: ACItoRGB(169, 2) = 38
	ACItoRGB(170, 0) = 1: ACItoRGB(170, 1) = 1: ACItoRGB(170, 2) = 255
	ACItoRGB(171, 0) = 127: ACItoRGB(171, 1) = 127: ACItoRGB(171, 2) = 255
	ACItoRGB(172, 0) = 0: ACItoRGB(172, 1) = 0: ACItoRGB(172, 2) = 165
	ACItoRGB(173, 0) = 82: ACItoRGB(173, 1) = 82: ACItoRGB(173, 2) = 165
	ACItoRGB(174, 0) = 0: ACItoRGB(174, 1) = 0: ACItoRGB(174, 2) = 127
	ACItoRGB(175, 0) = 63: ACItoRGB(175, 1) = 63: ACItoRGB(175, 2) = 127
	ACItoRGB(176, 0) = 0: ACItoRGB(176, 1) = 0: ACItoRGB(176, 2) = 76
	ACItoRGB(177, 0) = 38: ACItoRGB(177, 1) = 38: ACItoRGB(177, 2) = 76
	ACItoRGB(178, 0) = 0: ACItoRGB(178, 1) = 0: ACItoRGB(178, 2) = 38
	ACItoRGB(179, 0) = 19: ACItoRGB(179, 1) = 19: ACItoRGB(179, 2) = 38
	ACItoRGB(180, 0) = 63: ACItoRGB(180, 1) = 0: ACItoRGB(180, 2) = 255
	ACItoRGB(181, 0) = 159: ACItoRGB(181, 1) = 127: ACItoRGB(181, 2) = 255
	ACItoRGB(182, 0) = 41: ACItoRGB(182, 1) = 0: ACItoRGB(182, 2) = 165
	ACItoRGB(183, 0) = 103: ACItoRGB(183, 1) = 82: ACItoRGB(183, 2) = 165
	ACItoRGB(184, 0) = 31: ACItoRGB(184, 1) = 0: ACItoRGB(184, 2) = 127
	ACItoRGB(185, 0) = 79: ACItoRGB(185, 1) = 63: ACItoRGB(185, 2) = 127
	ACItoRGB(186, 0) = 19: ACItoRGB(186, 1) = 0: ACItoRGB(186, 2) = 76
	ACItoRGB(187, 0) = 47: ACItoRGB(187, 1) = 38: ACItoRGB(187, 2) = 76
	ACItoRGB(188, 0) = 9: ACItoRGB(188, 1) = 0: ACItoRGB(188, 2) = 38
	ACItoRGB(189, 0) = 23: ACItoRGB(189, 1) = 19: ACItoRGB(189, 2) = 38
	ACItoRGB(190, 0) = 127: ACItoRGB(190, 1) = 0: ACItoRGB(190, 2) = 255
	ACItoRGB(191, 0) = 191: ACItoRGB(191, 1) = 127: ACItoRGB(191, 2) = 255
	ACItoRGB(192, 0) = 82: ACItoRGB(192, 1) = 0: ACItoRGB(192, 2) = 165
	ACItoRGB(193, 0) = 124: ACItoRGB(193, 1) = 82: ACItoRGB(193, 2) = 165
	ACItoRGB(194, 0) = 63: ACItoRGB(194, 1) = 0: ACItoRGB(194, 2) = 127
	ACItoRGB(195, 0) = 95: ACItoRGB(195, 1) = 63: ACItoRGB(195, 2) = 127
	ACItoRGB(196, 0) = 38: ACItoRGB(196, 1) = 0: ACItoRGB(196, 2) = 76
	ACItoRGB(197, 0) = 57: ACItoRGB(197, 1) = 38: ACItoRGB(197, 2) = 76
	ACItoRGB(198, 0) = 19: ACItoRGB(198, 1) = 0: ACItoRGB(198, 2) = 38
	ACItoRGB(199, 0) = 28: ACItoRGB(199, 1) = 19: ACItoRGB(199, 2) = 38
	ACItoRGB(200, 0) = 191: ACItoRGB(200, 1) = 0: ACItoRGB(200, 2) = 255
	ACItoRGB(201, 0) = 223: ACItoRGB(201, 1) = 127: ACItoRGB(201, 2) = 255
	ACItoRGB(202, 0) = 124: ACItoRGB(202, 1) = 0: ACItoRGB(202, 2) = 165
	ACItoRGB(203, 0) = 145: ACItoRGB(203, 1) = 82: ACItoRGB(203, 2) = 165
	ACItoRGB(204, 0) = 95: ACItoRGB(204, 1) = 0: ACItoRGB(204, 2) = 127
	ACItoRGB(205, 0) = 111: ACItoRGB(205, 1) = 63: ACItoRGB(205, 2) = 127
	ACItoRGB(206, 0) = 57: ACItoRGB(206, 1) = 0: ACItoRGB(206, 2) = 76
	ACItoRGB(207, 0) = 66: ACItoRGB(207, 1) = 38: ACItoRGB(207, 2) = 76
	ACItoRGB(208, 0) = 28: ACItoRGB(208, 1) = 0: ACItoRGB(208, 2) = 38
	ACItoRGB(209, 0) = 33: ACItoRGB(209, 1) = 19: ACItoRGB(209, 2) = 38
	ACItoRGB(210, 0) = 255: ACItoRGB(210, 1) = 1: ACItoRGB(210, 2) = 255
	ACItoRGB(211, 0) = 255: ACItoRGB(211, 1) = 127: ACItoRGB(211, 2) = 255
	ACItoRGB(212, 0) = 165: ACItoRGB(212, 1) = 0: ACItoRGB(212, 2) = 165
	ACItoRGB(213, 0) = 165: ACItoRGB(213, 1) = 82: ACItoRGB(213, 2) = 165
	ACItoRGB(214, 0) = 127: ACItoRGB(214, 1) = 0: ACItoRGB(214, 2) = 127
	ACItoRGB(215, 0) = 127: ACItoRGB(215, 1) = 63: ACItoRGB(215, 2) = 127
	ACItoRGB(216, 0) = 76: ACItoRGB(216, 1) = 0: ACItoRGB(216, 2) = 76
	ACItoRGB(217, 0) = 76: ACItoRGB(217, 1) = 38: ACItoRGB(217, 2) = 76
	ACItoRGB(218, 0) = 38: ACItoRGB(218, 1) = 0: ACItoRGB(218, 2) = 38
	ACItoRGB(219, 0) = 38: ACItoRGB(219, 1) = 19: ACItoRGB(219, 2) = 38
	ACItoRGB(220, 0) = 255: ACItoRGB(220, 1) = 0: ACItoRGB(220, 2) = 191
	ACItoRGB(221, 0) = 255: ACItoRGB(221, 1) = 127: ACItoRGB(221, 2) = 223
	ACItoRGB(222, 0) = 165: ACItoRGB(222, 1) = 0: ACItoRGB(222, 2) = 124
	ACItoRGB(223, 0) = 165: ACItoRGB(223, 1) = 82: ACItoRGB(223, 2) = 145
	ACItoRGB(224, 0) = 127: ACItoRGB(224, 1) = 0: ACItoRGB(224, 2) = 95
	ACItoRGB(225, 0) = 127: ACItoRGB(225, 1) = 63: ACItoRGB(225, 2) = 111
	ACItoRGB(226, 0) = 76: ACItoRGB(226, 1) = 0: ACItoRGB(226, 2) = 57
	ACItoRGB(227, 0) = 76: ACItoRGB(227, 1) = 38: ACItoRGB(227, 2) = 66
	ACItoRGB(228, 0) = 38: ACItoRGB(228, 1) = 0: ACItoRGB(228, 2) = 28
	-4-D:\_S\Excel workbook\CAD_Col\File\CAD vba files.txt Thursday, May 20, 2021 8:52 PM
	ACItoRGB(229, 0) = 38: ACItoRGB(229, 1) = 19: ACItoRGB(229, 2) = 33
	ACItoRGB(230, 0) = 255: ACItoRGB(230, 1) = 0: ACItoRGB(230, 2) = 127
	ACItoRGB(231, 0) = 255: ACItoRGB(231, 1) = 127: ACItoRGB(231, 2) = 191
	ACItoRGB(232, 0) = 165: ACItoRGB(232, 1) = 0: ACItoRGB(232, 2) = 82
	ACItoRGB(233, 0) = 165: ACItoRGB(233, 1) = 82: ACItoRGB(233, 2) = 124
	ACItoRGB(234, 0) = 127: ACItoRGB(234, 1) = 0: ACItoRGB(234, 2) = 63
	ACItoRGB(235, 0) = 127: ACItoRGB(235, 1) = 63: ACItoRGB(235, 2) = 95
	ACItoRGB(236, 0) = 76: ACItoRGB(236, 1) = 0: ACItoRGB(236, 2) = 38
	ACItoRGB(237, 0) = 76: ACItoRGB(237, 1) = 38: ACItoRGB(237, 2) = 57
	ACItoRGB(238, 0) = 38: ACItoRGB(238, 1) = 0: ACItoRGB(238, 2) = 19
	ACItoRGB(239, 0) = 38: ACItoRGB(239, 1) = 19: ACItoRGB(239, 2) = 28
	ACItoRGB(240, 0) = 255: ACItoRGB(240, 1) = 0: ACItoRGB(240, 2) = 63
	ACItoRGB(241, 0) = 255: ACItoRGB(241, 1) = 127: ACItoRGB(241, 2) = 159
	ACItoRGB(242, 0) = 165: ACItoRGB(242, 1) = 0: ACItoRGB(242, 2) = 41
	ACItoRGB(243, 0) = 165: ACItoRGB(243, 1) = 82: ACItoRGB(243, 2) = 103
	ACItoRGB(244, 0) = 127: ACItoRGB(244, 1) = 0: ACItoRGB(244, 2) = 31
	ACItoRGB(245, 0) = 127: ACItoRGB(245, 1) = 63: ACItoRGB(245, 2) = 79
	ACItoRGB(246, 0) = 76: ACItoRGB(246, 1) = 0: ACItoRGB(246, 2) = 19
	ACItoRGB(247, 0) = 76: ACItoRGB(247, 1) = 38: ACItoRGB(247, 2) = 47
	ACItoRGB(248, 0) = 38: ACItoRGB(248, 1) = 0: ACItoRGB(248, 2) = 9
	ACItoRGB(249, 0) = 38: ACItoRGB(249, 1) = 19: ACItoRGB(249, 2) = 23
	ACItoRGB(250, 0) = 84: ACItoRGB(250, 1) = 84: ACItoRGB(250, 2) = 84
	ACItoRGB(251, 0) = 118: ACItoRGB(251, 1) = 118: ACItoRGB(251, 2) = 118
	ACItoRGB(252, 0) = 152: ACItoRGB(252, 1) = 152: ACItoRGB(252, 2) = 152
	ACItoRGB(253, 0) = 186: ACItoRGB(253, 1) = 186: ACItoRGB(253, 2) = 186
	ACItoRGB(254, 0) = 220: ACItoRGB(254, 1) = 220: ACItoRGB(254, 2) = 220
	ACItoRGB(255, 0) = 252: ACItoRGB(255, 1) = 252: ACItoRGB(255, 2) = 252
	Dim arr(2) As Integer
	arr(0) = ACItoRGB(ACI, 0)
	arr(1) = ACItoRGB(ACI, 1)
	arr(2) = ACItoRGB(ACI, 2)
	lookUpRGB = arr
End Function
'line.dvb sample code for AutoCAD
Option Explicit
' Code by Jimmy B 2000-06-08
' No shortcut menu when in line command but in all others
' When no command is active right click is treated as an enter
Dim line As Boolean

Private Sub AcadDocument_Activate()
	line = False
End Sub

Private Sub AcadDocument_BeginCommand(ByVal CommandName As String)
	line = (CommandName = "LINE")
End Sub

Private Sub AcadDocument_BeginRightClick(ByVal PickPoint As Variant)
	If ThisDrawing.GetVariable("cmdactive") > 0 And line Then
	ThisDrawing.Application.Preferences.User.ShortCutMenuDisplay = False
	Else
	ThisDrawing.SetVariable "ShortcutMenu", 10
	End If
End Sub

Option Explicit
Dim objDbx As AxDbDocument
' 2000-03-08
' By Jimmy Bergmark
' Copyright (C) 1997-2003 JTB World, All Rights Reserved
' Website: www.jtbworld.com
' E-mail: info@jtbworld.com
' Runs in AutoCAD 2000 with axdb15.dll (must be referenced)
' Example of batch for listing all layers on all drawings in a directory.


Private Sub ListLayers()
	-5-D:\_S\Excel workbook\CAD_Col\File\CAD vba files.txt Thursday, May 20, 2021 8:52 PM
	Set objDbx = GetInterfaceObject("ObjectDBX.AxDbDocument")
	Dim inDir As String
	Dim elem As Object
	Dim filenom As String
	Dim WholeFile As String
	Dim newHeight As Double
	inDir = "r:\projekt\3828\A"
	filenom = Dir$(inDir & "\*.dwg")
	Do While filenom <> ""
	ThisDrawing.Utility.Prompt vbCrLf & "File: " & filenom
	ThisDrawing.Utility.Prompt vbCrLf & "-----------------"
	WholeFile = inDir & "\" & filenom
	objDbx.Open WholeFile
	For Each elem In objDbx.Layers
	ThisDrawing.Utility.Prompt vbCrLf & elem.Name
	Next
	Set elem = Nothing
	objDbx.SaveAs WholeFile
	filenom = Dir$
	ThisDrawing.Utility.Prompt vbCrLf
	Loop
End Sub
jDbx As AxDbDocument


Private Sub ListXREF()
	Set objDbx = GetInterfaceObject("ObjectDBX.AxDbDocument")
	Dim inDir As String
	Dim elem As Object
	Dim filenom As String
	Dim WholeFile As String
	Dim newHeight As Double
	inDir = "r:\projekt\3828\A"
	filenom = Dir$(inDir & "\*.dwg")
	Do While filenom <> ""
	ThisDrawing.Utility.Prompt vbCrLf & "File: " & filenom
	ThisDrawing.Utility.Prompt vbCrLf & "-----------------"
	WholeFile = inDir & "\" & filenom
	objDbx.Open WholeFile
	For Each elem In objDbx.Blocks
	If elem.IsXRef = True Then
	ThisDrawing.Utility.Prompt vbCrLf & elem.Name
	End If
	Next
	Set elem = Nothing
	objDbx.SaveAs WholeFile
	filenom = Dir$
	ThisDrawing.Utility.Prompt vbCrLf
	Loop
End Sub

Public Sub revert()
	If Documents.Count > 0 Then
	If ThisDrawing.GetVariable("DWGTITLED") = 1 Then
	If ThisDrawing.GetVariable("dbmod") > 0 Then
	Dim strName As String
	strName = ThisDrawing.FullName
	If MsgBox("Abandon changes to " & strName, vbOKCancel + vbExclamation + _
	vbDefaultButton2, "AutoCAD - REVERT VBA") <> vbCancel Then
	If AcadApplication.Preferences.System.SingleDocumentMode = False Then
	ActiveDocument.Close False
	Documents.Open strName
	Else
	ThisDrawing.Open strName
	End If
	End If
	End If
	Else
	MsgBox "Drawing has never been saved.", vbCritical + vbOKOnly, "AutoCAD - REVERT VBA"
	End If
	End If
	By Jimmy Bergmark
	' Copyright (C) 1997-2003 JTB World, All Rights Reserved
	' Website: www.jtbworld.com
	' E-mail: info@jtbworld.com
	Public Function DrawingVersion(strFullPath As String) As String
	On Error Resume Next
	Dim i As Long
	Dim bytVersion(0 To 5) As Byte
	Dim strVersion As String
	Dim lngFile As Long
	If Len(Dir(strFullPath)) > 0 Then
	lngFile = FreeFile
	Open strFullPath For Binary Access Read As lngFile
	Get #lngFile, , bytVersion
	Close lngFile
	strVersion = StrConv(bytVersion(), vbUnicode)
	End If
	If Len(strVersion) > 0 Then
	DrawingVersion = strVersion
	Else
	DrawingVersion = "NEWNEW"
	End If
	End Function
	Private Sub AcadDocument_EndSave(ByVal FileName As String)
	On Error Resume Next
	Dim dv As String
	If ThisDrawing.Application.Preferences.OpenSave.SaveAsType <> ac2004_dwg Then
	dv = DrawingVersion(FileName)
	If (dv = "AC402b" Or dv = "AC1018") Then
	ThisDrawing.Utility.Prompt vbNewLine & FileName & " Saved in 2004 format" & vbNewLine
	End If
	End If
End Sub


Private Sub AcadDocument_BeginCommand(ByVal CommandName As String)
	strCMD = CommandName
End Sub

Private Sub AcadDocument_EndCommand(ByVal CommandName As String)
	If CommandName = "WBLOCK" Or CommandName = "-WBLOCK" Then
	If ThisDrawing.Application.Preferences.OpenSave.SaveAsType <> ac2004_dwg Then
	If MsgBox("Do you want to open this wblock that is saved in 2004 DWG format" &
	vbNewLine & _
	"to be able to QSAVE it to 2000 DWG format?" & vbNewLine & vbNewLine & _
	"Just QSAVE and CLOSE it if you want to.", vbYesNo) = vbYes Then
	ThisDrawing.Application.Documents.Open (strFileName)
	ThisDrawing.Application.Documents.Item(ThisDrawing.Application.Documents.Count
	- 1).Activate
	End If
	End If
	End If
End Sub

Private Sub AcadDocument_EndSave(ByVal FileName As String)
	strFileName = FileName
End Sub

sub xref_test()
	Dim ps_str As String
	Dim po_blk As AcadBlock
	Dim po_blkref As AcadBlockReference
	Dim pi_dxf70 As Integer
	Dim pv_1 As Variant
	Dim userr1 As Integer
	ThisDrawing.Utility.GetEntity po_blkref, pv_1
	userr1 = ThisDrawing.GetVariable("USERR1")
	Set po_blk = ThisDrawing.Blocks(po_blkref.Name)
	ps_str = "(SETVAR ""USERR1"" (cdr (assoc 70 (tblsearch ""BLOCK"" """ & po_blkref.Name &
	""")))) "
	ThisDrawing.SendCommand ps_str
	pi_dxf70 = ThisDrawing.GetVariable("USERR1")
	ps_str = "(SETVAR ""USERR1"" " & userr1 & ") "
	ThisDrawing.SendCommand ps_str
	If pi_dxf70 = 44 Then MsgBox "XREF " & po_blk.Name & " is Overlaid."
	If pi_dxf70 = 36 Then MsgBox "XREF " & po_blk.Name & " is Attached."
End Sub

Option Explicit
'' based on listing written by Ken Puls (www.exelguru.ca)
Sub TextStreamTest()
	Const ForReading = 1, ForWriting = 2, ForAppending = 3
	Dim fso As Object
	Dim fdw As Object, sdw As Object
	Dim fdr As Object, sdr As Object
	Dim strText As String
	'create file sysytem object
	Set fso = CreateObject("Scripting.FileSystemObject")
	'get file to read and open text stream
	Set fdr = fso.GetFile("C:\Wat\Lookdb.ldb")
	Set sdr = fdr.OpenAsTextStream(ForReading, False)
	'create new file to write to
	fso.CreateTextFile ("C:\Wat\Lookdb.txt")
	Set fdw = fso.GetFile("C:\Wat\Lookdb.txt")
	Set sdw = fdw.OpenAsTextStream(ForAppending, False)
	'iterate through the end of first file
	Do Until sdr.AtEndOfStream
	'read string from the first file
	strText = sdr.ReadLine & vbNewLine '<--added carriage return to jump on the next line
	'write to the second one
	sdw.Write strText
	Loop
	'close both files and clean up
	sdr.Close
	sdw.Close
	Set sdr = Nothing
	Set sdw = Nothing
	Set fdr = Nothing
	Set fdw = Nothing
	Set fso = Nothing
End Sub

Option Explicit
Sub ExportPoints()
	'Declare variables
	Dim currentSelectionSet As AcadSelectionSet
	Dim ent As AcadEntity
	Dim pnt As AcadPoint
	Dim csvFile As String
	Dim FSO As FileSystemObject
	Dim textFile As TextStream
	'Create a reference to the selection set of the currently selected objects
	Set currentSelectionSet = ThisDrawing.ActiveSelectionSet
	'Check if anything is selected, and give exit with a warning if not
	If currentSelectionSet.Count = 0 Then
	ThisDrawing.Utility.Prompt "There are no currently selected objects. Please select some
	points to export, and run this command again." & vbNewLine
	End If
	'Use a For Each statement to look through every item in CurrentSelectionSet
	For Each ent In currentSelectionSet
	'In here, ent will be one of the selected entities.
	'If ent is not a point object, we should ignore it
	If TypeOf ent Is AcadPoint Then
	'Only points will make it this far
	'Now that we know we are dealing with a point,
	'we can use the specific AcadPoint type of variable.
	Set pnt = ent
	'You'll notice that after doing this, you have more
	'intellisense methods when you type "pnt."
	'Add a line to the string variable csvFile.
	'We are concatenating two numbers together with a comma in between,
	'and adding a new line character at the end to complete the row.
	csvFile = csvFile & pnt.Coordinates(0) & "," & pnt.Coordinates(1) & vbNewLine
	'Saying that csvFile = csvFile & whatever is a useful
	'way to repeatedly add to the end of a string variable.
	End If
	Next
	'Write the contents of the csvFile variable to a file on the C:\ with the same name
	'FileSystemObjects are really useful for manipulating files
	'But, you'll need a reference to the Microsoft Scripting Runtime in your VBA project.
	'Go Tools>References, and select the Microsoft Scripting Runtime.
	'Create a new File System Object
	Set FSO = New FileSystemObject
	'Using FSO.CreateTextFile, create the text file csvFile.csv,
	'and store a reference to it in the variable textFile
	Set textFile = FSO.CreateTextFile("C:\csvFile.csv")
	'Write the string variable csvFile to textFile
	textFile.Write csvFile
	'Close textFile, as we are finished with it.
	textFile.Close
	'Alert the user that the file has been created
	-1-D:\_S\Excel workbook\CAD_Col\File\junction vba code 1.txt Thursday, May 20, 2021 8:52 PM
	ThisDrawing.Utility.Prompt "Points have been exported to C:\csvFile.csv" & vbNewLine
End Sub

Sub getatts_Extract()
	Dim Excel As Excel.Application
	Dim ExcelSheet As Object
	Dim ExcelWorkbook As Object
	Dim RowNum As Integer
	Dim Header As Boolean
	Dim elem As AcadEntity
	Dim Array1 As Variant
	Dim Count As Integer
	‘ Launch Excel.
	Set Excel = New Excel.Application
	‘ Create a new workbook and find the active sheet.
	Set ExcelWorkbook = Excel.Workbooks.Add
	Set ExcelSheet = Excel.ActiveSheet
	ExcelWorkbook.SaveAs “c:\temp\Attribute.xls”
	RowNum = 1
	Header = False
	‘ Iterate through model space finding
	‘ all block references.
	For Each elem In ThisDrawing.ModelSpace
	With elem
	‘ When a block reference has been found,
	‘ check it for attributes
	If StrComp(.EntityName, “AcDbBlockReference”, 1) _
	= 0 Then
	If .HasAttributes Then
	‘ Get the attributes
	Array1 = .GetAttributes
	‘ Copy the Tagstrings for the
	‘ Attributes into Excel
	For Count = LBound(Array1) To UBound(Array1)
	If Header = False Then
	If StrComp(Array1(Count).EntityName, _
	“AcDbAttribute”, 1) = 0 Then
	ExcelSheet.Cells(RowNum, _
	Count + 1).Value = _
	Array1(Count).TagString
	End If
	End If
	Next Count
	RowNum = RowNum + 1
	For Count = LBound(Array1) To UBound(Array1)
	ExcelSheet.Cells(RowNum, Count + 1).Value _
	= Array1(Count).TextString
	Next Count
	Header = True
	End If
	End If
	End With
	Next elem
	Excel.Application.Quit
End Sub

Private Sub CommandButton2_Click()
	Dim rox,coy as Double
	Dim sname as String
	rox=2:coy=1
	sname="Sheet1"
	MsgBox "- Select Polyline/Polylines from AutoCAD for And Press Enter - ", vbInformation,
	"Select Polyline"
	Call getOtherPolylineCoordinatesFromAutoCAD(rox, coy, sname)
End Sub

Sub getOtherPolylineCoordinatesFromAutoCAD(ByVal cx As Integer, ByVal cy As Integer, ByVal
	sname As String)
	Set NewDC = GetObject(, "AutoCAD.Application")
	Set A2Kdwg = NewDC.ActiveDocument
	Dim Selection As AcadSelectionSet
	Dim poly As AcadLWPolyline
	Dim Obj As AcadEntity
	Dim Bound As Double
	Dim x, y As Double
	Dim rows, i, scount As Integer
	'---Search Object from SelectionSet and Delete If Found ----''
	For i = 0 To A2Kdwg.SelectionSets.count - 1
	If A2Kdwg.SelectionSets.Item(i).Name = "AcDbPolyline" Then
	''-- Delete Object Name from AutoCAD SelectionSet ---''
	A2Kdwg.SelectionSets.Item(i).Delete
	Exit For
	End If
	Next i
	''-- Add Object to AutoCad SelectionSet ----''
	Set Selection = A2Kdwg.SelectionSets.Add("AcDbPolyline")
	''-- Select Object from AutoCad Screen ---'''
	Selection.SelectOnScreen
	''-- Get Coordinates of Object if Object name is ACadPolyline--''
	rows = cx
	For Each Obj In Selection
	If Obj.ObjectName = "AcDbPolyline" Then
	''- Set Obj as Polyline--''
	Set poly = Obj
	On Error Resume Next
	''-- Set Size of Coordinates Like array Size--''
	Bound = UBound(poly.Coordinates)
	'' Starting Index of Excel Row to insert Coordinates --''
	rows = rows
	''-- Display Coordinates one by one to Excel Columns --'''
	For i = 0 To Bound
	''-- Set Coordinates into Variables--'''
	x = Round(poly.Coordinates(i), 3)
	y = Round(poly.Coordinates(i + 1), 3)
	''-- Set Coordinates into Excel Columns --'''
	Worksheets(sname).Cells(rows, cy) = Round(x, 3)
	Worksheets(sname).Cells(rows, cy + 1) = Round(y, 4)
	''- Increment variable for Excel Rows ---''
	rows = rows + 1
	''--- Increment Counter variable to get Next point of Polyline --'''
	i = i + 1
	Next
	Else
	MsgBox "--- This is not a Polyline --- ", vbInformation, "Please Select a
	Polyline"
	End If
	rows = rows + 1
	Next Obj
End Sub

Sub Main()
	Dim xAP As Excel.Application
	Dim xWB As Excel.Workbook
	Dim xWS As Excel.Worksheet
	Set xAP = Excel.Application
	Set xWB = xAP.Workbooks.Open("C:\A2K2_VBA\IUnknown.xls")
	Set xWS = xWB.Worksheets("Sheet1")
	MsgBox "Excel says: """ & Cells(1, 1) & """"
	Dim A2K As AcadApplication
	Dim A2Kdwg As AcadDocument
	Set A2K = CreateObject("AutoCAD.Application")
	Set A2Kdwg = A2K.Application.Documents.Add
	MsgBox A2K.Name & " version " & A2K.Version & _
	" is running."
	Dim Height As Double
	Dim P(0 To 2) As Double
	Dim TxtObj As AcadText
	Dim TxtStr As String
	Height = 1
	P(0) = 1: P(1) = 1: P(2) = 0
	TxtStr = Cells(1, 1)
	Set TxtObj = A2Kdwg.ModelSpace.AddText(TxtStr, _
	P, Height)
	A2Kdwg.SaveAs "C:\A2K2_VBA\IUnknown.dwg"
	A2K.Documents.Close
	A2K.Quit
	Set A2K = Nothing
	xAP.Workbooks.Close
	xAP.Quit
	Set xAP = Nothing
End Sub
	-1-


	'fillet

	D:\_S\Excel workbook\CAD_Col\fillet\filletLines.txt Thursday, May 20, 2021 8:53 PM

Sub FilletLines()
	Dim objLine1 As AcadLine
	Dim objLine2 As AcadLine
	Dim varPt As Variant
	Dim dblFill, newSysVar As Double
	Dim comString As String
	dblFill = ThisDrawing.GetVariable("FILLETRAD")
	On Error GoTo ProblemHere
	MsgBox CStr(dblFill)
	newSysVar = CDbl(InputBox("Enter fillet radius:", "Fillet Radius Value", "2,5"))
	ThisDrawing.Utility.GetEntity objLine1, varPt, "Select first line"
	ThisDrawing.Utility.GetEntity objLine2, varPt, "Select second line"
	If objLine1 Is Nothing Or objLine2 Is Nothing Then
	Exit Sub
	Else
	If TypeOf objLine1 Is AcadLine And _
	TypeOf objLine2 Is AcadLine Then
	ThisDrawing.SetVariable "FILLETRAD", newSysVar
	comString = "_FILLET" & vbCr & "(HANDENT " & Chr(34) & CStr(objLine1.Handle) & Chr(34) &
	")" & vbCr & _
	"(HANDENT " & Chr(34) & CStr(objLine2.Handle) & Chr(34) & ")" & vbCr
	ThisDrawing.SendCommand comString
	Else
	MsgBox "Incorrect object type"
	Exit Sub
	End If
	End If
	ThisDrawing.SetVariable "FILLETRAD", dblFill
	ProblemHere:
	If Err Then
	ThisDrawing.SetVariable "FILLETRAD", dblFill
	MsgBox vbCr & Err.Description
	End If
End Sub

Sub FilletLines()
	Dim objLine1 As AcadLine
	Dim objLine2 As AcadLine
	Dim varPt As Variant
	Dim dblFill, newSysVar As Double
	Dim comString As String
	dblFill = ThisDrawing.GetVariable("FILLETRAD")
	On Error GoTo ProblemHere
	MsgBox CStr(dblFill)
	newSysVar = CDbl(InputBox("Enter fillet radius:", "Fillet Radius Value", "2,5"))
	ThisDrawing.Utility.GetEntity objLine1, varPt, "Select first line"
	ThisDrawing.Utility.GetEntity objLine2, varPt, "Select second line"
	If objLine1 Is Nothing Or objLine2 Is Nothing Then
	Exit Sub
	Else
	If TypeOf objLine1 Is AcadLine And _
	TypeOf objLine2 Is AcadLine Then
	ThisDrawing.SetVariable "FILLETRAD", newSysVar
	comString = "_FILLET" & vbCr & "(HANDENT " & Chr(34) & CStr(objLine1.Handle) & Chr(34) & ")" &
	vbCr & _
	"(HANDENT " & Chr(34) & CStr(objLine2.Handle) & Chr(34) & ")" & vbCr
	-1-D:\_S\Excel workbook\CAD_Col\fillet\filletLines.txt Thursday, May 20, 2021 8:53 PM
	ThisDrawing.SendCommand comString
	Else
	MsgBox "Incorrect object type"
	Exit Sub
	End If
	End If
	ThisDrawing.SetVariable "FILLETRAD", dblFill
	ProblemHere:
	If Err Then
	ThisDrawing.SetVariable "FILLETRAD", dblFill
	MsgBox vbCr & Err.Description
	End If
End Sub


'join_lines

Function Is2DPointsEqual(p1 As Variant, p2 As Variant, gap As Double) As Boolean
	Is2DPointsEqual = False
	Dim a, b
	a = Abs(CDbl(p1(0)) - CDbl(p2(0)))
	b = Abs(CDbl(p1(1)) - CDbl(p2(1)))
	If a <= gap And b <= gap Then Is2DPointsEqual = True
End Function

Sub JoinLines()
	' based on idea by Norman Yuan
	' Fatty T.O.H. () 2007 * all rights removed
	' edited 02.04.2008
	Dim oSsets As AcadSelectionSets
	Dim pSset As AcadSelectionSet
	Dim oSset As AcadSelectionSet
	Dim setName As String
	Dim fType(0) As Integer
	Dim fData(0) As Variant
	Dim varPt As Variant
	Dim pickPt As Variant
	Dim fLine As AcadLine
	Dim oLine As AcadEntity
	Dim oEnt As AcadEntity
	Dim commStr As String
	Dim stPt(1) As Double
	Dim endPt(1) As Double
	Dim dxftype, dxfcode
	Dim n As Integer
	Dim sp As Variant
	Dim ep As Variant
	Dim ps(1) As Double
	Dim pe(1) As Double
	Dim vexs As Variant
	Dim oSpace As AcadBlock
	With ThisDrawing
	If .ActiveSpace = acModelSpace Then
	Set oSpace = .ModelSpace
	Else
	Set oSpace = .PaperSpace
	End If
	End With
	On Error GoTo Error_Trapp
	Dim osm
	osm = ThisDrawing.GetVariable("OSMODE")
	ThisDrawing.SetVariable "OSMODE", 1
	ThisDrawing.SetVariable "PICKBOX", 1
	pickPt = ThisDrawing.Utility.GetPoint(, vbCr & "Select the starting point of the chain of lines
	:")
	ZoomExtents
	Set oSsets = ThisDrawing.SelectionSets
	fType(0) = 0: fData(0) = "LINE"
	dxftype = fType: dxfcode = fData
	setName = "FirstLine"
	With ThisDrawing.SelectionSets
	While .Count > 0
	.Item(0).Delete
	Wend
	-1-D:\_S\Excel workbook\CAD_Col\join_lines\join_lines.txt Thursday, May 20, 2021 8:54 PM
	End With
	setName = "LineSset"
	Set pSset = oSsets.Add("FirstLine")
	pSset.SelectAtPoint pickPt, dxftype, dxfcode
	If pSset.Count > 1 Then
	MsgBox "More than one line selected" & vbCr & _
	"Error"
	Exit Sub
	ElseIf pSset.Count = 1 Then
	Set fLine = pSset.Item(0)
	ElseIf pSset.Count = 0 Then
	MsgBox "Nothing selected" & vbCr & _
	"Error"
	Exit Sub
	End If
	sp = fLine.StartPoint
	ep = fLine.EndPoint
	ps(0) = sp(0): ps(1) = sp(1)
	pe(0) = ep(0): pe(1) = ep(1)
	If Is2DPointsEqual(pickPt, ps, 0.01) Then
	stPt(0) = ps(0): stPt(1) = ps(1)
	endPt(0) = pe(0): endPt(1) = pe(1)
	ElseIf Is2DPointsEqual(pickPt, pe, 0.01) Then
	stPt(0) = pe(0): stPt(1) = pe(1)
	endPt(0) = ps(0): endPt(1) = ps(1)
	End If
	Dim oPline As AcadLWPolyline
	Dim coors(3) As Double
	coors(0) = stPt(0): coors(1) = stPt(1)
	coors(2) = endPt(0): coors(3) = endPt(1)
	Set oPline = oSpace.AddLightWeightPolyline(coors)
	pSset.Delete
	Set pSset = Nothing
	Set oSset = oSsets.Add("LineSset")
	Dim remLine(0) As AcadEntity
	Set remLine(0) = fLine
	oSset.Select acSelectionSetAll, , , dxftype, dxfcode
	oSset.RemoveItems remLine
	fLine.Delete
	Dim i As Long
	i = 1
	Dim Pokey As Boolean
	Pokey = True
	Do Until Not Pokey
	Pokey = False
	Gumby:
	For n = oSset.Count - 1 To 0 Step -1
	Set oLine = oSset.Item(n)
	sp = oLine.StartPoint
	ep = oLine.EndPoint
	ps(0) = sp(0): ps(1) = sp(1)
	pe(0) = ep(0): pe(1) = ep(1)
	If Is2DPointsEqual(ps, endPt, 0.01) Then
	i = i + 1
	oPline.AddVertex i, pe
	Set remLine(0) = oLine
	oSset.RemoveItems remLine
	-2-D:\_S\Excel workbook\CAD_Col\join_lines\join_lines.txt Thursday, May 20, 2021 8:54 PM
	oLine.Delete
	vexs = oPline.Coordinate(i)
	endPt(0) = vexs(0): endPt(1) = vexs(1)
	Pokey = True
	Exit For
	ElseIf Is2DPointsEqual(pe, endPt, 0.01) Then
	i = i + 1
	oPline.AddVertex i, ps
	Set remLine(0) = oLine
	oSset.RemoveItems remLine
	oLine.Delete
	vexs = oPline.Coordinate(i)
	endPt(0) = vexs(0): endPt(1) = vexs(1)
	Pokey = True
	Exit For
	End If
	Next n
	If oSset.Count > 0 Then
	GoTo Gumby
	Else
	Exit Do
	End If
	Loop
	oSset.Delete
	Set oSset = Nothing
	Error_Trapp:
	ZoomPrevious
	If Err.Number <> 0 Then
	MsgBox "Error number: " & Err.Number & vbCr & Err.Description
	End If
	On Error Resume Next
	ThisDrawing.SetVariable "OSMODE", osm
	ThisDrawing.SetVariable "PICKBOX", 4 '<--change size to your suit
End Sub

Public Function MeJoinPline(FstPol As AcadLWPolyline, NxtPol As AcadLWPolyline, _ FuzVal as
	Double) As Boolean
	Dim FstArr() As Double
	Dim NxtArr() As Double
	Dim TmpPnt(0 To 1) As Double
	Dim FstLen As Long
	Dim NxtLen As Long
	Dim VtxCnt As Long
	Dim FstCnt As Long
	Dim NxtCnt As Long
	Dim RevFlg As Boolean
	Dim RetVal As Boolean
	With FstPol FstArr = .Coordinates NxtArr = NxtPol.Coordinates FstLen = UBound(FstArr)
	NxtLen = UBound(NxtArr) '<-Fst<-Nxt
	If MePointsEqual(FstArr, 1, NxtArr, NxtLen, FuzVal)Then
	MeReversePline FstPol FstArr = .Coordinates MeReversePline NxtPol
	NxtArr = NxtPol.Coordinates RevFlg = True RetVal = True '<-FstNxt->
	ElseIf MePointsEqual(FstArr, 1, NxtArr, 1, FuzVal) Then
	MeReversePline FstPol FstArr = .Coordinates
	RevFlg = True RetVal = True 'Fst-><-Nxt
	ElseIf MePointsEqual(FstArr, FstLen, NxtArr, NxtLen, FuzVal) Then
	MeReversePline NxtPol NxtArr = NxtPol.Coordinates
	RevFlg = False RetVal = True 'Fst->Nxt->
	ElseIf MePointsEqual(FstArr, FstLen, NxtArr, 1, FuzVal) Then
	RevFlg = False
	RetVal = True Else
	RetVal = False
	End If
	If RetVal Then
	FstCnt = (FstLen - 1) / 2 NxtCnt = 0 .SetBulge FstCnt,
	NxtPol.GetBulge(NxtCnt)
	For VtxCnt = 2 To NxtLen Step 2
	FstCnt = FstCnt + 1
	NxtCnt = NxtCnt + 1
	TmpPnt(0) = NxtArr(VtxCnt)
	TmpPnt(1) = NxtArr(VtxCnt + 1)
	.AddVertex FstCnt, TmpPnt
	.SetBulge FstCnt, NxtPol
	.GetBulge(NxtCnt)
	Next
	VtxCnt.Update
	NxtPol.Delete
	If RevFlg Then
	MeReversePline FstPol
	End If
	End With
	MeJoinPline = RetVal
	-1-D:\_S\Excel workbook\CAD_Col\join_lines\join_Polyline.txt Thursday, May 20, 2021 8:53 PM
End Function
' -----
Public Function MeReversePline(PolObj As AcadLWPolyline)
	Dim NewArr() As Double
	Dim BlgArr() As Double
	Dim OldArr() As Double
	Dim SegCnt As Long
	Dim ArrCnt As Long
	Dim ArrLen As Long
	With PolObj OldArr =
	.Coordinates ArrLen = UBound(OldArr) SegCnt = (ArrLen - 1) / 2
	ReDim NewArr(0 To ArrLen)
	ReDim BlgArr(0 To SegCnt + 1)
	For ArrCnt = SegCnt To 0 Step -1
	BlgArr(ArrCnt) = .GetBulge(SegCnt - ArrCnt) * -1
	Next ArrCnt
	For ArrCnt = ArrLen To 0 Step -2
	NewArr(ArrLen - ArrCnt + 1) = OldArr(ArrCnt)
	NewArr(ArrLen - ArrCnt) = OldArr(ArrCnt - 1)
	Next
	ArrCnt.Coordinates = NewArr
	For ArrCnt = 0 To SegCnt
	.SetBulge ArrCnt, BlgArr(ArrCnt + 1) Next
	ArrCnt.Update
	End With
End Function
' -----
Public Function MePointsEqual(FstArr, FstPos As Long, NxtArr, NxtPos As Long, _ FuzVal As
	Double) As Boolean
	Dim XcoDst As Double
	Dim YcoDst As Double
	XcoDst = FstArr(FstPos - 1) - NxtArr(NxtPos - 1)
	YcoDst = FstArr(FstPos) - NxtArr(NxtPos)
	MePointsEqual = (Sqr(XcoDst ^ 2 + YcoDst ^ 2) < FuzVal)
End Function

Public Function MeJoinPline(FstPol As AcadLWPolyline, NxtPol As AcadLWPolyline,
	_
	FuzVal as Double) As Boolean
	Dim FstArr() As Double
	Dim NxtArr() As Double
	Dim TmpPnt(0 To 1) As Double
	Dim FstLen As Long
	Dim NxtLen As Long
	Dim VtxCnt As Long
	Dim FstCnt As Long
	Dim NxtCnt As Long
	Dim RevFlg As Boolean
	Dim RetVal As Boolean
	With FstPol
	FstArr = .Coordinates
	NxtArr = NxtPol.Coordinates
	FstLen = UBound(FstArr)
	NxtLen = UBound(NxtArr)
	'<-Fst<-Nxt
	If MePointsEqual(FstArr, 1, NxtArr, NxtLen, FuzVal) Then
	MeReversePline FstPol
	FstArr = .Coordinates
	MeReversePline NxtPol
	NxtArr = NxtPol.Coordinates
	RevFlg = True
	RetVal = True
	'<-FstNxt->
	ElseIf MePointsEqual(FstArr, 1, NxtArr, 1, FuzVal) Then
	MeReversePline FstPol
	FstArr = .Coordinates
	RevFlg = True
	RetVal = True
	'Fst-><-Nxt
	ElseIf MePointsEqual(FstArr, FstLen, NxtArr, NxtLen, FuzVal) Then
	MeReversePline NxtPol
	NxtArr = NxtPol.Coordinates
	RevFlg = False
	RetVal = True
	'Fst->Nxt->
	ElseIf MePointsEqual(FstArr, FstLen, NxtArr, 1, FuzVal) Then
	RevFlg = False
	RetVal = True
	Else
	RetVal = False
	End If
	If RetVal Then
	FstCnt = (FstLen - 1) / 2
	NxtCnt = 0
	.SetBulge FstCnt, NxtPol.GetBulge(NxtCnt)
	For VtxCnt = 2 To NxtLen Step 2
	FstCnt = FstCnt + 1
	NxtCnt = NxtCnt + 1
	TmpPnt(0) = NxtArr(VtxCnt)
	TmpPnt(1) = NxtArr(VtxCnt + 1)
	.AddVertex FstCnt, TmpPnt
	.SetBulge FstCnt, NxtPol.GetBulge(NxtCnt)
	Next VtxCnt
	.Update
	NxtPol.Delete
	If RevFlg Then MeReversePline FstPol
	End If
	End With
	-1-D:\_S\Excel workbook\CAD_Col\join_lines\join_Polyline_2.txt Thursday, May 20, 2021 8:53 PM
	MeJoinPline = RetVal
End Function
' -----
Public Function MeReversePline(PolObj As AcadLWPolyline)
	Dim NewArr() As Double
	Dim BlgArr() As Double
	Dim OldArr() As Double
	Dim SegCnt As Long
	Dim ArrCnt As Long
	Dim ArrLen As Long
	With PolObj
	OldArr = .Coordinates
	ArrLen = UBound(OldArr)
	SegCnt = (ArrLen - 1) / 2
	ReDim NewArr(0 To ArrLen)
	ReDim BlgArr(0 To SegCnt + 1)
	For ArrCnt = SegCnt To 0 Step -1
	BlgArr(ArrCnt) = .GetBulge(SegCnt - ArrCnt) * -1
	Next ArrCnt
	For ArrCnt = ArrLen To 0 Step -2
	NewArr(ArrLen - ArrCnt + 1) = OldArr(ArrCnt)
	NewArr(ArrLen - ArrCnt) = OldArr(ArrCnt - 1)
	Next ArrCnt
	.Coordinates = NewArr
	For ArrCnt = 0 To SegCnt
	.SetBulge ArrCnt, BlgArr(ArrCnt + 1)
	Next ArrCnt
	.Update
	End With
End Function
' -----
Public Function MePointsEqual(FstArr, FstPos As Long, NxtArr, NxtPos As Long, _
	FuzVal As Double) As Boolean
	Dim XcoDst As Double
	Dim YcoDst As Double
	XcoDst = FstArr(FstPos - 1) - NxtArr(NxtPos - 1)
	YcoDst = FstArr(FstPos) - NxtArr(NxtPos)
	MePointsEqual = (Sqr(XcoDst ^ 2 + YcoDst ^ 2) < FuzVal)
End Function

Private Sub CommandButtonSmooth_Click()
	Dim sset As AcadSelectionSet
	Dim v(0) As Variant
	Dim lifiltertype(0) As Integer
	Dim plineObj As AcadLWPolyline
	Dim oLWP As AcadLWPolyline
	Dim i As Long
	Dim var As Variant
	Dim oSS() As AcadEntity
	Dim oGr As AcadGroup
	Set oGr = ThisDrawing.Groups.Add("QWERT")
	Set sset = Nothing
	For i = 0 To ThisDrawing.SelectionSets.Count - 1
	Set sset = ThisDrawing.SelectionSets.Item(i)
	If sset.Name = "ss1" Then
	sset.Clear
	Exit For
	Else
	Set sset = Nothing
	End If
	Next i
	If sset Is Nothing Then
	Set sset = ThisDrawing.SelectionSets.Add("ss1")
	End If
	'create a selection set of all the entities on a given layer
	'here they are all lw polylines
	lifiltertype(0) = 8
	v(0) = "jhl_9.25_begin"
	sset.Select acSelectionSetAll, , , lifiltertype, v
	ReDim Preserve oSS(0 To sset.Count - 1) As AcadEntity
	For i = 0 To sset.Count - 1
	Set oSS(i) = sset.Item(i)
	Next i
	'add plines to group
	oGr.AppendItems oSS
	Dim GRname As String
	GRname = oGr.Name
	' using SendCommand method with Group
	ThisDrawing.SendCommand "_PEDIT" & vbCr & "M" & vbCr & "G" & vbCr & GRname & vbCr & vbCr &
	"J" & vbCr & "0.0" & vbCr & vbCr
	' deleting group and clearing selection set
	oGr.Delete
	sset.Clear
	'start to pedit spline or fit here
	Dim oGroup As AcadGroup
	Set oGroup = ThisDrawing.Groups.Add("ZERO")
	GRname = oGroup.Name
	'select all the joined plines
	sset.Select acSelectionSetAll, , , lifiltertype, v
	ReDim Preserve oSS(0) As AcadEntity
	For i = 0 To sset.Count - 1
	Set oLWP = sset.Item(i)
	Set oSS(0) = sset.Item(i)
	'create group with one item
	oGroup.AppendItems oSS
	If oLWP.Closed Then
	'Spline
	ThisDrawing.SendCommand "_PEDIT" & vbCr & "M" & vbCr & "G" & vbCr & GRname & vbCr &
	vbCr & "S" & vbCr & vbCr
	Else
	'Fit
	ThisDrawing.SendCommand "_PEDIT" & vbCr & "M" & vbCr & "G" & vbCr & GRname & vbCr &
	vbCr & "F" & vbCr & vbCr
	End If
	'remove the pline from the group
	oGroup.RemoveItems oSS
	-1-D:\_S\Excel workbook\CAD_Col\join_lines\join_spline.txt Thursday, May 20, 2021 8:53 PM
	Next i
	sset.Delete
	oGroup.Delete
End Sub
-2-


'selection

'draw circle by selection of excel ranges....
Public Sub TestAddCircle()
	Dim varCenter(0 To 2) As Double
	Dim dblRadius As Double
	Dim objEnt As AcadCircle
	On Error Resume Next
	dblRadius = 0.45
	e = Selection.Row
	For Each x In Selection
	y = Selection.Column
	varCenter(0) = Cells(e, y)
	varCenter(1) = Cells(e, y + 1)
	Set objEnt = ActiveDocument.ModelSpace.AddCircle(varCenter, dblRadius)
	objEnt.Update
	e = e + 1
	Next
End Sub

sub dfd()
	'get point...
	For Each x In Selection
	With ActiveDocument.Utility
	varPick = .GetPoint(, vbCr & "Pick a point: ")
	End With
	Cells(i, j) = varPick(0)
	Cells(i, j + 1) = varPick(1) '= Cells(e, j)
	i = i + 1
	Next
	'get entity...
	Dim ent As AcadEntity
	Dim textEnt As AcadText
	Dim mtextEnt As AcadMText
	Dim objSS As AcadEntity
	With ActiveDocument.Utility
	.GetEntity objSS, varPick, vbCr & "Pick a text near "
	End With
	If TypeOf ent Is AcadText Then
	Set textEnt = ent
	n = textEnt.InsertionPoint(0)
	y = textEnt.InsertionPoint(1)
	y = n + y
	Cells(4 + n, 1).Value = y
	'' Do whatever wuth the AcadText object
	ElseIf TypeOf ent Is AcadMText Then
	Set mtextEnt = ent
	n = textEnt.InsertionPoint(0)
	y = textEnt.InsertionPoint(1)
	Cells(4 + n, 1).Value = c
	End If

end sub

Public Sub TestGetSubEntity()
	Dim objEnt As AcadEntity
	Dim varPick As Variant
	Dim varMatrix As Variant
	Dim varParents As Variant
	Dim intI As Integer
	Dim intJ As Integer
	Dim varID As Variant
	With ThisDrawing.Utility
	'' get the subentity from the user
	.GetSubEntity objEnt, varPick, varMatrix, varParents, _
	vbCr & "Pick an entity: "
	'' print some information about the entity
	.Prompt vbCr & "You picked a " & objEnt.ObjectName
	.Prompt vbCrLf & "At " & varPick(0) & "," & varPick(1)
	'' dump the varMatrix
	If Not IsEmpty(varMatrix) Then
	.Prompt vbLf & "MCS to WCS Translation varMatrix:"
	'' format varMatrix row
	For intI = 0 To 3
	.Prompt vbLf & "["
	'' format varMatrix column
	For intJ = 0 To 3
	.Prompt "(" & varMatrix(intI, intJ) & ")"
	Next intJ
	.Prompt "]"
	Next intI
	.Prompt vbLf
	End If
	'' if it has a parent nest
	If Not IsEmpty(varParents) Then
	.Prompt vbLf & "Block nesting:"
	'' depth counter
	intI = -1
	'' traverse most to least deep (reverse order)
	For intJ = UBound(varParents) To LBound(varParents) Step -1
	'' increment depth
	intI = intI + 1
	'' indent output
	.Prompt vbLf & Space(intI * 2)
	'' parent object ID
	varID = varParents(intJ)
	'' parent entity
	Set objEnt = ThisDrawing.ObjectIdToObject(varID)
	'' print info about parent
	.Prompt objEnt.ObjectName & " : " & objEnt.Name
	Next intJ
	.Prompt vbLf
	End If
	.Prompt vbCr
	End With
End Sub

Sub draw()
	'pick the start point to draw the slab...
	pto = AutoCAD.Application.ActiveDocument.Utility.GetPoint(, vbCr & "pick the first point ")
	x = pto(0): y = pto(1)
	Dim plineObj As AcadPolyline
	Dim plineObj As AcadPolyline
	Dim pta(0 To 14) As Double
	pta(0) = x: pta(1) = y
	' Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	pta(3) = x + 52.101: pta(4) = y
	' Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	pta(6) = x + 52.101: pta(7) = y + 21.752
	' Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	pta(9) = x: pta(10) = y + 21.752
	' Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	pta(12) = x: pta(13) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	plineObj.Color = cyan
	plineObj.color=cyan
	'draw hz
	Dim ptb(0 To 5) As Double
	ptb(0) = x: ptb(1) = y+1.472
	ptb(3) = x+52.101: ptb(4) = y+1.472
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(ptb)
	plineObj.color=color8
	'the diagonal
	Dim ptc(0 To 5) As Double
	ptc(0) = x: ptc(1) = y
	ptc(3) = x+2.049: ptc(4) = y+1.472
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(ptc)
	plineObj.color=color8
	'draw the horizontal
	Dim ptd(0 To 5) As Double
	ptd(0) = x+1.648: ptd(1) = y+1.472
	ptd(3) = x+1.648+50.420: ptd(4) = y+1.472
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(ptd)
	plineObj.color=cyan
	Dim pte(0 To 5) As Double
	pte(0) = x+52.101-0.433: pte(1) = y+1.072
	pte(3) = x+52.101-0.433: pte(4) = y+1.072+20.648
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pte)
	plineObj.color=cyan
	-1-D:\_S\Excel workbook\CAD_Col\selection\polyline.txt Thursday, May 20, 2021 8:56 PM
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	-2-D:\_S\Excel workbook\CAD_Col\selection\polyline.txt Thursday, May 20, 2021 8:56 PM
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
	Dim pte(0 To 17) As Double
	pta(0) = x: pta(1) = y
	Set plineObj = ActiveDocument.ModelSpace.AddPolyline(pta)
End Sub

Sub selEntByPline()
	On Error Resume Next
	Dim objCadEnt As AcadEntity
	Dim vrRetPnt As Variant
	ThisDrawing.Utility.GetEntity objCadEnt, vrRetPnt
	If objCadEnt.ObjectName = "AcDbPolyline" Then '|-- Checking for 2D Polylines --|
	Dim objLWPline As AcadLWPolyline
	Dim objSSet As AcadSelectionSet
	Dim dblCurCords() As Double
	Dim dblNewCords() As Double
	Dim iMaxCurArr, iMaxNewArr As Integer
	Dim iCurArrIdx, iNewArrIdx, iCnt As Integer
	Set objLWPline = objCadEnt
	dblCurCords = objLWPline.Coordinates '|-- The returned coordinates are 2D only --|
	iMaxCurArr = UBound(dblCurCords)
	If iMaxCurArr = 3 Then
	ThisDrawing.Utility.Prompt "The selected polyline should have minimum 2 segments..."
	Exit Sub
	Else
	'|-- The 2D Coordinates are insufficient to use in SelectByPolygon method --|
	'|-- So convert those into 3D coordinates --|
	iMaxNewArr = ((iMaxCurArr + 1) * 1.5) - 1 '|-- New array dimension
	ReDim dblNewCords(iMaxNewArr) As Double
	iCurArrIdx = 0: iCnt = 1
	For iNewArrIdx = 0 To iMaxNewArr
	If iCnt = 3 Then '|-- The z coordinate is set to 0 --|
	dblNewCords(iNewArrIdx) = 0
	iCnt = 1
	Else
	dblNewCords(iNewArrIdx) = dblCurCords(iCurArrIdx)
	iCurArrIdx = iCurArrIdx + 1
	iCnt = iCnt + 1
	End If
	Next
	Set objSSet = ThisDrawing.SelectionSets.Add("SEL_ENT")
	objSSet.SelectByPolygon acSelectionSetWindowPolygon, dblNewCords
	objSSet.Highlight True
	objSSet.Delete
	End If
	Else
	ThisDrawing.Utility.Prompt "The selected object is not a 2D Polyline...."
	End If
	If Err.Number <> 0 Then
	MsgBox Err.Description
	Err.Clear
	End If
End Sub
'******** Text ************************************************

Public Sub TestAddText()
	Dim varStart As Variant
	Dim dblHeight As Double
	Dim strText As String
	Dim objEnt As AcadText
	Dim textStyle1 As AcadTextStyle
	Dim newFontFile As String
	On Error Resume Next
	'' get input from user
	With activedocument.Utility
	varStart = .GetPoint(, vbCr & "Pick the start point: ")
	End With
	dblHeight = 0.5
	Stop
	'convert the text style .... to want of what we aim to do...
	Set textStyle1 = activedocument.ActiveTextStyle
	' Change the value for FontFile
	newFontFile = "C:/AutoCAD/Fonts/simplex.shx"
	textStyle1.fontFile = newFontFile
	' iterate through texts
	For Each x In Range("h5", "h14")
	strText = x
	'' create the text
	Set objEnt = activedocument.ModelSpace.AddText(strText, varStart, dblHeight)
	objEnt.Update
	objEnt.Color = acCyan
	''iterate through the pick point
	varStart(0) = varStart(0): varStart(1) = varStart(1) - 0.7863
	Next
	Stop
	'quantity with numbers....
	For Each y In Range("i34", "i34")
	Next
	Stop
End Sub
'******************* END *************************************
'******************** text style convert to ...simplex.shx ************************************
Sub Example_FontFile()
	' This example returns the current setting of
	' the FontFile property. It then changes the value, and
	' finally resets the value back to the original setting.
	Dim textStyle1 As AcadTextStyle
	Dim currFontFile As String
	Dim newFontFile As String
	Set textStyle1 = ThisDrawing.ActiveTextStyle
	' Retrieve the current FontFile value
	-1-D:\_S\Excel workbook\CAD_Col\selection\text_2CAD_fm_Excel.txt Thursday, May 20, 2021 8:56 PM
	currFontFile = textStyle1.fontFile
	MsgBox "The current value for FontFile is " & currFontFile, vbInformation, "FontFile
	Example"
	' Change the value for FontFile
	newFontFile = "C:/AutoCAD/Fonts/italic.shx"
	textStyle1.fontFile = newFontFile
	MsgBox "The new value for FontFile is " & textStyle1.fontFile, vbInformation, "FontFile
	Example"
	' Reset font file
	textStyle1.fontFile = currFontFile
	MsgBox "The value for FontFile has been reset to " & textStyle1.fontFile,
	vbInformation, "FontFile Example"
End Sub
'********************************************************

'trim


'============================================================================================
'Break_Line main program
'============================================================================================
Function BL()
	On Error Resume Next
	Set osel = SelectionSetFromScreen
	Set osel2 = osel
	kk = 0
	For Each ent1 In osel
	strs = ""
	For Each ent2 In osel2
	xp = ""
	xp = double2string(ent1.IntersectWith(ent2, acExtendNone))
	If xp <> "" Then
	strs = strs & xp & vbCrLf
	End If
	Next
	strs = strs & double2string(ent1.EndPoint)
	x0 = ent1.StartPoint
	cc = string2doublen(strs, 3)
	strs = ""
	For i = 0 To UBound(cc, 1)
	strs = strs & VectorSize(x0, string2double(cc(i, 0) & " " & cc(i, 1) & " " & cc(i, 2))) & vbCrLf
	Next i
	aa = sortArray(string2double(strs))
	For i = 0 To UBound(cc, 1)
	k0 = aa(i)
	x1 = string2double(cc(k0, 0) & " " & cc(k0, 1) & " " & cc(k0, 2))
	If VectorSize(x0, x1) > 0.0000000001 Then
	Set lineObj = ActiveDocument.ModelSpace.AddLine(x0, x1)
	x0 = x1
	Else
	x0 = x0
	End If
	Next i
	Next
	For Each ent In osel
	ent.Delete
	Next
	ActiveDocument.Regen (True)
End Function
'============================================================================================
'return selection set
'============================================================================================
Function SelectionSetFromScreen(Optional ww As String) As AcadSelectionSet
	-1-D:\_S\Excel workbook\CAD_Col\trim\Break_at_intersection.bas Thursday, May 20, 2021 9:02 PM
	On Error Resume Next
	selsname = "S_" & IntRandomCall(1, 10000000#)
	ActiveDocument.SelectionSets(selsname).Delete
	Set SelectionSetFromScreen = ActiveDocument.SelectionSets.Add(selsname)
	SelectionSetFromScreen.SelectOnScreen
End Function
'============================================================================================
'change double array to string
'============================================================================================
Function double2string(ar) As String
	For i = 0 To UBound(ar)
	double2string = double2string & ar(i) & " "
	Next
	double2string = RTrim(double2string)
End Function
'============================================================================================
'vector size
'============================================================================================
Function VectorSize(p0, p1) As Double 'vector_size vector size
	a1 = p1(0) - p0(0)
	a2 = p1(1) - p0(1)
	a3 = p1(2) - p0(2)
	VectorSize = (a1 ^ 2 + a2 ^ 2 + a3 ^ 2) ^ 0.5
End Function
'============================================================================================
'array sort
'============================================================================================
Function sortArray(a, Optional identifier = "up") As Long()
	n1 = LBound(a)
	n2 = UBound(a)
	Dim nn() As Long
	ReDim nn(n2) As Long
	For i = n1 To n2
	nn(i) = i
	Next i
	If LCase(identifier) = "dn" Then
	For i = n1 To n2
	a0 = a(i)
	For j = n1 To n2
	a1 = a(j)
	If a0 > a1 Then
	a(i) = a1
	-2-D:\_S\Excel workbook\CAD_Col\trim\Break_at_intersection.bas Thursday, May 20, 2021 9:02 PM
	a(j) = a0
	a0 = a1
	kk = nn(j)
	nn(j) = nn(i)
	nn(i) = kk
	End If
	Next j
	Next i
	Else
	For i = n1 To n2
	a0 = a(i)
	For j = n1 To n2
	a1 = a(j)
	If a0 < a1 Then
	a(i) = a1
	a(j) = a0
	a0 = a1
	kk = nn(j)
	nn(j) = nn(i)
	nn(i) = kk
	End If
	Next j
	Next i
	End If
	sortArray = nn
End Function
'============================================================================================
'split string to array
'============================================================================================
Function string2double(ByVal ss As String) As Double()
	On Error Resume Next
	Dim cc() As Double
	ss = Replace(ss, vbTab, " ")
	ss = Replace(ss, ",", " ")
	ss = Replace(ss, ";", " ")
	ss = Replace(ss, "(", " ")
	ss = Replace(ss, ")", " ")
	ss = Replace(ss, "\", " ")
	ss = Replace(ss, "|", " ")
	ss = Replace(ss, vbCrLf, " ")
	qq = Split(ss)
	kk = 0
	For i = 0 To UBound(qq)
	If IsNumeric(qq(i)) Then
	kk = kk + 1
	End If
	Next i
	-3-D:\_S\Excel workbook\CAD_Col\trim\Break_at_intersection.bas Thursday, May 20, 2021 9:02 PM
	ReDim cc(kk - 1)
	kk = 0
	For i = 0 To UBound(qq)
	If IsNumeric(qq(i)) Then
	cc(kk) = qq(i)
	kk = kk + 1
	End If
	Next i
	string2double = cc
	Exit Function
	'errh:
	' MsgBox "ERR!!!"
End Function
'============================================================================================
'split string to n-dimensional array
'============================================================================================
Function string2doublen(strs, n) As Double()
	ss = string2double(strs)
	n1 = LBound(ss)
	n2 = UBound(ss)
	nn = (n2 - n1 + 1) / n
	Dim cc() As Double
	ReDim cc(nn - 1, n - 1) As Double
	kk = 0
	For i = 0 To nn - 1
	For j = 0 To n - 1
	cc(i, j) = ss(kk)
	kk = kk + 1
	Next
	Next
	string2doublen = cc
End Function
'============================================================================================
'randomize integer
'============================================================================================
Function IntRandomCall(n1, n2)
	Randomize
	IntRandomCall = Int((n2 - n1 + 1) * Rnd + n1)
End Function

Public Sub Trim()
	Dim objEnt As AcadEntity
	Dim objCut As AcadCircle
	Dim objTrim As AcadLine
	Dim varPnt(0 To 2) As Double
	Dim varSPnt As Variant
	Dim varEPnt As Variant
	Dim strPrmpt As String
	Dim varTrimPnt As Variant
	Dim dblTrimPnt(2) As Double
	Dim varInterSectns As Variant
	Dim GetLength
	On Error GoTo Err_Control
	'Line to cut to
	varPnt(0) = 0
	varPnt(1) = 0
	varPnt(2) = 0
	ActiveDocument.Utility.GetEntity objCut, Array(varPnt(0), varPnt(1), varPnt(2))
	'Do
	'line to trim
	ActiveDocument.Utility.GetEntity objEnt, varPnt, strPrmpt
	'trimming
	If TypeOf objEnt Is AcadLine Then
	Set objTrim = objEnt
	varInterSectns = objTrim.IntersectWith(objCut, acExtendNone)
	If IsArray(varInterSectns) Then
	If UBound(varInterSectns) > 0 Then
	varSPnt = objTrim.StartPoint
	varEPnt = objTrim.EndPoint
	dblTrimPnt(0) = varInterSectns(0)
	dblTrimPnt(1) = varInterSectns(1)
	dblTrimPnt(2) = varInterSectns(2)
	varTrimPnt = Array(varInterSectns(0), varInterSectns(1), varInterSectns(2))
	' objTrim.startPoint = dblTrimPnt
	objTrim.EndPoint = dblTrimPnt
	End If
	End If
	End If
	'Loop
	Exit_Here:
	If Not objCut Is Nothing Then
	objCut.Highlight False
	End If
	Exit Sub
	Err_Control:
	'If they select anything other than A line
	If Err.Description = "Type mismatch" Then
	-1-D:\_S\Excel workbook\CAD_Col\trim\trim_4_circle.bas Thursday, May 20, 2021 9:02 PM
	Err.Clear
	End If
End Sub

Sub trim_Lay()
	Call cutlines(20, "0")
End Sub

Sub cutlines(Trim As Integer, LayerName As String)
	Dim CurrentLine As AcadLine
	Dim Oldstartpoint() As Double
	Dim NewStartpoint(0 To 2) As Double
	Dim Oldendpoint() As Double
	Dim Newendpoint(0 To 2) As Double
	'loops through every object in the current layout
	For a = 0 To (ActiveDocument.ActiveLayout.Block.Count - 1)
	'checks whether object is a line
	If ActiveDocument.ActiveLayout.Block.Item(a).ObjectName = "AcDbLine" Then
	'puts the current line in an object variable for easy writing
	Set CurrentLine = ActiveDocument.ActiveLayout.Block.Item(a)
	'checks whether line is in correct layer
	If CurrentLine.Layer = LayerName Then
	Oldstartpoint = CurrentLine.StartPoint
	Oldendpoint = CurrentLine.EndPoint
	'formulae that determine the new startpoints and endpoints of the line
	For i = 0 To 2
	NewStartpoint(i) = Oldstartpoint(i) + (CurrentLine.Delta(i) * (Trim / 2) /
	CurrentLine.Length)
	Newendpoint(i) = Oldendpoint(i) - (CurrentLine.Delta(i) * (Trim / 2) /
	CurrentLine.Length)
	Next
	CurrentLine.StartPoint = NewStartpoint
	CurrentLine.EndPoint = Newendpoint
	End If
	End If
	Next
End Sub

Sub Closed_PLine_From_IntPts()
	Dim pL As AcadLWPolyline
	Dim rL As AcadLWPolyline
	Dim PlCopy As AcadLWPolyline
	Set pL = ThisDrawing.ModelSpace.Item(0)
	Set rL = ThisDrawing.ModelSpace.Item(1)
	Dim IntPt
	Set PlCopy = pL.Copy 'Copy Polyline to keep draworder at the top and to trim
	IntPt = pL.IntersectWith(rL, acExtendOtherEntity)
	'Trim using Break command---------------------------------------------
	Dim Pt1, Pt2
	Pt1 = IntPt(0) & "," & IntPt(1) & "," & IntPt(2)
	Pt2 = pL.Coordinates(0) & "," & pL.Coordinates(1) & "," & pL.Elevation
	ThisDrawing.SendCommand "break" & vbCr & Pt1 & vbCr & Pt2 & vbCr
	Pt1 = IntPt(3) & "," & IntPt(4) & "," & IntPt(5)
	Pt2 = pL.Coordinates(UBound(pL.Coordinates) - 1) & "," & pL.Coordinates(UBound(pL.Coordinates))
	& "," & pL.Elevation
	ThisDrawing.SendCommand "break" & vbCr & Pt1 & vbCr & Pt2 & vbCr
	'--------------------------------------------------------------------
	'Add Vertices from the similar points-------------------------------
	Dim Vert(0 To 1) As Double
	If Format(Abs(rL.Coordinates(0) - PlCopy.Coordinates(0)), "0.000000000") = 0 And
	Format(Abs(rL.Coordinates(1) - PlCopy.Coordinates(1)), "0.000000000") = 0 Then
	For i = 2 To UBound(PlCopy.Coordinates) - 2 Step 2
	Vert(0) = PlCopy.Coordinates(i): Vert(1) = PlCopy.Coordinates(i + 1)
	rL.AddVertex 0, Vert
	Next i
	ElseIf Format(Abs(rL.Coordinates(6) - PlCopy.Coordinates(0)), "0.000000000") = 0 And
	Format(Abs(rL.Coordinates(7) - PlCopy.Coordinates(1)), "0.000000000") = 0 Then
	For i = 2 To UBound(PlCopy.Coordinates) - 2 Step 2
	Vert(0) = PlCopy.Coordinates(i): Vert(1) = PlCopy.Coordinates(i + 1)
	rL.AddVertex ((UBound(rL.Coordinates)) - 1) / 2 + 1, Vert
	Next i
	End If
	rL.Closed = True
	'Delete copied pline
	PlCopy.Delete
End Sub

Sub stub()
	Dim cirs(1) As AcadCircle
	Dim cir1Cen(2) As Double
	Dim reg
	cir1Cen(0) = 10
	cir1Cen(1) = 10
	cir1Cen(2) = 0
	Dim cir2Cen(2) As Double
	cir2Cen(0) = 15
	cir2Cen(1) = 10
	cir2Cen(2) = 0
	Dim cirRad As Double: cirRad = 5
	'create an overlaping circles
	With ActiveDocument.ModelSpace
	Set cirs(0) = .AddCircle(cir1Cen, cirRad)
	Set cirs(1) = .AddCircle(cir2Cen, cirRad)
	reg = .AddRegion(cirs)
	Dim regs(1) As AcadRegion
	Set regs(0) = reg(0)
	Set regs(1) = reg(1)
	regs(0).Boolean acIntersection, regs(1)
	cirs(0).Delete
	cirs(1).Delete
	End With
End Sub

Public Sub BreakLineByBlock()
	'Break lines around block insertions.
	Dim str As String
	Dim strHandle As String
	Dim objLine As AcadLine
	Dim objLine1 As AcadLine
	Dim objLine2 As AcadLine
	Dim objSubEnt As AcadEntity
	Dim objBlock As AcadBlockReference
	Dim ssBlocks As AcadSelectionSet
	Dim ssLines As AcadSelectionSet
	Dim vSubEnts As Variant
	Dim vMinPoint As Variant
	Dim vMaxPoint As Variant
	Dim vIntPoint As Variant
	Dim vCPoint As Variant 'compare point
	Dim vSPoint As Variant 'start point
	Dim vSPoint1 As Variant 'start point prime
	Dim vEPoint As Variant 'end point
	Dim vEPoint1 As Variant 'end point prime
	Dim dPickPoint(0 To 1) As Double
	Dim dPoint(0 To 2) As Double
	Dim dDistSP As Double 'shortest distance from start point
	Dim dDistEP As Double 'shortest distance from end point
	Dim dDistC As Double 'comparison distance
	Dim dVertList(0 To 7) As Double
	Dim iL As Integer 'lines counter
	Dim iP As Integer 'points counter
	Dim iSE As Integer 'sub entities counter
	Dim iCntL As Integer 'line count
	Dim iCntP As Integer 'point count
	Dim iCntSE As Integer 'sub entity count
	Dim PtsInsideBB As Integer '0=none: 1=StartPoint: 2=EndPoint
	Dim varFilterType(0) As Integer
	Dim varFilterData(0) As Variant
	Dim vFT As Variant
	Dim vFD As Variant
	Dim BBpoints(0 To 4) As Point 'Bounding box points list
	Dim Cpoint As Point 'compare point
	' On Error GoTo Err_Control
	'Set up undo for this command
	ActiveDocument.StartUndoMark
	'get blocks
	ActiveDocument.Utility.Prompt "Lines will be broken around selected blocks."
	Set ssBlocks = toolbox.ejSelectionSets.GetSS_BlockFilter
	For Each objBlock In ssBlocks
	'Use the block's bounding box to select ents that intersect with it.
	objBlock.GetBoundingBox vMinPoint, vMaxPoint
	BBpoints(0).x = vMinPoint(0): BBpoints(0).y = vMinPoint(1)
	BBpoints(1).x = vMaxPoint(0): BBpoints(1).y = vMinPoint(1)
	BBpoints(2).x = vMaxPoint(0): BBpoints(2).y = vMaxPoint(1)
	BBpoints(3).x = vMinPoint(0): BBpoints(3).y = vMaxPoint(1)
	BBpoints(4).x = vMinPoint(0): BBpoints(4).y = vMinPoint(1)
	Set ssLines = toolbox.ejSelectionSets.AddSelectionSet("ssLines")
	ssLines.Clear
	varFilterType(0) = 0: varFilterData(0) = "LINE"
	vFT = varFilterType: vFD = varFilterData
	ssLines.Select acSelectionSetCrossing, vMaxPoint, vMinPoint, vFT, vFD
	'get subent's of block
	vSubEnts = objBlock.Explode
	iCntSE = UBound(vSubEnts)
	For Each objLine In ssLines
	'Compare subentity intersection points with line start and
	-1-D:\_S\Excel workbook\CAD_Col\trim\trim_ED_AUgi.bas Thursday, May 20, 2021 9:01 PM
	'end points to determine new line segment. Points creating the
	'shortest line segments should be the outer limits of the block.
	'Any other intersections are inside the block and are discarded.
	' Get reference info.
	vSPoint = objLine.StartPoint
	vEPoint = objLine.EndPoint
	dDistSP = toolbox.ejMath.XYZDistance(vSPoint, vEPoint)
	dDistEP = toolbox.ejMath.XYZDistance(vEPoint, vSPoint)
	Cpoint.x = vSPoint(0): Cpoint.y = vSPoint(1)
	If toolbox.ejMath.InsidePolygon(BBpoints, Cpoint) = True Then PtsInsideBB =
	PtsInsideBB Or 1
	Cpoint.x = vEPoint(0): Cpoint.y = vEPoint(1)
	If toolbox.ejMath.InsidePolygon(BBpoints, Cpoint) = True Then PtsInsideBB =
	PtsInsideBB Or 2
	For iSE = 0 To iCntSE
	'get list of points where the line intersects with the block
	Set objSubEnt = vSubEnts(iSE)
	vIntPoint = objSubEnt.IntersectWith(objLine, acExtendNone)
	'Compare to line segment lengths.
	If UBound(vIntPoint) > -1 Then
	iCntP = (UBound(vIntPoint) + 1) / 3
	For iP = 1 To iCntP
	vCPoint = toolbox.ejMath.Point3D((vIntPoint(iP * 3 - 3)), (vIntPoint(iP
	* 3 - 2)), (vIntPoint(iP * 3 - 1)))
	dDistC = toolbox.ejMath.XYZDistance(vSPoint, vCPoint)
	If dDistC < dDistSP Then
	dDistSP = dDistC
	vSPoint1 = vCPoint
	End If
	dDistC = toolbox.ejMath.XYZDistance(vCPoint, vEPoint)
	If dDistC < dDistEP Then
	dDistEP = dDistC
	vEPoint1 = vCPoint
	End If
	Next iP
	Else
	'the array returned by IntersectWith is dimensioned
	' (0 To -1) when there are no points.
	End If
	Next iSE
	Select Case Round(objLine.Length, 14)
	Case Is = Round(dDistSP, 14)
	'line did not intersect the block
	'do nothing
	Case Is = Round(dDistSP + dDistEP, 14)
	'One end of the line is inside the block and does
	'not pass through, only one intersection point.
	'Determine whether start point or end
	'point is in the block and trim it. Assume the smaller
	'distance is inside the block.
	If dDistSP > dDistEP Then
	'the endpoint is in the block
	objLine.EndPoint = vEPoint1
	objLine.Update
	Else
	'the startpoint is in the block
	objLine.StartPoint = vSPoint1
	objLine.Update
	End If
	Case Else
	'enough intersection points exist to break the line
	'create two new lines and delete the original
	Select Case PtsInsideBB
	Case Is = 0 'neither end is inside
	If ActiveDocument.ActiveSpace = acModelSpace Then
	Set objLine1 = ActiveDocument.ModelSpace.AddLine(vSPoint,
	-2-D:\_S\Excel workbook\CAD_Col\trim\trim_ED_AUgi.bas Thursday, May 20, 2021 9:01 PM
	vSPoint1)
	Set objLine2 = ActiveDocument.ModelSpace.AddLine(vEPoint1,
	vEPoint)
	Else
	Set objLine1 = ActiveDocument.PaperSpace.AddLine(vSPoint,
	vSPoint1)
	Set objLine2 = ActiveDocument.PaperSpace.AddLine(vEPoint1,
	vEPoint)
	'update new lines so that they will be seen by the next attempt
	to
	'get a selection set
	End If
	objLine1.Update
	objLine2.Update
	objLine.Delete
	Case Is = 1 'start point is inside
	objLine.StartPoint = vEPoint1
	objLine.Update
	Case Is = 2 'end point is inside
	objLine.EndPoint = vSPoint1
	objLine.Update
	Case Is = 3 'both ends are inside
	End Select
	PtsInsideBB = 0 'reset for next line
	End Select
	Next objLine
	For iSE = 0 To iCntSE
	Set objSubEnt = vSubEnts(iSE)
	objSubEnt.Delete
	Next iSE
	Next objBlock
	Exit_Here:
	ActiveDocument.EndUndoMark
	Exit Sub
	Err_Control:
	Select Case Err.Number
	Case -2147352567
	If GetAsyncKeyState(VK_ESCAPE) And &H8000 > 0 Then
	Err.Clear
	Resume Exit_Here
	ElseIf GetAsyncKeyState(VK_LBUTTON) > 0 Then
	Err.Clear
	Resume
	End If
	' Case -2145320928
	' 'User input is keyword or..
	' 'Right click
	' Err.Clear
	' Resume Exit_Here
	Case Else
	MsgBox Err.Description
	Resume Exit_Here
	End Select
End Sub

'break sample
Sub Break()
	Dim Pnt As Variant
	Dim entObj As AcadEntity
	ActiveDocument.Utility.GetEntity entObj, Pnt, "Select Object£º"
	Dim Pnt2 As Variant
	Pnt2 = ActiveDocument.Utility.GetPoint(, "Select point£º")
	Dim det As String
	det = GetDoubleEntTable(entObj, Pnt)
	Dim lspPnt As String
	lspPnt = axPoint2lspPoint(Pnt2)
	ActiveDocument.SendCommand "_break" & vbCr & det & vbCr & lspPnt & vbCr
End Sub
' Trim Sample
Sub Trim()
	Dim Pnt1 As Variant
	Dim entObj1 As AcadEntity
	ActiveDocument.Utility.GetEntity entObj1, Pnt1, "Select Object£º"
	Dim det1 As String
	det1 = axEnt2lspEnt(entObj1)
	Dim Pnt2 As Variant
	Dim entObj2 As AcadEntity
	ActiveDocument.Utility.GetEntity entObj2, Pnt2, "Select trim Ojbect£º"
	Dim det2 As String
	det2 = GetDoubleEntTable(entObj2, Pnt2)
	ActiveDocument.SendCommand "_trim" & vbCr & det1 & vbCr & vbCr & det2 & vbCr & vbCr
End Sub
'convert Point and Object to LISP format
Public Function GetDoubleEntTable(entObj As AcadEntity, Pnt As Variant) As String
	Dim entHandle As String
	entHandle = entObj.Handle
	GetDoubleEntTable = "(list(handent " & Chr(34) & entHandle & Chr(34) & ")(list " & str(Pnt(0))
	& str(Pnt(1)) & str(Pnt(2)) & "))"
End Function
'convert Point to LISP format
Public Function axPoint2lspPoint(Pnt As Variant) As String
	axPoint2lspPoint = Pnt(0) & "," & Pnt(1) & "," & Pnt(2)
End Function
'Convert Object to lisp format
Public Function axEnt2lspEnt(entObj As AcadEntity) As String
	Dim entHandle As String
	entHandle = entObj.Handle
	axEnt2lspEnt = "(handent " & Chr(34) & entHandle & Chr(34) & ")"
End Function

