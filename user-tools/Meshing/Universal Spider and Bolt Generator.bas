'API Written by Predictive Engineering 2011
'Predictive Engineering Assumes No Responsibility For Results Obtained From API
'Written for FEMAP 10.3
'This program automates the creation of two rigid links and one beam element
'connecting the rigid links together.  It allows for the creation of new properties and
'materials, or the specification of existing properties.

Sub Main
	Dim App As femap.model
    Set App = feFemap()


	Dim feNode As femap.Node
	Dim nodeSet As femap.Set
	Dim feElem As femap.Elem
	Dim nodeCount As Long
	Dim nodeX As Double
	Dim nodeY As Double
	Dim nodeZ As Double
	Dim nodeX2 As Double
	Dim nodeY2 As Double
	Dim nodeZ2 As Double
	Dim nodeID As Long
	Dim elemID As Long

	Dim vNodeArray As Variant
	Dim passCount As Long
	Dim eString(2) As String
	eString(0) = "Select Nodes for the First Rigid Spider"
	eString(1) = "Select Nodes for the Second Rigid Spider"

	Dim eSet As Object
	Set eSet = App.feSet

	Dim pSet As Object
	Set pSet = App.feSet

	Dim MT As Object
	Set MT = App.feMatl

	Set feElem = App.feElem




	' Get a set of nodes for this rigid element

	Set nodeSet = App.feSet()

	For p = 1 To 2

	nodeX = 0#
	nodeY = 0#
	nodeZ = 0#

	rc = nodeSet.Clear

	App.feAppMessage(0,eString(p-1))

	rc = nodeSet.Select(FT_NODE, True, eString(p-1))

	If rc = -1 Then 'return code FE_OK

    'Lets see how many nodes were selected
    nodeCount = nodeSet.count()

    If nodeCount > 0 Then

        ' Walk the nodes and find the average
        Set feNode = App.feNode()
        rc = nodeSet.Reset()
        nodeID = nodeSet.Next()

        ReDim nodeArray(nodeCount) As Long



        passCount = 0

        Do While nodeID <> 0

            nodeArray(passCount) = nodeID
            passCount = passCount + 1
            rc = feNode.Get(nodeID)

            nodeX = nodeX + feNode.x
            nodeY = nodeY + feNode.y
            nodeZ = nodeZ + feNode.z

            nodeID = nodeSet.Next()

        Loop

        vNodeArray = nodeArray

        nodeID = feNode.NextEmptyID

        feNode.ID = nodeID
        feNode.x = nodeX / nodeCount
        feNode.y = nodeY / nodeCount
        feNode.z = nodeZ / nodeCount

        rc = feNode.Put(nodeID)

        If rc = -1 Then 'return code FE_OK
            ' create the element


            elemID = feElem.NextEmptyID
			feElem.type = FET_L_RIGID
			feElem.topology = FTO_RIGIDLIST

            feElem.Node(0) = nodeID 'Independent Node
            feElem.release(0, 0) = 1
            feElem.release(0, 1) = 1
            feElem.release(0, 2) = 1
            feElem.release(0, 3) = 1
            feElem.release(0, 4) = 1
            feElem.release(0, 5) = 1

            feElem.ID = elemID

	        rc = feElem.PutNodeList(0, nodeCount, vNodeArray, Null, Null, Null)
            rc = feElem.Put(elemID)
            rc = eSet.Add(nodeID)
            App.feAppMessage(0,"Rigid Link Created")
        End If
    End If
Else
GoTo FAIL
End If

Next p

Dim pr As Object
Set pr = App.feProp

Dim pCount As Long
Dim mCount As Long
Dim pIDs()
Dim mIDs()


rc = pSet.AddAll(FT_PROP)
pCount = pSet.count+2

    ReDim lists(pCount+1) As String
    ReDim pIDs(pCount-2)

    lists(0) = "Select a Property Option:"
    lists(pCount-1) = "Create a New Property with Specified Diameter"
    lists(pCount) = "Create a Custom Property"

For i = 1 To pCount-2

	rc = pr.Get(pSet.Next)
	pIDs(i-1) = pr.ID
    lists(i) = Str$(pr.ID) + ".. " + pr.title
Next i

rc = pSet.Clear

rc = pSet.AddAll(FT_MATL)
mCount = pSet.count+2

    ReDim lists2(mCount) As String
    ReDim mIDs(mCount-2)

    lists2(0) = "Select a Material Option:"
    lists2(mCount-1) = "Create a New Material"

For i = 1 To mCount-2

	rc = MT.Get(pSet.Next)
	mIDs(i-1) = MT.ID
    lists2(i) = Str$(MT.ID) + ".. " + MT.title
Next i

App.feAppMessage(2,"Please Specify Materials and Properties")
App.feAppMessage(2,"If an existing property is selected, no other information is required.")
App.feAppMessage(2,"If 'Create a New Property' is selected, bolt diameter and material also must be entered.")

BOXTOP:


	Begin Dialog UserDialog 360,280,"Specify Material and Property" ' %GRID:10,7,1,1
		GroupBox 10,7,340,49,"Property Specification",.GroupBox2
		DropListBox 20,28,320,63,lists(),.list1
		GroupBox 10,63,340,49,"Material Specification",.GroupBox1
		DropListBox 20,84,320,63,lists2(),.list2
		TextBox 180,126,90,21,.boltDia
		text 50,126,100,14,"Bolt Diameter =",.Text1,1
		OKButton 60,168,90,21
		CancelButton 190,168,90,21
		text 60,203,220,63,"If an existing property is selected, no other information is required.     If 'Create a New Property' is selected, bolt diameter and material also must be entered.",.Text2
	End Dialog
	Dim dlg As UserDialog

	If Dialog(dlg) = 0 Then
	GoTo FAIL
	End If

Dim Mn, Pn As Long
Dim boltDia As Double

boltDia = Val(dlg.boltDia)

Pn = dlg.list1

App.feAppMessage(0,Str$(Pn))

Mn = dlg.list2

Dim checkSet As Object
Set checkSet = App.feSet

Dim checkCount As Long
Dim checkID As Long

rc = checkSet.AddAll(FT_PROP)
checkCount = checkSet.count

Dim MTID As Long
Dim PID As Long

If Pn = 0 Then
	rc = App.feAppMessageBox(1,"Properties Not Specified, Specify or Cancel")
	If rc = -1 Then
		GoTo BOXTOP
	Else
		GoTo FAIL
	End If
End If

If Pn = pCount-1 Then

	If boltDia = 0 Then
		rc = App.feAppMessageBox(1,"A value of zero was entered for the bolt diameter, Re-enter or Cancel")
		If rc = -1 Then
			GoTo BOXTOP
		Else
			GoTo FAIL
		End If
	End If

	If Mn = 0 Then
		rc = App.feAppMessageBox(1,"Materials Not Specified, Specify or Cancel")
		If rc = -1 Then
			GoTo BOXTOP
		Else
			GoTo FAIL
		End If
	End If

	If Mn = mCount - 1 Then
		MTID = MT.NextEmptyID
		rc = App.feRunCommand(1222,True)
	Else
		MTID = mIDs(Mn - 1)
	End If

pr.matlID = MTID
pr.type = 5     ' Beam
pr.flagI(1) = 5 ' Circular Bar
pr.pval(40) = boltDia/2 ' Radius
pr.pval(46) = 4
pr.pval(47) = 1
pr.pval(48) = 0
pr.pval(49) = 0
PID = pr.NextEmptyID
pr.ID=PID
rc = pr.ComputeShape(False, False, False)
pr.title = Str$(boltDia) + " Bolt Property"
pr.Put (PID)

App.feAppMessage(2,"Property #" + Str$(PID) + " has been created.")

ElseIf Pn  < pCount-1 Then
	PID = pIDs(Pn - 1)
End If

If Pn = pCount Then
	MTID = mIDs(Mn - 1)
	MT.Active = MTID
	rc = App.feRunCommand(1223,True)
	rc = pr.Reset

	For i = 1 To checkCount + 1
		rc = pr.Next
		If checkSet.IsAdded(pr.ID) = 0 Then
			PID = pr.ID
		End If
	Next i

End If

pr.Get(PID)
If pr.type <> 5 Then
	rc = App.feAppMessageBox(1,"Selected Property Type is not a Beam, Respecify or Cancel")
	If rc = -1 Then
		GoTo BOXTOP
	Else
		GoTo FAIL
	End If
End If

feElem.Reset
eSet.Reset

Dim vecIn(3) As Double

Dim nElem As Long
nElem = feElem.NextEmptyID

feElem.ClearNodeList(-1)

feElem.release(0, 0) = 0
feElem.release(0, 1) = 0
feElem.release(0, 2) = 0
feElem.release(0, 3) = 0
feElem.release(0, 4) = 0
feElem.release(0, 5) = 0
feElem.type = 5
feElem.topology = 0
feElem.propID = PID
feElem.orientID = 0
feElem.Node(0) = eSet.Next
rc = feNode.Get(feElem.Node(0))

nodeX = feNode.x
nodeY = feNode.y
nodeZ = feNode.z

feElem.Node(1) = eSet.Next
rc = feNode.Get(feElem.Node(1))

nodeX2 = feNode.x
nodeY2 = feNode.y
nodeZ2 = feNode.z

vecIn(0) = nodeX - nodeX2
vecIn(1) = nodeY - nodeY2
vecIn(2) = nodeZ - nodeZ2

App.feVectorPerpendicular(vecIn, arbVec)

feElem.orient(0) = arbVec(0)
feElem.orient(1) = arbVec(1)
feElem.orient(2) = arbVec(2)
feElem.ID = nElem
feElem.Put(nElem)

App.feAppMessage(0,"Element #" + Str$(nElem) + " has been created.")
App.feAppMessage(0,"Program Completed")

Call App.feViewRegenerate(0)

FAIL:

End Sub
