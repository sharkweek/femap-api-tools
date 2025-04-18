' Title: Spider (Fully Constrained).BAS
' Author: Andy Perez
' License: OSL-3.0
' Femap API Version: 2306

' Simcenter Femap, Simcenter Nastran, and related documentation are proprietary
' to Siemens Digital Industries Software Inc. Siemens and the Siemens logo are
' registered trademarks of Siemens AG. NX is a trademark or registered trade­mark
' of Siemens Digital Industries Software Inc. or its subsidiaries in the United
' States and in other countries.

' Usage Notes
' -----------
' Modified from default Spider.BAS included in Femap. This script creates
' a spider that is constrained in all 6 DOF instead of just translational

Sub Main
	Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage(FCM_COMMAND, "Create Fully Constrained Spider")

	Dim feNode As femap.Node
	Dim nodeSet As femap.Set
	Dim feElem As femap.Elem
	Dim nodeCount As Long
	Dim nodeX As Double
	Dim nodeY As Double
	Dim nodeZ As Double
	Dim nodeID As Long
	Dim elemID As Long

	Dim vNodeArray As Variant

	nodeX = 0#
	nodeY = 0#
	nodeZ = 0#

	' Get a set of nodes for this rigid element

	Set nodeSet = App.feSet()

	rc = nodeSet.Select(FT_NODE, True, "Select Nodes for Rigid Spider")

	If rc = -1 Then 'return code FE_OK

    'Lets see how many nodes were selected
    nodeCount = nodeSet.count()

    If nodeCount > 0 Then

        ' Walk the nodes and find the average
        Set feNode = App.feNode()
        rc = nodeSet.Reset()
        nodeID = nodeSet.Next()

        ReDim nodeArray(nodeCount) As Long

        Dim passCount As Long

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
            Set feElem = App.feElem()

            elemID = feElem.NextEmptyID
			feElem.type = FET_L_RIGID
			feElem.topology = FTO_RIGIDLIST

            ' Set independent node
            feElem.Node(0) = nodeID

            ' Set constrained degrees of freedom
            feElem.release(0, 0) = True
            feElem.release(0, 1) = True
            feElem.release(0, 2) = True
            feElem.release(0, 3) = True
            feElem.release(0, 4) = True
            feElem.release(0, 5) = True

            feElem.ID = elemID

	        rc = feElem.PutNodeList(0, nodeCount, vNodeArray, Null, Null, Null)
            rc = feElem.Put(elemID)
        End If
    End If
End If

Call App.feViewRegenerate(0)

End Sub
