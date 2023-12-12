' File: WriteReadSets.BAS
' Version: Version 1
' Author: Ali Gokce, ATA Engineering
' www.ata-plmsoftware.com
' Description: This API reads Nastran set definitions into FEMAP as groups,
' or writes out Nastran sets
'  from FEMAP groups in the message window.

'Usage Notes
'====================================
' In the API user interface (UI), select if you want to process elements or
' nodes. To read a set: Copy-paste set definition lines from a Nastran file
' into the text area of the API UI. The set will be saved in a FEMAP group;
' group ID will be posted in the message window. For very large set
' definitions, more than one copy-paste might be needed. To write a set: Select
' the elements or nodes that needs to be written out. One could copy-paste this
' message window content to the DAT files. Also it is possible to have FEMAP
' send message window output to a file under "List > Destination". You could
' either select a set of groups and\or nodes one at a time, or you could use
' multiple groups to write out sets.

' Copyright (c) <2008> <ATA Engineering, Inc.>
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software") to use,
' copy, modify, merge, and/or distribute royalty-free copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to
' the following conditions:
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Sub Main
        'Define a string indicating the version of the translator
        Const API_version As String = "1"

	Dim App As femap.model
	Set App = GetObject(,"femap.model")

        'Indicate which version and contact information for support
        App.feAppMessage(FCM_NORMAL,"*** Write Out Sets, Version " & API_version & " ***")
        App.feAppMessage(FCM_NORMAL,"Created by ATA Engineering, http://www.ata-plmsoftware.com")

	Begin Dialog UserDialog 790,225,"API to read & write Nastran sets" ' %GRID:5,3,0,1
		GroupBox 150,0,640,70,"Write parameters",.GroupBox2
		CheckBox 160,15,110,15,"Write nodes",.CheckBoxN
		TextBox 660,14,60,21,.setID
		text 160,49,540,14,"* Nodes are written to sets with odd IDs, elements are written to sets with even IDs. ",.Text1
		text 160,35,290,14,"* All sets are written to the message window.",.Text2
		text 570,14,90,14,"Initial set ID:",.Text3
		OKButton 10,56,60,21
		CancelButton 80,56,60,21
		CheckBox 410,14,140,14,"Write from groups",.CheckBoxG
		GroupBox 0,78,790,141,"Read parameters",.GroupBox3
		TextBox 15,117,760,96,.SetText,1
		CheckBox 280,14,120,14,"Write elements",.CheckBoxE
		GroupBox 0,0,130,51,"Select Action",.GroupBox1
		OptionGroup .Group1
			OptionButton 10,33,90,15,"Read Set",.OptionRead
			OptionButton 10,18,90,15,"Write Set",.OptionWrite
		CheckBox 35,96,105,15,"Read nodes",.ChkReadN
		CheckBox 150,96,135,15,"Read elements",.ChkReadE
		text 325,87,440,15,"* Copy Nastran SET cards and paste below without modification.",.Text4
		text 325,102,405,15,"* Handles THRU but not EXCEPT or INCLUDE",.Text5
	End Dialog
	Dim dlg As UserDialog
    dlg.setID = "1"
    dlg.CheckBoxE = 1
    dlg.CheckBoxN = 1
    dlg.CheckBoxG = 1
    dlg.ChkReadN = 1
    dlg.ChkReadE = 1
    'dlg.OptionRead = False
	Dialog dlg

    'App.feAppMessage(FCM_ERROR, ">>>  " & dlg.Group1)

    Set_ID = Val(dlg.setID)

    Dim OutStr As String

    Dim Groupset As femap.Set
    Dim Elemset As femap.Set
    Dim Nodeset As femap.Set
    Dim Entityset As femap.Set
  	Dim groupID As Long
    Dim feGroup As femap.Group

    Set Groupset = App.feSet
    Set Elemset = App.feSet
    Set Nodeset = App.feSet
    Set Entityset = App.feSet
  	Set feGroup = App.feGroup()

    If dlg.Group1 = 1 Then  '======================= "WRITE SET" BELOW ========
		If dlg.CheckBoxE=0 Then
        	If dlg.CheckBoxN=0 Then
				App.feAppMessage(FCM_WARNING, "Neither elements nor nodes were checked to write!")
            	Exit All
            End If
        End If
	    If dlg.CheckBoxG Then
            rc= Groupset.Select(FT_GROUP,True,"Select groups to write out to sets")
       		pointer1 = Groupset.First()
       		While pointer1
            	rc = Elemset.Clear()
				rc = Nodeset.Clear()
       			If dlg.CheckBoxE Then
           	   		rc = Elemset.AddGroup(FT_ELEM, pointer1)
               		If Elemset.count > 0 Then
                    	App.feAppMessage(FCM_HIGHLIGHT, "$$$$$$ Elements from Group " & pointer1 & " below. $$$$$$")
                  		WriteOutSets(App, Elemset, Set_ID, 1)
               		End If
            	End If
            	If dlg.CheckBoxN Then
               		rc = Nodeset.AddGroup(FT_NODE, pointer1)
               		If Nodeset.count > 0 Then
                    	App.feAppMessage(FCM_HIGHLIGHT, "$$$$$$ Nodes from Group " & pointer1 & " below. $$$$$$")
                  		WriteOutSets(App, Nodeset, Set_ID, 0)
               		End If
            	End If
		    	pointer1=Groupset.Next()
        	Wend
        Else
        	If dlg.CheckBoxE Then
            	rc = Elemset.Select(FT_ELEM, True, "Select elements to write to the set")
                If Elemset.count > 0 Then
                   WriteOutSets(App, Elemset, Set_ID, 1)
                End If
            End If
            If dlg.CheckBoxN Then
            	rc = Nodeset.Select(FT_NODE, True, "Select elements to write to the set")
                If Nodeset.count > 0 Then
                   WriteOutSets(App, Nodeset, Set_ID, 0)
                End If
            End If
    	End If
    Else '=================================== "READ SET" BELOW ====================
    	If dlg.ChkReadN = 0 Then
        	 If dlg.ChkReadE = 0 Then
				App.feAppMessage(FCM_WARNING, "Neither elements nor nodes were checked to read!" & dlg.ChkReadN & " -   " & dlg.ChkReadE)
            	Exit All
            End If
        End If
		SetStr = dlg.SetText
        If Len(SetStr) < 3 Then
        	App.feAppMessage(FCM_WARNING, "Insufficient data pasted in the text window!")
            Exit All
        End If
        SetStr = Replace(UCase$(SetStr), vbNewLine,"")',,,vbTextCompare))
        ind = 1
        Dim SetStrCurr As String
        Dim StrArray As Variant
        Strlen = Len(SetStr)
        While ind < StrLen
        	rc = Entityset.Clear
    		groupID = feGroup.NextEmptyID
    		feGroup.ID = groupID
            typestr = ""
			'
        	Set1Ind = InStr(ind, SetStr, "SET")
			EqualInd = InStr(ind, SetStr, "=")
            SetID = Val(Trim(Mid(SetStr, Set1Ind+3, EqualInd-Set1Ind-3)))
            Set2Ind =InStr(EqualInd+1, SetStr, "SET")
            If IsNull(Set2Ind) Or Set2Ind=0 Then
            	Set2Ind = Strlen+1
            End If
            SetStrCurr = Mid$(SetStr, EqualInd+2, Set2Ind-EqualInd-2)
            StrArray = Split(SetStrCurr, ",")

            For i = 0 To UBound(StrArray)
          		CurrText = Trim(StrArray(i))
				ThruInd = InStr(1, CurrText, "THRU")
       	        If IsNull(ThruInd) Or ThruInd=0 Then
   	            	StartInd = Val(CurrText)
                    If StartInd > 0 Then
                        rc = Entityset.Add(StartInd)
                   	End If
				Else
               		StartInd = Val(Trim(Split(CurrText, "THRU")(0)))
           	        EndInd =Val(Trim(Split(CurrText,"THRU")(1)))
                    If StartInd > 0 And EndInd > 0 Then
                       	rc = Entityset.AddRange(StartInd, EndInd, 1)
                     End If
   	            End If
					'App.feAppMessage(FCM_HIGHLIGHT, Trim(StrArray(i)))
            Next
            If dlg.ChkReadN Then
      			rc = feGroup.SetAdd(FT_NODE,Entityset.ID)
                typestr = " nodes"
            End If
            If dlg.ChkReadE Then
            	rc = feGroup.SetAdd(FT_ELEM,Entityset.ID)
                If Len(typestr) > 2 Then
                	typestr = typestr & " and elements"
                Else
                    typestr = " elements"
                End If
            End If
    		rc = feGroup.Put(groupID)
			App.feAppMessage(FCM_NORMAL, "Entities in SET " & SetID & " are written to group " & groupID & " as" & typestr & "!")
			ind = Set2Ind
        Wend
    End If
End Sub
Sub WriteOutSets(App, Set1, ByRef Set_ID, Flag1)

    Dim  OutStr As String
    Dim AddStr As String

    If Fix(Set_ID/2) = Set_ID/2 And Not Flag1 Then
            Set_ID = Set_ID+1
     End If
     If Fix(Set_ID/2) <> Set_ID/2 And Flag1 Then
            Set_ID = Set_ID+1
     End If

    pointer2 = Set1.First()
    BeginID = pointer2
    OutStr = "SET " & Set_ID & "= "
    AddStr = "" & BeginID
    While Set1.NextID
      	If Set1.NextID > pointer2 + 1 Then
        	If pointer2 = BeginID+1 Then
        	   AddStr = AddStr & ", " & pointer2
            ElseIf pointer2 > BeginID Then
        	   AddStr = AddStr & " THRU " & pointer2
        	'Else
        	   'AddStr = "" & BeginID
            End If
            BeginID=Set1.NextID
            If Len(OutStr) + Len(AddStr) < 55 Then
            	OutStr = OutStr & AddStr & ", "
            Else
            	App.feAppMessage(1, OutStr)
                OutStr = "    " & AddStr & ", "
            End If
            AddStr = "" & BeginID
        End If
        pointer2 = Set1.Next
      Wend

      	If pointer2 = BeginID+1 Then
       	   AddStr = AddStr & ", " & pointer2
        ElseIf pointer2 > BeginID Then
       	   AddStr = AddStr & " THRU " & pointer2
       	'Else
       	   'AddStr = "" & BeginID
        End If
        BeginID=Set1.NextID
        If Len(OutStr) + Len(AddStr) < 55 Then
          	OutStr = OutStr & AddStr
        Else
           	App.feAppMessage(1, OutStr)
            OutStr = "    " & AddStr
        End If

      App.feAppMessage(1, OutStr)
End Sub
' CODE I USED FOR DEBUGGING
'            App.feAppMessage(FCM_NORMAL, SetStr)
'            App.feAppMessage(FCM_NORMAL, "" & EqualInd & "   " & Set2Ind)
'            App.feAppMessage(FCM_HIGHLIGHT,"Set curr= " & SetStrCurr)
'            App.feAppMessage(FCM_HIGHLIGHT, "bound=" & UBound( StrArray))
'            App.feAppMessage( FCM_HIGHLIGHT, " waaa " & Split(SetStrCurr, ",")(2) )
'            App.feAppMessage( FCM_HIGHLIGHT, "array=" & StrArray(0) )
'App.feAppMessage(FCM_NORMAL, "" & Set1Ind & "   " & EqualInd & "    " &  SetID) 'SetStr)
'                    	App.feAppMessage(FCM_HIGHLIGHT, "range = " & StartInd & " +++++>  " & EndInd)
'                           App.feAppMessage(FCM_NORMAL, ">>>" & EqualInd & "   " & Set1Ind)
'                App.feAppMessage(FCM_NORMAL, ">>>" & EqualInd & "   " & Set2Ind)
'                        						App.feAppMessage(FCM_HIGHLIGHT, "single = " & CurrText)
'     	App.feAppMessage(FCM_WARNING, "Ne " & dlg.ChkReadN & " -   " & dlg.ChkReadE)
