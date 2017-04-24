Sub WRITE_CUSTOM_OUTPUT_VECTOR()
‘ Creates a custom output vector

Dim app As femap.Model
Set app = GetObject(, “femap.model”)

Dim new_output_set As femap.OutputSet
Set new_output_set = app.feOutputSet

With new_output_set
    .Title = “New Output Set”
    .ID = new_output_set.NextEmptyID
    .program = FAP_UNKNOWN
    .analysis = FAT_UNKNOWN
    .Value = 0#
    .Notes = “Data manually added”
    .Put (new_output_set.ID)
End With

‘Output Vector 1
Dim new_vector_1 As femap.output
Set new_vector_1 = app.feOutput
rc = new_vector_1.InitScalarAtElem(new_output_set.ID, 9000000, “Custom OV 1”, FOT_ANY, True)

‘Output Vector 2
Dim new_vector_2 As femap.output
Set new_vector_2 = app.feOutput
rc = new_vector_2.InitScalarAtElem(new_output_set.ID, 9000001, “Custom OV 2”, FOT_ANY, True)

‘Loop condition
new_vector_1.Value(currELEMID) = ‘Value
new_vector_2.Value(currELEMID) = ‘Value

‘Loop over every element
rc = new_vector_1.Put(9000000)
rc = new_vector_2.Put(9000001)

rc = app.feappmessage(2, “Write operation complete…”)
rc = app.feFileRebuild(False, False)

End Sub
