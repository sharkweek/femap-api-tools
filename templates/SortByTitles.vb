Function Private SortByTitle(entity_set, entity) As Variant
    ' sort entities by title

    ' Arguments
    ' ---------
    ' entity_set : femap.Set
    ' entity : femap object matching type of entities in `entity_set`

    ' Return
    ' ------
    ' array
    ' all entity IDs sorted by title

    Dim listID() As Variant
    Dim titles() As Variant
    Dim swapped As Boolean
    Dim tSwap As Variant
    Dim idSwap As Variant
    Dim iMax As Long

    ' Get a list of titles for specified group set
    entity_set.Reset()
    k = 0
    Do While entity.NextInSet(entity_set.ID)
        ReDim Preserve titles(k)
        ReDim Preserve listID(k)
        entity.Get(entity.ID)
        titles(k) = entity.title
        listID(k) = entity.ID
        k = k + 1
    Loop

    ' bubble sort by entity titles
    iMax = entity_set.Count() - 2
    Do
        swapped = False
        For i = 0 To iMax
            If titles(i) > titles(i + 1) Then
                ' swap titles
                tSwap = titles(i)
                titles(i) = titles(i + 1)
                titles(i + 1) = tSwap

                ' swap IDs
                idSwap = listID(i)
                listID(i) = listID(i + 1)
                listID(i + 1) = idSwap

                swapped = True
            End If
        Next i
        iMax = iMax - 1
    Loop Until swapped = False

    SortByTitle = listID

End Function