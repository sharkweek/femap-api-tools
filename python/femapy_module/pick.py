"""A set of functions for prompting entity selection by the user."""

# some useful shorcut functions
def pick_entities(entityType, clear=True, prompt="Select entities..."):
    """Return user-selected the IDs as an array."""

    pickSet = app.feSet
    pickSet.Select(entityType, clear, prompt)

    rc, _, setArray = pickSet.GetArray()  # get the array
    del _, pickSet  # delete intermediate variables from cache
    if rc == fc.FE_CANCEL:
        response = "User cancelled selection..."
        app.feAppMessage(fc.FCM_ERROR, response)
        print(response)
    elif rc == fc.FE_NOT_EXIST:
        response = "No entities of selected type exist."
        app.feAppMessage(fc.FCM_ERROR, response)
        print(response)

    return setArray


def pick_id(entityType, prompt="Select entities..."):
    """Prompt user to pick a single entity and returns the ID."""

    pickSet = app.feSet
    rc, enity_id = pickSet.SelectID(entityType, prompt)

    #  Ensure user selects a valide entity
    if rc == fc.FE_CANCEL:
        response = "User cancelled selection..."
        app.feAppMessage(fc.FCM_ERROR, response)
        print(response)
    elif rc == fc.FE_NOT_EXIST:
        response = "No entities of selected type exist."
        app.feAppMessage(fc.FCM_ERROR, response)
        print(response)

    return enity_id