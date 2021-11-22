Attribute VB_Name = "starWarsUtility"
Sub getAPIData(strUrl As String)

    Dim objRequest As Object
    Dim blnAsync As Boolean

    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    blnAsync = True

    With objRequest
        .Open "GET", strUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        'spin wheels whilst waiting for response
        While objRequest.readyState <> 4
            DoEvents
        Wend
        apiResponse = .ResponseText
    End With

    'Debug.Print apiResponse

End Sub

Sub planetsHeader(Optional offsetRow = 0)
    cellRow = ActiveCell.Row + offsetRow
    cellColumn = ActiveCell.Column

    Cells(cellRow + 1, cellColumn).Value = "Planets"
    Cells(cellRow + 1, cellColumn + 1).Value = "Diameter"
    Cells(cellRow + 1, cellColumn + 2).Value = "Climate"
    Cells(cellRow + 1, cellColumn + 3).Value = "Gravity"
    Cells(cellRow + 1, cellColumn + 4).Value = "Terrain"
    Cells(cellRow + 1, cellColumn + 5).Value = "Surface Water"
    Cells(cellRow + 1, cellColumn + 6).Value = "Population"
    Cells(cellRow + 1, cellColumn + 7).Value = "Known Residents"
End Sub

Sub peopleHeader(Optional offsetRow = 0)
    cellRow = ActiveCell.Row + offsetRow
    cellColumn = ActiveCell.Column

    Cells(cellRow + 1, cellColumn).Value = "Name"
    Cells(cellRow + 1, cellColumn + 1).Value = "Height"
    Cells(cellRow + 1, cellColumn + 2).Value = "Weight"
    Cells(cellRow + 1, cellColumn + 3).Value = "Hair Color"
    Cells(cellRow + 1, cellColumn + 4).Value = "Skin Color"
    Cells(cellRow + 1, cellColumn + 5).Value = "Eye Color"
    Cells(cellRow + 1, cellColumn + 6).Value = "Birth year"
    Cells(cellRow + 1, cellColumn + 7).Value = "Gender"
    Cells(cellRow + 1, cellColumn + 8).Value = "Homeworld"
    Cells(cellRow + 1, cellColumn + 9).Value = "Vehicles"
    Cells(cellRow + 1, cellColumn + 10).Value = "Starships"
    Cells(cellRow + 1, cellColumn + 11).Value = "Film Appearances"
    
End Sub

Sub starshipsHeader(Optional offsetRow = 0)
    cellRow = ActiveCell.Row + offsetRow
    cellColumn = ActiveCell.Column

    Cells(cellRow + 1, cellColumn).Value = "Name"
    Cells(cellRow + 1, cellColumn + 1).Value = "Manufacturer"
    Cells(cellRow + 1, cellColumn + 2).Value = "Cost"
    Cells(cellRow + 1, cellColumn + 3).Value = "Length"
    Cells(cellRow + 1, cellColumn + 4).Value = "Max Speed"
    Cells(cellRow + 1, cellColumn + 5).Value = "Crew"
    Cells(cellRow + 1, cellColumn + 6).Value = "Passengers"
    Cells(cellRow + 1, cellColumn + 7).Value = "Cargo Capacity"
    Cells(cellRow + 1, cellColumn + 8).Value = "Hyperdrive Rating"
    Cells(cellRow + 1, cellColumn + 9).Value = "Starship class"
    Cells(cellRow + 1, cellColumn + 10).Value = "Pilots"
    Cells(cellRow + 1, cellColumn + 11).Value = "Film Appearances"
    
End Sub
Sub vehiclesHeader(Optional offsetRow = 0)
    cellRow = ActiveCell.Row + offsetRow
    cellColumn = ActiveCell.Column

    Cells(cellRow + 1, cellColumn).Value = "Name"
    Cells(cellRow + 1, cellColumn + 1).Value = "Manufacturer"
    Cells(cellRow + 1, cellColumn + 2).Value = "Cost"
    Cells(cellRow + 1, cellColumn + 3).Value = "Length"
    Cells(cellRow + 1, cellColumn + 4).Value = "Max Speed"
    Cells(cellRow + 1, cellColumn + 5).Value = "Crew"
    Cells(cellRow + 1, cellColumn + 6).Value = "Passengers"
    Cells(cellRow + 1, cellColumn + 7).Value = "Cargo Capacity"
    Cells(cellRow + 1, cellColumn + 8).Value = "Vehicle class"
    Cells(cellRow + 1, cellColumn + 9).Value = "Pilots"
    Cells(cellRow + 1, cellColumn + 10).Value = "Film Appearances"
    
End Sub

Sub speciesHeader(Optional offsetRow = 0)
    cellRow = ActiveCell.Row + offsetRow
    cellColumn = ActiveCell.Column

    Cells(cellRow + 1, cellColumn).Value = "Name"
    Cells(cellRow + 1, cellColumn + 1).Value = "Classification"
    Cells(cellRow + 1, cellColumn + 2).Value = "Designation"
    Cells(cellRow + 1, cellColumn + 3).Value = "Average Height"
    Cells(cellRow + 1, cellColumn + 4).Value = "Skin Colors"
    Cells(cellRow + 1, cellColumn + 5).Value = "Hair Color"
    Cells(cellRow + 1, cellColumn + 6).Value = "Average Lifespan"
    Cells(cellRow + 1, cellColumn + 7).Value = "Homeworld"
    Cells(cellRow + 1, cellColumn + 8).Value = "Language"
    Cells(cellRow + 1, cellColumn + 9).Value = "People"
    
End Sub
