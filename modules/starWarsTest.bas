Attribute VB_Name = "starWarsTest"
Global apiResponse As String
Sub say_hello(control As IRibbonControl)
    MsgBox "Hello " & Application.UserName & "!" & vbNewLine & "I hope you would like my Star Wars Add-In."
End Sub
Sub getListOfPlanets(control As IRibbonControl)
    'Example of sub with slow performance, sending multiple requests for grabbing data for each Planet.
    Dim Json As Object
    Dim countPlanets As Integer
    Dim residentsURL As String
    
    apiResponse = ""
    cellRow = ActiveCell.Row
    cellColumn = ActiveCell.Column
    
    'Adding the header
    Call planetsHeader(-1)
    
    Call getAPIData("https://swapi.dev/api/planets/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    countPlanets = Json("count")
    
    'Request API calls for each planet
    'Grab the info needed
    For i = 1 To countPlanets
        Call getAPIData("https://swapi.dev/api/planets/" & i & "/")
        Set Planet = JsonConverter.ParseJson(apiResponse)
    
        Cells(cellRow + i, cellColumn).Value = Planet("name")
        Cells(cellRow + i, cellColumn + 1).Value = Planet("diameter")
        Cells(cellRow + i, cellColumn + 2).Value = Planet("climate")
        Cells(cellRow + i, cellColumn + 3).Value = Planet("gravity")
        Cells(cellRow + i, cellColumn + 4).Value = Planet("terrain")
        Cells(cellRow + i, cellColumn + 5).Value = Planet("surface_water")
        Cells(cellRow + i, cellColumn + 6).Value = Planet("population")
        
        countResidents = 0
        For Each Value In Planet("residents")
            countResidents = countResidents + 1
        Next Value
        Cells(cellRow + i, cellColumn + 7).Value = countResidents
        
    Next i
    
End Sub
Sub getListOfPlanets_betterPerformance(control As IRibbonControl)
    'Example of sub with better performance, sending few requests for grabbing data for each Planet.
    'Using JSON results pages
    Dim Json As Object
    Dim countPlanets As Integer
    Dim firstTime As Boolean
    
    apiResponse = ""
    firstTime = True
    cellRow = ActiveCell.Row
    cellColumn = ActiveCell.Column
    i = 1 'Used for incrementing rows
    
    'Adding the header
    Call planetsHeader(-1)
    
    Call getAPIData("https://swapi.dev/api/planets/")
    Set Json = JsonConverter.ParseJson(apiResponse)

    'Grab the info needed
continueNext:
    'Handles the first page from JSON results
    If firstTime = False Then
        Call getAPIData(Json("next"))
        Set Json = JsonConverter.ParseJson(apiResponse)
    End If
    
    For Each result In Json("results")
        Cells(cellRow + i, cellColumn).Value = result("name")
        Cells(cellRow + i, cellColumn + 1).Value = result("diameter")
        Cells(cellRow + i, cellColumn + 2).Value = result("climate")
        Cells(cellRow + i, cellColumn + 3).Value = result("gravity")
        Cells(cellRow + i, cellColumn + 4).Value = result("terrain")
        Cells(cellRow + i, cellColumn + 5).Value = result("surface_water")
        Cells(cellRow + i, cellColumn + 6).Value = result("population")
        
        'Count the number of known residents
        x = 0
        For Each Value In result("residents")
            x = x + 1
        Next Value
        Cells(cellRow + i, cellColumn + 7).Value = x
        
        'Increment to continue to the next row
        i = i + 1
        firstTime = False
        
        
    Next result
    
    If IsNull(Json("next")) = False Then GoTo continueNext
    
End Sub

Sub getListOfPeople(control As IRibbonControl)
    'Example of function with a better performance, sending few requests to get data for People.
    'This function makes API calls when the input is a API link(planetURL, vehicleURL,starshipsURL) and grab the relevant names.
    Dim Json As Object
    Dim firstTime As Boolean
    Dim planetURL As String
    Dim vehicleURL As String
    Dim starshipsURL As String
    
    
    apiResponse = ""
    firstTime = True
    cellRow = ActiveCell.Row
    cellColumn = ActiveCell.Column
    i = 1 'Used for incrementing rows
    
    'Adding the header
    Call peopleHeader(-1)
    
    Call getAPIData("https://swapi.dev/api/people/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    
    'Grab the info needed
continueNext:
    'Handles the first page from JSON results
    If firstTime = False Then
        Call getAPIData(Json("next"))
        Set Json = JsonConverter.ParseJson(apiResponse)
    End If
    
    For Each result In Json("results")
        Cells(cellRow + i, cellColumn).Value = result("name")
        Cells(cellRow + i, cellColumn + 1).Value = result("height")
        Cells(cellRow + i, cellColumn + 2).Value = result("mass")
        Cells(cellRow + i, cellColumn + 3).Value = result("hair_color")
        Cells(cellRow + i, cellColumn + 4).Value = result("skin_color")
        Cells(cellRow + i, cellColumn + 5).Value = result("eye_color")
        Cells(cellRow + i, cellColumn + 6).Value = result("birth_year")
        Cells(cellRow + i, cellColumn + 7).Value = result("gender")
        
        'Grab Homeworld URL and return its name
        planetURL = result("homeworld")
        Call getAPIData(planetURL)
        Set Planets = JsonConverter.ParseJson(apiResponse)
        Cells(cellRow + i, cellColumn + 8).Value = Planets("name")
        
        'Grab Vehicles URL and return its name
        vehiclesNames = ""
        firstVehicle = True
        For Each vehicle In result("vehicles")
            vehicleURL = vehicle
            Call getAPIData(vehicleURL)
            Set Vehicles = JsonConverter.ParseJson(apiResponse)
            If firstVehicle = True Then
                vehiclesNames = Vehicles("name")
            Else
                vehiclesNames = vehiclesNames & vbCrLf & Vehicles("name")
            End If
            firstVehicle = False
        Next vehicle
        Cells(cellRow + i, cellColumn + 9).Value = vehiclesNames
        
        'Grab Starships URL and return its name
        starshipsNames = ""
        firstStarship = True
        For Each starship In result("starships")
            starshipsURL = starship
            Call getAPIData(starshipsURL)
            Set Starships = JsonConverter.ParseJson(apiResponse)
            If firstStarship = True Then
                starshipsNames = Starships("name")
            Else
                starshipsNames = starshipsNames & "; " & Starships("name")
            End If
            firstStarship = False
        Next starship
        Cells(cellRow + i, cellColumn + 10).Value = starshipsNames
        
        'Count the number of film appearances
        x = 0
        For Each Value In result("films")
            x = x + 1
        Next Value
        Cells(cellRow + i, cellColumn + 11).Value = x
        
        'Increment to continue to the next row
        i = i + 1
        firstTime = False
        
        
    Next result
    
    If IsNull(Json("next")) = False Then GoTo continueNext
    
End Sub

Sub getListOfSpaceships(control As IRibbonControl)
    'Example of function with a better performance, sending few requests to get data for Spaceships/Starships.
    
    Dim Json As Object
    Dim firstTime As Boolean
    
    apiResponse = ""
    firstTime = True
    cellRow = ActiveCell.Row
    cellColumn = ActiveCell.Column
    i = 1 'Used for incrementing rows
    
    'Adding the header
    Call starshipsHeader(-1)
    
    Call getAPIData("https://swapi.dev/api/starships/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    
    'Grab the info needed
continueNext:
    'Handles the first page from JSON results
    If firstTime = False Then
        Call getAPIData(Json("next"))
        Set Json = JsonConverter.ParseJson(apiResponse)
    End If
    
    For Each result In Json("results")
        Cells(cellRow + i, cellColumn).Value = result("name")
        Cells(cellRow + i, cellColumn + 1).Value = result("manufacturer")
        Cells(cellRow + i, cellColumn + 2).Value = result("cost_in_credits")
        Cells(cellRow + i, cellColumn + 3).Value = result("length")
        Cells(cellRow + i, cellColumn + 4).Value = result("max_atmosphering_speed")
        Cells(cellRow + i, cellColumn + 5).Value = result("crew")
        Cells(cellRow + i, cellColumn + 6).Value = result("passengers")
        Cells(cellRow + i, cellColumn + 7).Value = result("cargo_capacity")
        Cells(cellRow + i, cellColumn + 8).Value = result("hyperdrive_rating")
        Cells(cellRow + i, cellColumn + 9).Value = result("starship_class")
        
        'Count the number of pilots
        x = 0
        For Each Value In result("pilots")
            x = x + 1
        Next Value
        Cells(cellRow + i, cellColumn + 10).Value = x
        
        'Count the number of film appearances
        x = 0
        For Each Value In result("films")
            x = x + 1
        Next Value
        Cells(cellRow + i, cellColumn + 11).Value = x
        
        'Increment to continue to the next row
        i = i + 1
        firstTime = False
        
    Next result
    
    If IsNull(Json("next")) = False Then GoTo continueNext
    
End Sub

Sub getListOfVehicles(control As IRibbonControl)
    'Example of function with a better performance, sending few requests to get data for Vehicles.
    Dim Json As Object
    Dim firstTime As Boolean
    
    apiResponse = ""
    firstTime = True
    cellRow = ActiveCell.Row
    cellColumn = ActiveCell.Column
    i = 1 'Used for incrementing rows
    
    'Adding the header
    Call vehiclesHeader(-1)
    
    Call getAPIData("https://swapi.dev/api/vehicles/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    
    'Grab the info needed
continueNext:
    'Handles the first page from JSON results
    If firstTime = False Then
        Call getAPIData(Json("next"))
        Set Json = JsonConverter.ParseJson(apiResponse)
    End If
    
    For Each result In Json("results")
        Cells(cellRow + i, cellColumn).Value = result("name")
        Cells(cellRow + i, cellColumn + 1).Value = result("manufacturer")
        Cells(cellRow + i, cellColumn + 2).Value = result("cost_in_credits")
        Cells(cellRow + i, cellColumn + 3).Value = result("length")
        Cells(cellRow + i, cellColumn + 4).Value = result("max_atmosphering_speed")
        Cells(cellRow + i, cellColumn + 5).Value = result("crew")
        Cells(cellRow + i, cellColumn + 6).Value = result("passengers")
        Cells(cellRow + i, cellColumn + 7).Value = result("cargo_capacity")
        Cells(cellRow + i, cellColumn + 8).Value = result("vehicle_class")
        
        'Count number of pilots
        x = 0
        For Each Value In result("pilots")
            x = x + 1
        Next Value
        Cells(cellRow + i, cellColumn + 9).Value = x
        
        'Count number of film appearances
        x = 0
        For Each Value In result("films")
            x = x + 1
        Next Value
        Cells(cellRow + i, cellColumn + 10).Value = x
        
        'Increment to continue to the next row
        i = i + 1
        firstTime = False
        
        
    Next result
    
    If IsNull(Json("next")) = False Then GoTo continueNext
    
End Sub

Sub getListOfSpecies(control As IRibbonControl)
    'Example of function with a better performance, sending few requests to get data for Species.
    Dim Json As Object
    Dim firstTime As Boolean
    
    apiResponse = ""
    firstTime = True
    cellRow = ActiveCell.Row
    cellColumn = ActiveCell.Column
    i = 1 'Used for incrementing rows
    
    'Adding the header
    Call speciesHeader(-1)
    
    Call getAPIData("https://swapi.dev/api/species/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    
    'Grab the info needed
continueNext:
    'Handles the first page from JSON results
    If firstTime = False Then
        Call getAPIData(Json("next"))
        Set Json = JsonConverter.ParseJson(apiResponse)
    End If
    
    For Each result In Json("results")
        Cells(cellRow + i, cellColumn).Value = result("name")
        Cells(cellRow + i, cellColumn + 1).Value = result("classification")
        Cells(cellRow + i, cellColumn + 2).Value = result("designation")
        Cells(cellRow + i, cellColumn + 3).Value = result("average_height")
        Cells(cellRow + i, cellColumn + 4).Value = result("skin_colors")
        Cells(cellRow + i, cellColumn + 5).Value = result("hair_colors")
        Cells(cellRow + i, cellColumn + 6).Value = result("average_lifespan")
        Cells(cellRow + i, cellColumn + 7).Value = result("homeworld")
        Cells(cellRow + i, cellColumn + 8).Value = result("language")
        
        'Count number of pilots
        x = 0
        For Each Value In result("people")
            x = x + 1
        Next Value
        Cells(cellRow + i, cellColumn + 9).Value = x
        
        'Increment to continue to the next row
        i = i + 1
        firstTime = False
        
    Next result
    
    If IsNull(Json("next")) = False Then GoTo continueNext
    
End Sub

Sub insertData(control As IRibbonControl)
    InsertDataForm.Show
End Sub
