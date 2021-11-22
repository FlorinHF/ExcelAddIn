Attribute VB_Name = "starWarsFunctions"
Function swapiGetListOfPlanets(Optional includeHeader As Boolean) As Variant()
    'Example of function with slow performance, sending multiple requests for each Planet.
    Dim Json As Object
    Dim temArr As Variant
    Dim countPlanets As Integer

    apiResponse = ""

    Call getAPIData("https://swapi.dev/api/planets/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    countPlanets = Json("count")
    
    'Parameter to let the user decide it wants to have the header included or not
    If includeHeader = False Then
        ReDim tempArr(1 To countPlanets, 1 To 8)
        j = 0
    Else
        ReDim tempArr(1 To countPlanets + 1, 1 To 8)
        j = 1
        tempArr(1, 1) = "Planet Name"
        tempArr(1, 2) = "Diameter"
        tempArr(1, 3) = "Climate"
        tempArr(1, 4) = "Gravity"
        tempArr(1, 5) = "Terrain"
        tempArr(1, 6) = "Surface water"
        tempArr(1, 7) = "Population"
        tempArr(1, 8) = "Known residents"
    End If
    
    'Request API calls for each planet
    'Grab the info needed
    For i = 1 To countPlanets
        Call getAPIData("https://swapi.dev/api/planets/" & i & "/")
        Set Planet = JsonConverter.ParseJson(apiResponse)
        tempArr(i + j, 1) = Planet("name")
        tempArr(i + j, 2) = Planet("diameter")
        tempArr(i + j, 3) = Planet("climate")
        tempArr(i + j, 4) = Planet("gravity")
        tempArr(i + j, 5) = Planet("terrain")
        tempArr(i + j, 6) = Planet("surface_water")
        tempArr(i + j, 7) = Planet("population")
        
        'Count number of known residents
        x = 0
        For Each Value In Planet("residents")
            x = x + 1
        Next Value
        tempArr(i + j, 8) = x
        
    Next i
    
    swapiGetListOfPlanets = tempArr

End Function

Function swapiGetListOfPeople(Optional includeHeader As Boolean) As Variant()
    'Example of function with a better performance, sending few requests to get data for People.
    'This function makes API calls when the input is a API link(planetURL, vehicleURL,starshipsURL) and grab the relevant names.
    Dim Json As Object
    Dim firstTime As Boolean
    Dim planetURL As String
    Dim vehicleURL As String
    Dim starshipsURL As String
    
    apiResponse = ""
    firstTime = True
    i = 1 'Used for incrementing rows
    
    Call getAPIData("https://swapi.dev/api/people/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    countPeople = Json("count")
    
    'Parameter to let the user decide it wants to have the header included or not
    If includeHeader = False Then
        ReDim tempArr(1 To countPeople, 1 To 12)
        j = 0
    Else
        ReDim tempArr(1 To countPeople + 1, 1 To 12)
        j = 1
        tempArr(1, 1) = "Name"
        tempArr(1, 2) = "Height"
        tempArr(1, 3) = "Weight"
        tempArr(1, 4) = "Hair Color"
        tempArr(1, 5) = "Skin Color"
        tempArr(1, 6) = "Eye Color"
        tempArr(1, 7) = "Birth year"
        tempArr(1, 8) = "Gender"
        tempArr(1, 9) = "Homeworld"
        tempArr(1, 10) = "Vehicles"
        tempArr(1, 11) = "Starships"
        tempArr(1, 12) = "Film Appearances"
    End If
    
    'Request API calls for results pages for people
    'Grab the info needed
continueNext:
    'Handles the first page from JSON results
    If firstTime = False Then
        Call getAPIData(Json("next"))
        Set Json = JsonConverter.ParseJson(apiResponse)
    End If
    
    For Each result In Json("results")
        tempArr(i + j, 1) = result("name")
        tempArr(i + j, 2) = result("height")
        tempArr(i + j, 3) = result("mass")
        tempArr(i + j, 4) = result("hair_color")
        tempArr(i + j, 5) = result("skin_color")
        tempArr(i + j, 6) = result("eye_color")
        tempArr(i + j, 7) = result("birth_year")
        tempArr(i + j, 8) = result("gender")
        
        'Grab Homeworld URL and return its name
        planetURL = result("homeworld")
        Call getAPIData(planetURL)
        Set Planets = JsonConverter.ParseJson(apiResponse)
        
        tempArr(i + j, 9) = Planets("name")
        
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
                vehiclesNames = vehiclesNames & "; " & Vehicles("name")
            End If
            firstVehicle = False
        Next vehicle
        
        tempArr(i + j, 10) = vehiclesNames
        
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
        
        tempArr(i + j, 11) = starshipsNames
        
        
        'Count the number of film appearances
        x = 0
        For Each Value In result("films")
            x = x + 1
        Next Value
        tempArr(i + j, 12) = x
        
        'Increment to continue to the next row
        i = i + 1
        firstTime = False
        
        
    Next result
    
    If IsNull(Json("next")) = False Then GoTo continueNext
    
    swapiGetListOfPeople = tempArr

End Function
Function swapiGetListOfSpaceships(Optional includeHeader As Boolean) As Variant()
    'Example of function with a better performance, sending few requests to get data for Spaceships/Starships.
    
    Dim Json As Object
    Dim firstTime As Boolean
    
    apiResponse = ""
    firstTime = True
    i = 1 'Used for incrementing rows
    Call getAPIData("https://swapi.dev/api/starships/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    countSpaceships = Json("count")
    
    'Parameter to let the user decide it wants to have the header included or not
    If includeHeader = False Then
        ReDim tempArr(1 To countSpaceships, 1 To 12)
        j = 0
    Else
        ReDim tempArr(1 To countSpaceships + 1, 1 To 12)
        j = 1
        tempArr(1, 1) = "Name"
        tempArr(1, 2) = "Manufacturer"
        tempArr(1, 3) = "Cost"
        tempArr(1, 4) = "Length"
        tempArr(1, 5) = "Max Speed"
        tempArr(1, 6) = "Crew"
        tempArr(1, 7) = "Passengers"
        tempArr(1, 8) = "Cargo Capacity"
        tempArr(1, 9) = "Hyperdrive Rating"
        tempArr(1, 10) = "Starship class"
        tempArr(1, 11) = "Pilots"
        tempArr(1, 12) = "Film Appearances"
    End If
    
    'Request API calls for results pages for people
    'Grab the info needed
continueNext:
    'Handles the first page from JSON results
    If firstTime = False Then
        Call getAPIData(Json("next"))
        Set Json = JsonConverter.ParseJson(apiResponse)
    End If
    
    For Each result In Json("results")
        tempArr(i + j, 1) = result("name")
        tempArr(i + j, 2) = result("manufacturer")
        tempArr(i + j, 3) = result("cost_in_credits")
        tempArr(i + j, 4) = result("length")
        tempArr(i + j, 5) = result("max_atmosphering_speed")
        tempArr(i + j, 6) = result("crew")
        tempArr(i + j, 7) = result("passengers")
        tempArr(i + j, 8) = result("cargo_capacity")
        tempArr(i + j, 9) = result("hyperdrive_rating")
        tempArr(i + j, 10) = result("starship_class")

        'Count the number of pilots
        x = 0
        For Each Value In result("pilots")
            x = x + 1
        Next Value
        tempArr(i + j, 11) = x
        
        'Count the number of film appearances
        x = 0
        For Each Value In result("films")
            x = x + 1
        Next Value
        tempArr(i + j, 12) = x
        
        'Increment to continue to the next row
        i = i + 1
        firstTime = False
        
        
    Next result
    
    If IsNull(Json("next")) = False Then GoTo continueNext
    
    swapiGetListOfSpaceships = tempArr

End Function
Function swapiGetListOfVehicles(Optional includeHeader As Boolean) As Variant()
    'Example of function with a better performance, sending few requests to get data for Vehicles.
    Dim Json As Object
    Dim firstTime As Boolean
    
    apiResponse = ""
    firstTime = True
    i = 1 'Used for incrementing rows
    Call getAPIData("https://swapi.dev/api/vehicles/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    countVehicles = Json("count")
    
    'Parameter to let the user decide it wants to have the header included or not
    If includeHeader = False Then
        ReDim tempArr(1 To countVehicles, 1 To 11)
        j = 0
    Else
        ReDim tempArr(1 To countVehicles + 1, 1 To 11)
        j = 1
        tempArr(1, 1) = "Name"
        tempArr(1, 2) = "Manufacturer"
        tempArr(1, 3) = "Cost"
        tempArr(1, 4) = "Length"
        tempArr(1, 5) = "Max Speed"
        tempArr(1, 6) = "Crew"
        tempArr(1, 7) = "Passengers"
        tempArr(1, 8) = "Cargo Capacity"
        tempArr(1, 9) = "Vehicle class"
        tempArr(1, 10) = "Pilots"
        tempArr(1, 11) = "Film Appearances"
    End If
    
    'Request API calls for results pages for people
    'Grab the info needed
continueNext:
    'Handles the first page from JSON results
    If firstTime = False Then
        Call getAPIData(Json("next"))
        Set Json = JsonConverter.ParseJson(apiResponse)
    End If
    
    For Each result In Json("results")
        tempArr(i + j, 1) = result("name")
        tempArr(i + j, 2) = result("manufacturer")
        tempArr(i + j, 3) = result("cost_in_credits")
        tempArr(i + j, 4) = result("length")
        tempArr(i + j, 5) = result("max_atmosphering_speed")
        tempArr(i + j, 6) = result("crew")
        tempArr(i + j, 7) = result("passengers")
        tempArr(i + j, 8) = result("cargo_capacity")
        tempArr(i + j, 9) = result("vehicle_class")

        'Count the number of pilots
        x = 0
        For Each Value In result("pilots")
            x = x + 1
        Next Value
        tempArr(i + j, 10) = x
        
        'Count the number of film appearances
        x = 0
        For Each Value In result("films")
            x = x + 1
        Next Value
        tempArr(i + j, 11) = x
        
        'Increment to continue to the next row
        i = i + 1
        firstTime = False
        
    Next result
    
    If IsNull(Json("next")) = False Then GoTo continueNext
    
    swapiGetListOfVehicles = tempArr

End Function
'Function swapiGetListOfSpecies(Optional includeHeader As Boolean) As Variant()
'
'    Dim Json As Object
'    Dim firstTime As Boolean
'
'
'    apiResponse = ""
'    firstTime = True
'
'    i = 1 'increment row
'
'    Call getAPIData("https://swapi.dev/api/species/")
'
'    Set Json = JsonConverter.ParseJson(apiResponse)
'    countSpecies = Json("count")
'
'    If includeHeader = False Then
'        ReDim tempArr(1 To countSpecies, 1 To 10)
'        j = 0
'    Else
'        ReDim tempArr(1 To countSpecies + 1, 1 To 10)
'        j = 1
'        tempArr(1, 1) = "Name"
'        tempArr(1, 2) = "Classification"
'        tempArr(1, 3) = "Designation"
'        tempArr(1, 4) = "Average Height"
'        tempArr(1, 5) = "Skin Colors"
'        tempArr(1, 6) = "Hair Color"
'        tempArr(1, 7) = "Average Lifespan"
'        tempArr(1, 8) = "Homeworld"
'        tempArr(1, 9) = "Language"
'        tempArr(1, 10) = "Known People"
'    End If
'
'continueNext:
'
'    If firstTime = False Then
'        Call getAPIData(Json("next"))
'        Set Json = JsonConverter.ParseJson(apiResponse)
'    End If
'
'    For Each result In Json("results")
'        tempArr(i + j, 1) = result("name")
'        tempArr(i + j, 2) = result("classification")
'        tempArr(i + j, 3) = result("designation")
'        tempArr(i + j, 4) = result("average_height")
'        tempArr(i + j, 5) = result("skin_colors")
'        tempArr(i + j, 6) = result("hair_colors")
'        tempArr(i + j, 7) = result("average_lifespan")
'        tempArr(i + j, 8) = result("homeworld")
'        tempArr(i + j, 9) = result("language")
'
'        'Count number of people
'        x = 0
'        For Each Value In result("people")
'            x = x + 1
'        Next Value
'
'        tempArr(i + j, 10) = x
'
'
'        i = i + 1
'        firstTime = False
'
'
'    Next result
'
'    If IsNull(Json("next")) = False Then GoTo continueNext
'
'    swapiGetListOfSpecies = tempArr
'
'End Function
Function swapiGetNumberOfPlanets()
    'Returns how many planets are in the API
    Dim Json As Object
    Call getAPIData("https://swapi.dev/api/planets/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    
    swapiGetNumberOfPlanets = Json("count")
    
End Function

Function swapiGetNumberOfPeople()
    'Returns how many people are in the API
    Dim Json As Object
    Call getAPIData("https://swapi.dev/api/people/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    
    swapiGetNumberOfPeople = Json("count")
    
End Function

Function swapiGetNumberOfSpaceships()
    'Returns how many spaceships are in the API
    Dim Json As Object
    Call getAPIData("https://swapi.dev/api/starships/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    
    swapiGetNumberOfSpaceships = Json("count")
    
End Function

Function swapiGetNumberOfSpecies()
    'Returns how many species are in the API
    Dim Json As Object
    Call getAPIData("https://swapi.dev/api/species/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    
    swapiGetNumberOfSpecies = Json("count")
    
End Function

Function swapiGetNumberOfVehicles()
    'Returns how many vehicles are in the API
    Dim Json As Object
    Call getAPIData("https://swapi.dev/api/vehicles/")
    Set Json = JsonConverter.ParseJson(apiResponse)
    
    swapiGetNumberOfVehicles = Json("count")
End Function

Function swapiGetAPIData(strUrl As String, Optional parameter1 As String)
    'Returns the API call in JSON format as text in Excel cell
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
    'the user can insert the column name for retrieve the value it's looking for
    If parameter1 = "" Then
        swapiGetAPIData = apiResponse
    Else
        Set Json = JsonConverter.ParseJson(apiResponse)
        swapiGetAPIData = Json(parameter1)
    End If

End Function


