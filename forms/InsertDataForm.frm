VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertDataForm 
   Caption         =   "Insert Data Form"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4485
   OleObjectBlob   =   "InsertDataForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InsertDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
        
    With Me.ComboBox1
        .Clear
        .AddItem "Spaceships"
        .AddItem "Species"
        .AddItem "People"
        .AddItem "Planets"
        .AddItem "Vehicles"
    
    End With
End Sub

Private Sub CommandButton1_Click()
    Dim FakeControl As IRibbonControl

    datasetToInsert = Me.ComboBox1.Value
    
    If datasetToInsert = "" Then
        MsgBox "Please select a dataset to insert"
    End If
    
    Select Case datasetToInsert
        Case "Spaceships"
            Call getListOfSpaceships(FakeControl)
        Case "Species"
            Call getListOfSpecies(FakeControl)
        Case "People"
            Call getListOfPeople(FakeControl)
        Case "Planets"
            Call getListOfPlanets(FakeControl)
        Case "Vehicles"
            Call getListOfVehicles(FakeControl)
    End Select
    
    Unload InsertDataForm
    
    
End Sub



