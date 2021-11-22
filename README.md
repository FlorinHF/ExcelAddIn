# Excel Add-In using Star Wars API
## Introduction
This is an Excel Add-in to connect to Star Wars API available on https://swapi.dev/. All the VBA functions and subs can be seen using the VB Editor. This Add-in is compatible with all Excel versions starting with Excel 2007 on Windows operating system. Use [`Star Wars Universe AddIn.xlam`](https://github.com/FlorinHF/ExcelAddIn/blob/main/Star%20Wars%20Universe%20AddIn.xlam)   file to get started

## Installation

Firstly, you need to have the Developer tab displayed in the ribbon. The **Developer** tab isn't displayed by default, but you can add it to the ribbon.
  - On the **File** tab, go to **Options** > **Customize Ribbon**.
  - Under **Customize the Ribbon** and under **Main Tabs**, select the **Developer** check box.

Navigate to **Developer** tab then click on **Add-Ins** (see image below). There you can click on **Browse...** and add the [`Star Wars Universe AddIn.xlam`](https://github.com/FlorinHF/ExcelAddIn/blob/main/Star%20Wars%20Universe%20AddIn.xlam) file. At the end of the ribbon you should be able to see **STAR WARS** tab added.

![image](https://user-images.githubusercontent.com/22075409/142901052-ad633847-4b0a-4523-ad37-4944ab6b71f7.png)

## How to use the Add-In
This add-in has main two methods to fetch data from Star Wars API (https://swapi.dev/) :
  - Insert data using the **STAR WARS** tab
  - Insert data using the `'=swapi*'` formulae

### Insert data using the **STAR WARS** tab
![image](https://user-images.githubusercontent.com/22075409/142904665-3dc1e1ab-7b80-404f-80bb-d49f9916f3fb.png)

**Hello** group is simply a message box thats displaying your name.

**Performance Comparison** group contains two buttons: 
  - *Insert Slow* - fetches data from the API using multiple requests for each record
  - *Insert Fast* - fetches data from the API using few requests for each results page, resulting in better performance

**Insert Datasets** group contains a variety of datasets from the API formatted in different ways.

**Test** group contains a *Insert Data Form*(custom) that lets the user choose what dataset to insert.

*All **Insert...** buttons will be inserted in the active cell. If the user wants to insert the dataset into C6, select C6 and click any insert button from the *STAR WARS* tab*

### Insert data using the `'=swapi*'` formulae
![image](https://user-images.githubusercontent.com/22075409/142909135-885fbc46-6032-481a-92ed-437cd1467d42.png)

All the formulae that start with `=swapi*` are fetching data from the STAR WARS API. **Some of the formulae take longer to run as it contains more data.**

Formulae returning a single value
```
=swapiGetAPIData (URL, *Optional Parameter1)
=swapiGetNumberOfPlanets
=swapiGetNumberOfPeople
=swapiGetNumberOfSpecies
=swapiGetNumberOfSpaceships
=swapiGetNumberOfVehicles

```
Examples:
`swapiGetAPIData(URL, *Optional Parameter1)`
![image](https://user-images.githubusercontent.com/22075409/142911499-6bde26a9-51d9-43e4-a732-fb54bc44b06e.png) ![image](https://user-images.githubusercontent.com/22075409/142912535-116ae6a5-fb22-45a9-a7f0-1d1d2dd4954c.png)

`=swapiGetNumberOfPlanets`

![image](https://user-images.githubusercontent.com/22075409/142911669-fea0e406-ebbd-4e26-8906-17eebf5cf296.png)

Formulae returning multiple values (hit CRTL+Shift+Enter on keyboard)
```
=swapiGetListOfPlanets(Optional includeHeader)
=swapiGetListOfPeople(Optional includeHeader)
=swapiGetListOfSpaceships(Optional includeHeader)
=swapiGetListOfVehicles(Optional includeHeader)
```
Examples:
`=swapiGetListOfPlanets(Optional includeHeader)`
![image](https://user-images.githubusercontent.com/22075409/142911971-3f61f868-2527-4e4c-a346-79eae492fff6.png)
![image](https://user-images.githubusercontent.com/22075409/142912322-6e2cac85-1d08-4392-aa0f-1cd9724d5a3a.png)




