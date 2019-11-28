Attribute VB_Name = "wbTag"
Option Explicit

'This little VBA module was written by Christoph Kleine and published on internetzkidz.de
'It is licensed under the GNU General Public License v3.0 and can be improved upon and distributed
'For documentation check out https://internetzkidz.de/open-source/wbtag
'Version #0.8

Function workbookTag(usageType As String)
    
    'Define variable types
    Dim eventType As String, Timestamp As Date, Url As String, ID As String, Count As Integer, User As String, OS As String
    Dim pageviews As Integer, events As Integer, saves As Integer, opens As Integer, pageAdd As Integer
    Dim dataLine() As Variant
    Dim TheNow As Date
    
    'Pre calculate current time and date for timestamp and ID usage
    TheNow = Now()
    
    'Define variables to be used
    eventType = usageType
    Timestamp = Format(TheNow, "DD.MM.YYYY hh:mm:ss")
    Url = "/" & Replace(ActiveSheet.Name, " ", "-")
    ID = Format(TheNow, "YYYYMMDDhhmmss") & Left(eventType, 1)
    Count = 1
    User = Application.UserName
    OS = Application.OperatingSystem
    
    'Calculate numeric / sum values for datamodel / I hate Pivot Tables
    If eventType = "pageview" Then pageviews = 1
    If eventType = "event" Then events = 1
    If eventType = "open" Then opens = 1
    If eventType = "save" Then saves = 1
    If eventType = "newpage" Then pageAdd = 1
    
    'Bundle the acquired information in an array
    dataLine = Array(ID, Timestamp, eventType, Url, Count, User, OS, pageviews, events, opens, saves, pageAdd)
    
    'send it to the database sheet and create rows
    wbTag2Database (dataLine)
    
End Function

Private Function wbTag2Database(dataset As Variant)

    'Define variable types
    Dim wbTagDB As Worksheet
    Dim lastRow As Integer, nextRow As Integer
    Dim i As Integer, datasetLength As Integer
    
    'Define the database sheet
    'if this causes issues look into installer module
    'or make sure your sheet includes a sheet by the name of "_wbTagDB"
    Set wbTagDB = thisWorkbook.Sheets("_wbTagDB")
    
    'find last row to determine row for next dataset line
    lastRow = wbTagDB.Range("A1").CurrentRegion.Rows.Count
    nextRow = lastRow + 1
    datasetLength = UBound(dataset)
    
    
    'print dataset to last row in wbTagDB table
    For i = 0 To datasetLength
        wbTagDB.Cells(nextRow, i + 1).Value = dataset(i)
    Next i

End Function
