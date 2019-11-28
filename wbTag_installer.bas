Attribute VB_Name = "wbTag_installer"
Option Explicit

'This little VBA module was written by Christoph Kleine and published on internetzkidz.de
'It is licensed under the GNU General Public License v3.0 and can be improved upon and distributed
'For documentation check out https://internetzkidz.de/open-source/wbtag
'Version #0.8

Private Sub wbTag_installer()

    'Declare variables that are going to be used
    Dim thisWorkbook As Workbook
    Dim pageSheet As Worksheet, addedDB As Worksheet
    Dim installed As Boolean
    Dim Headlines As Variant, HeadlinesLength As Integer
    Dim i As Integer, sheetCount As Integer
    Dim installerSuccess As String, installerObsolete As String
    
    'this is where we define the message that explains that wbTag is already installed or that it has been successfully installed
    installerObsolete = "wbTag ist in diesem Arbeitsblatt bereits installiert. Sollte das Programm Probleme bereiten, entferne das wbTag-VBA-Modul und das _wbtagDB Sheet und installiere wbTag erneut."
    installerSuccess = "wbTag wurde erfolgreich installiert und zeichnet Nutzerdaten auf. Binde das wbTag-Pixel als nächstes im Arbeitsmappen-Element ein. Vergiss nicht das _wbTagDB-Sheet und andere Analyse-Sheets auszublenden!"
    
    Set thisWorkbook = ActiveWorkbook
    installed = False
    sheetCount = thisWorkbook.Sheets.Count
    
    'loop through each sheet to see wether there is a _wbTagDB sheet
    For Each pageSheet In thisWorkbook.Worksheets
                
        'if the sheet is already there set installed to true and prompt the user
        If pageSheet.Name = "_wbTagDB" Then
            installed = True
            MsgBox installerObsolete
            
            'Exit installer if DB-sheet is already present
            Exit Sub
        End If
                
    Next pageSheet
    
    'initialize array for column headlines
    Headlines = Array("ID", "Timestamp", "Eventtype", "URL", "Count", "User", "OS", "pageviews", "events", "opens", "saves", "pageAdd")
    HeadlinesLength = UBound(Headlines)
    
    'if the check reveals that there is no _wbTagDB sheet ad it an pre-fill the column headlines
    If installed = False Then
        Set addedDB = thisWorkbook.Sheets.Add(, thisWorkbook.Sheets(sheetCount))
        addedDB.Name = "_wbTagDB"
        addedDB.Cells(1, 1).EntireRow.Font.Bold = True
        addedDB.Cells(1, 1).EntireRow.Interior.ColorIndex = 15
        
        For i = 0 To HeadlinesLength
            addedDB.Cells(1, i + 1).Value = Headlines(i)
        Next i
        
        'Prompt user for end of installation.
        MsgBox installerSuccess
    
    End If
    
    'To install the pageview on-Event-Tag use the following code in the workbook module:
    '### Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    '###    Call workbookTag("pageview")
    '### End Sub

End Sub
