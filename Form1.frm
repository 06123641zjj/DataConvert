VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox DBPath 
      Height          =   2340
      Left            =   4935
      TabIndex        =   2
      Top             =   495
      Width           =   2700
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2730
      Left            =   1005
      TabIndex        =   1
      Top             =   135
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   4815
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   15
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   435
      Left            =   8295
      TabIndex        =   0
      Top             =   1845
      Width           =   1260
   End
   Begin VB.Label lbl 
      Caption         =   "Double click the folder contain the database to select the path to database, then click Convert and select the xls file"
      Height          =   825
      Left            =   7995
      TabIndex        =   5
      Top             =   435
      Width           =   2205
   End
   Begin VB.Label lblPathToDB 
      Caption         =   "Path to database:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4980
      TabIndex        =   4
      Top             =   120
      Width           =   3090
   End
   Begin VB.Label lblRow 
      Height          =   390
      Left            =   8310
      TabIndex        =   3
      Top             =   1425
      Width           =   1125
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdConvert_Click()
    Dim recordSet As ADODB.recordSet
    Dim xlsFilePath As String
    ' Check if all targe database exist
    If checkDatabase() <> "0" Then
        MsgBox ("Path to database is wrong. Miss " & checkDatabase())
        Debug.Print ("Path to database is wrong. Miss " & checkDatabase())
        Exit Sub
    Else
        'Get the path to the Excle file
        cdFile.ShowOpen
        
        'back up target database
        'Call backUpdatabase
        
        'Open these target database
        Call openDatabase(CStr(DBPath.path))
        
        'get path to xls file
        xlsFilePath = cdFile.fileName
        If InStr(xlsFilePath, ".xls") = 0 Then
            MsgBox ("Please select the .xls file to be convert.")
        Else
            'Start convert
            convert (xlsFilePath)
            lblRow.Caption = "DONE"
        End If
        
        'Close Database
        Call closeDB
    End If
       
    
End Sub

'This method convert a excel file to a recordSet
Function getRecordSetForExcels(sFilePath As String) As ADODB.recordSet
    Const adOpenStatic = 3
    Const adLockOptimistic = 3
    Const adCmdText = &H1
    Dim objRecordset As New ADODB.recordSet
    
    Set objConnection = CreateObject("ADODB.Connection")
    'Set objRecordset = CreateObject("ADODB.Recordset")
      
    objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & sFilePath & ";Extended Properties=""Excel 8.0;HDR=Yes;"";"
    
    objRecordset.Open "Select * FROM [RO HISTORY 8-15-2017$]", _
        objConnection, adOpenStatic, adLockOptimistic, adCmdText
        
    Set getRecordSetForExcels = objRecordset
End Function



Private Sub convert(path As String)
    Dim rsData As ADODB.recordSet 'Excel
    Dim ct As Integer
    Dim row As Collection
    Set row = New Collection
    
    
    'need database: invoice, items, act
    
    'Convert excel to recordSet "rsData"
    Set rsData = getRecordSetForExcels(path)

    ct = 0
    'Go to first row of recordSet
    rsData.MoveFirst
    Do While Not rsData.EOF
        ct = ct + 1
        'read current row, put into a collection
        Call getCurrRow(rsData, row)
        'write current row to database
        Call writeToDB(row)
        
        Set row = New Collection
        'Show number of row have been converted
        lblRow.Caption = "In row: " & CStr(ct) & "/" & rsData.RecordCount
        'Go to next record
        rsData.MoveNext
        DoEvents
    Loop
    
    d4flush dbInvoice
    d4unlock dbInvoice
     
End Sub

'Use the info stored in collection "row", store to database
'So far, just store info to INVOICE database
Private Sub writeToDB(row As Collection)
    Call d4appendBlank(dbInvoice)
    Call f4assign(d4field(dbInvoice, "INVOICE"), getInvoiceNum(row("INVOICE")))
    Call f4assign(d4field(dbInvoice, "DATE"), Format(row("DATE"), "yyyymmdd"))
    Call f4assign(d4field(dbInvoice, "CDATE"), Format(row("CDATE"), "yyyymmdd"))
    Call f4assign(d4field(dbInvoice, "CONAME"), row("CONAME"))
    Call f4assign(d4field(dbInvoice, "ADDRESS"), row("ADDRESS"))
    Call f4assign(d4field(dbInvoice, "CITY"), row("CITY"))
    Call f4assign(d4field(dbInvoice, "STATE"), row("STATE"))
    Call f4assign(d4field(dbInvoice, "ZIP"), row("ZIP"))
    Call f4assign(d4field(dbInvoice, "EMAIL"), row("EMAIL"))
    Call f4assign(d4field(dbInvoice, "PHONE"), row("PHONE"))
    Call f4assign(d4field(dbInvoice, "PHONE1"), row("PHONE1"))
    Call f4assign(d4field(dbInvoice, "TIMEIN"), row("TIMEIN"))
    Call f4memoAssign(d4field(dbInvoice, "lmemo"), row("lmemo"))
    Call f4assign(d4field(dbInvoice, "YEAR"), CInt(row("YEAR")))
    Call f4assign(d4field(dbInvoice, "MAKE"), row("MAKE"))
    Call f4assign(d4field(dbInvoice, "MODEL"), row("MODEL"))
    Call f4assign(d4field(dbInvoice, "VIN"), row("VIN"))
    Call f4assign(d4field(dbInvoice, "LICNO"), row("LICNO"))
    Call f4assign(d4field(dbInvoice, "COLOR"), row("COLOR"))
    Call f4assign(d4field(dbInvoice, "mileage"), row("mileage"))
    Call f4assign(d4field(dbInvoice, "milesout"), row("milesout"))
    Call f4assign(d4field(dbInvoice, "Tag3"), "T")
End Sub



'get the information in current row of recordset, do some conversion and eliminate the Null data, store to a collection called "row"
'So far, just extract info for INVOICE database (Format the phone number) store to a collection called "row"
Private Sub getCurrRow(rsData As ADODB.recordSet, row As Collection)
    Dim homePhone As String
    Dim workPhone As String
    Dim notes As String
    Call row.Add(getData(rsData, "R/O NUMBER"), "INVOICE")
    Call row.Add(getData(rsData, "OPEN DATE"), "DATE")
    Call row.Add(getData(rsData, "DATE CLOSED"), "CDATE")
    Call row.Add(getData(rsData, "CUSTOMER NAME"), "CONAME")
    Call row.Add(getData(rsData, "CUSTOMER ADDRESS"), "ADDRESS")
    Call row.Add(getData(rsData, "CITY"), "CITY")
    Call row.Add(getData(rsData, "CUSTOMER STATE"), "STATE")
    Call row.Add(getData(rsData, "CUSTOMER ZIP"), "ZIP")
    Call row.Add(getData(rsData, "CUSTOMER E-MAIL"), "EMAIL")
    homePhone = getData(rsData, "HOME PHONE1") & getData(rsData, "HOME PHONE2") & getData(rsData, "HOME PHONE3")
    Call row.Add(FormatPhone(homePhone), "PHONE")
    workPhone = getData(rsData, "WORK PHONE1") & getData(rsData, "WORK PHONE2") & getData(rsData, "WORK PHONE3")
    Call row.Add(FormatPhone(workPhone), "PHONE1")
    Call row.Add(getData(rsData, "TIME IN HR") & ":" & getData(rsData, "TIME IN MIN") & ":" & "00", "TIMEIN")
    notes = "CELL PHONE: " & getData(rsData, "CELL PHONE AREA CODE") & getData(rsData, "CELL PHONE EXCHANGE") & getData(rsData, "CELL PHONE NUMBER") & vbCrLf
    notes = "SERVICE ADVISOR NAME: " & getData(rsData, "SERVICE ADVISOR NAME") & vbCrLf & " NEXT SERVICE DESCRIPTION: " & getData(rsData, "NEXT SERVICE DESCRIPTION") & "."
    notes = notes & vbCrLf & "SERVICE WRITER: " & getData(rsData, "SERVICE WRITER")
    notes = notes & vbCrLf & "PRE-WRITE NUMBER: " & getData(rsData, "PRE-WRITE NUMBER")
    notes = notes & vbCrLf & "DEFAULT LABOR LEVEL: " & getData(rsData, "DEFAULT LABOR LEVEL")
    notes = notes & vbCrLf & "COME BACK Y/N: " & getData(rsData, "COME BACK Y/N")
    notes = notes & vbCrLf & "VEHICLE LINE NUMBER: " & getData(rsData, "VEHICLE LINE NUMBER")
    notes = notes & vbCrLf & "CUSTOMER ACCOUNT NUMBER: " & getData(rsData, "CUSTOMER ACCOUNT NUMBER")
    notes = notes & vbCrLf & "MARKETING FOLLOWUP: " & getData(rsData, "MARKETING FOLLOWUP")
    notes = notes & vbCrLf & "SCHEDULED MAINTENANCE: " & getData(rsData, "SCHEDULED MAINTENANCE")
    notes = notes & vbCrLf & "JOB NUMBER: " & getData(rsData, "JOB NUMBER")
    notes = notes & vbCrLf & "CASH/RECEIVABLE: " & getData(rsData, "CASH/RECEIVABLE")
    notes = notes & vbCrLf & "SERVICE/BODY SHOP: " & getData(rsData, "SERVICE/BODY SHOP")
    notes = notes & vbCrLf & "CUSTOMER/WARRANTY/INTERNAL: " & getData(rsData, "CUSTOMER/WARRANTY/INTERNAL")
    notes = notes & vbCrLf & "TIME IN HR: " & getData(rsData, "TIME IN HR")
    notes = notes & vbCrLf & "TIME IN MIN: " & getData(rsData, "TIME IN MIN")
    notes = notes & vbCrLf & "TAXABLE FLAG: " & getData(rsData, "TAXABLE FLAG")
    notes = notes & vbCrLf & "DELIVERY DATE: " & getData(rsData, "DELIVERY DATE MONTH") & "/" & getData(rsData, "DELIVERY DATE DAY") & "/" & getData(rsData, "DELIVERY DATE YEAR")
    
    notes = notes & vbCrLf & "ODOMETER AT DELIVERY: " & getData(rsData, "ODOMETER AT DELIVERY")
    notes = notes & vbCrLf & "FIRST USE: " & getData(rsData, "FIRST USE MONTH") & "/" & getData(rsData, "FIRST USE DAY") & "/" & getData(rsData, "FIRST USE YEAR")

    notes = notes & vbCrLf & "INSPECTION MONTH: " & getData(rsData, "INSPECTION MONTH")
    notes = notes & vbCrLf & "NEXT SERVICE DATE(MM/YY): " & getData(rsData, "NEXT SERVICE DATE MONTH") & "/" & getData(rsData, "NEXT SERVICE DATE YEAR")
    

    notes = notes & vbCrLf & "PREVIOUS SERVICE ODOMETER: " & getData(rsData, "PREVIOUS SERVICE ODOMETER")
    notes = notes & vbCrLf & "SERVICE CONTRACT: " & getData(rsData, "SERVICE CONTRACT")
    notes = notes & vbCrLf & "SERVICE CONTRACT TERM: " & getData(rsData, "SERVICE CONTRACT TERM")
    
    notes = notes & vbCrLf & "SERVICE CONTRACT EXPIRES(MM/YY): " & getData(rsData, "SERVICE CONTRACT EXPIRES MONTH") & "/" & getData(rsData, "SERVICE CONTRACT EXPIRES YEAR")

    notes = notes & vbCrLf & "SERVICE CONTRACT EXPIRES ODOMETER: " & getData(rsData, "SERVICE CONTRACT EXPIRES ODOMETER")
    notes = notes & vbCrLf & "NEW/USED/OTHER: " & getData(rsData, "NEW/USED/OTHER")
    notes = notes & vbCrLf & "WARRANTY FRANCHISE: " & getData(rsData, "WARRANTY FRANCHISE")
    notes = notes & vbCrLf
    notes = notes & vbCrLf & "W/C LABOR: " & getData(rsData, "W/C LABOR")
    notes = notes & vbCrLf & "W/C LABOR COST: " & getData(rsData, "W/C LABOR COST")
    notes = notes & vbCrLf & "W/C PARTS: " & getData(rsData, "W/C PARTS")
    notes = notes & vbCrLf & "W/C PARTS COST: " & getData(rsData, "W/C PARTS COST")
    notes = notes & vbCrLf & "W/C GAS/OIL/GRS: " & getData(rsData, "W/C GAS/OIL/GRS")
    notes = notes & vbCrLf & "W/C G/O/G COST: " & getData(rsData, "W/C G/O/G COST")
    notes = notes & vbCrLf & "W/C SUBLET: " & getData(rsData, "W/C SUBLET")
    notes = notes & vbCrLf & "W/C SUBLET COST: " & getData(rsData, "W/C SUBLET COST")
    notes = notes & vbCrLf & "W/C DEDUCTIBLE: " & getData(rsData, "W/C DEDUCTIBLE")
    notes = notes & vbCrLf & "W/C TOTAL: " & getData(rsData, "W/C TOTAL")
    notes = notes & vbCrLf & "W/C CONTROL: " & getData(rsData, "W/C CONTROL")
    notes = notes & vbCrLf & "INT CONTROL: " & getData(rsData, "INT CONTROL")
    notes = notes & vbCrLf
    notes = notes & vbCrLf & "C/P LABOR: " & getData(rsData, "C/P LABOR")
    notes = notes & vbCrLf & "C/P LABOR COST: " & getData(rsData, "C/P LABOR COST")
    notes = notes & vbCrLf & "C/P PARTS: " & getData(rsData, "C/P PARTS")
    notes = notes & vbCrLf & "C/P PARTS COST: " & getData(rsData, "C/P PARTS COST")
    notes = notes & vbCrLf & "C/P GAS/OIL/GRS: " & getData(rsData, "C/P GAS/OIL/GRS")
    notes = notes & vbCrLf & "C/P G/O/G COST: " & getData(rsData, "C/P G/O/G COST")
    notes = notes & vbCrLf & "SHOP SUPPLY COST: " & getData(rsData, "SHOP SUPPLY COST")
    notes = notes & vbCrLf & "C/P SUBLET: " & getData(rsData, "C/P SUBLET")
    notes = notes & vbCrLf & "C/P SUBLET COST: " & getData(rsData, "C/P SUBLET COST")
    notes = notes & vbCrLf & "C/P TOTAL: " & getData(rsData, "C/P TOTAL")
    notes = notes & vbCrLf & "C/P TAX: " & getData(rsData, "C/P TAX")
    notes = notes & vbCrLf & "C/P A/R CONTROL: " & getData(rsData, "C/P A/R CONTROL")
    notes = notes & vbCrLf & "C/P CHARGE: " & getData(rsData, "C/P CHARGE")
    notes = notes & vbCrLf & "C/P CASH: " & getData(rsData, "C/P CASH")
    notes = notes & vbCrLf & "TAXABLE AMT: " & getData(rsData, "TAXABLE AMT")
    notes = notes & vbCrLf & "TECH NUMBER: " & getData(rsData, "TECH NUMBER")
    
    
    Call row.Add(notes, "lmemo")
    Call row.Add(getData(rsData, "YEAR"), "YEAR")
    Call row.Add(getData(rsData, "MAKE"), "MAKE")
    Call row.Add(getData(rsData, "MODEL"), "MODEL")
    Call row.Add(getData(rsData, "VIN"), "VIN")
    Call row.Add(getData(rsData, "LIC"), "LICNO")
    Call row.Add(getData(rsData, "COLOR"), "COLOR")
    Call row.Add(getData(rsData, "ODOMETER REEDING IN"), "mileage")
    Call row.Add(getData(rsData, "ODOMETER READING OUT"), "milesout")
End Sub


Private Function getData(rsData As ADODB.recordSet, columnName As String) As String
    If IsNull(rsData(columnName)) Then
        getData = ""
    Else
        getData = rsData(columnName)
    End If
    
End Function

'set default database path
Private Sub Form_Load()
    'DBPath.path = "C:\Users\Jingjie\Desktop\newDB"

    DBPath.path = Environ$("USERPROFILE") & "\Desktop"

End Sub

'check the existance of needed database
Private Function checkDatabase() As String
    Dim database(2) As String
    Dim I As Integer

    database(0) = "invoice.dbf"
    database(1) = "items.dbf"
    database(2) = "act.dbf"
    
    For I = 0 To 2 Step 1
        If Not Dir(DBPath.path & "\" & database(I)) <> "" Then
            checkDatabase = database(I)
            Exit Function
        End If
    Next I
    checkDatabase = "0"
End Function

Private Sub backUpdatabase()

    If Dir(DBPath.path & "\backUpDatabaseForConvert", vbDirectory) = "" Then
        MkDir DBPath.path & "\backUpDatabaseForConvert\"
    End If
    'FileCopy DBPath.path & "\invoice.dbf", Environ$("USERPROFILE") & "\Desktop\backUpDatabaseForConvert\invoice.dbf"
    'FileCopy DBPath.path & "\invoice.dbf", Environ$("USERPROFILE") & "\Desktop\backUpDatabaseForConvert\"
    FileCopy DBPath.path & "\invoice.dbf", DBPath.path & "\backUpDatabaseForConvert\invoice.dbf"

'    Dim cmd1 As String
'    Dim cmd2 As String
'    cmd1 = "cd " & DBPath.path & "\backUpDatabaseForConvert"
'    cmd2 = "xcopy " & DBPath.path & "\invoice.dbf"
'    Call Shell(cmd1, vbHide)
'    Call Shell(cmd2, vbHide)
End Sub


'test code
Private Sub testAddRow()
    Call d4appendBlank(dbAct)
    Call f4assign(d4field(dbAct, "STOCKNO"), "A123")
    d4flush dbAct
    d4unlock dbAct
End Sub

' if we can not use the number in excel file directly, we will need to implement this method
Private Function getInvoiceNum(noFromExcel As Integer) As String
    getInvoiceNum = noFromExcel
    'getInvoiceNum = noFromExcel
End Function

Private Function getAccountNum(noFromExcel As Integer) As Integer
    getAccountNum = noFromExcel
End Function
