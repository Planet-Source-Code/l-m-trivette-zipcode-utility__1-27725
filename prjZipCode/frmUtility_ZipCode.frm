VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUtility_ZipCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ZipCode Search Utility"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   Icon            =   "frmUtility_ZipCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6510
   Begin VB.Frame Frame2 
      Caption         =   "Search Results"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   6255
      Begin MSComctlLib.ListView ListView1 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "City"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "State"
            Object.Width           =   1032
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Zipcode"
            Object.Width           =   1667
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Areacode"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "County"
            Object.Width           =   2390
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search "
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton Command2 
         Caption         =   "&Print"
         Height          =   495
         Left            =   3840
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "Zip"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "County"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "Area"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "State"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Search"
         Height          =   495
         Left            =   2400
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   3975
      End
      Begin VB.OptionButton Opt_Category 
         Caption         =   "City"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Search Criteria"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   480
         Width           =   3495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "frmUtility_ZipCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Written by L. "Mike" Trivette 10/1/2001
' Please send me comments at mtrivette@yahoo.com - ICQ 129312768.
' Dont forget to give credit where credit is due.
'
'
Dim category As String ' Public varibale that describes the chosen category to search from
Dim strDatabase As String ' Public varibale that describes the location of the database

Private Sub printfunc()
    ' just a very quick and simple print routine.
    ' this will definately need some work.
    '
    '
    If ListView1.ListItems.Count = 0 Then Exit Sub
    CommonDialog1.Copies = 1
    CommonDialog1.ShowPrinter
    Printer.Copies = CommonDialog1.Copies
    Printer.Print " "
    Printer.Print Tab(95); Now()
    Printer.Print " "
    Printer.Print " "
    For i = 1 To ListView1.ListItems.Count
        Printer.Print Tab(5); ListView1.ListItems(i).Text;
        Printer.Print Tab(25); ListView1.ListItems(i).SubItems(1);
        Printer.Print Tab(35); ListView1.ListItems(i).SubItems(2);
        Printer.Print Tab(45); ListView1.ListItems(i).SubItems(3);
        Printer.Print Tab(55); ListView1.ListItems(i).SubItems(4)
    Next i
    Printer.EndDoc
End Sub

Private Sub load_zips()
    ' This sub searches then load the found data into the listview control
    '
    '
    '
    Dim ii As Long ' Set Variable to hold loop integer
    Set dbs = OpenDatabase(strDatabase) ' Open database
    Set rst = dbs.OpenRecordset("SELECT * FROM zipcodes where " & (category) & " = '" & (txtSearch.Text) & "';") ' Search for data
    
    ListView1.ListItems.Clear ' Clear out the listview for the new data
    
    rst.MoveLast ' Populate the recordset - Very important
    rst.MoveFirst
    
    For ii = 1 To rst.RecordCount ' Main loop starts here
    Set itmx = ListView1.ListItems.Add(, , "" & rst.Fields("city"))
               itmx.SubItems(1) = "" & rst.Fields("state")
               itmx.SubItems(2) = "" & rst.Fields("zip")
               itmx.SubItems(3) = "" & rst.Fields("area")
               itmx.SubItems(4) = "" & rst.Fields("county")
    rst.MoveNext
    Next ii ' Main loop ends here.
    
    rst.Close ' close recordset
    dbs.Close ' close database
End Sub

Private Sub Command1_Click()
    If category <> "" Then load_zips
End Sub

Private Sub Command2_Click()
    If ListView1.ListItems.Count > 0 Then printfunc
End Sub

Private Sub Form_Load()
    ' Variable describes the location of the zipcodes.mdb database
    ' You may have to adjust this file location
    '
    strDatabase = App.Path & "\zipcodes.mdb"
End Sub

Private Sub Opt_Category_Click(Index As Integer)
    category = Opt_Category(Index).Caption
End Sub
