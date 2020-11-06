VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAsyncFetch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Async Fetch"
   ClientHeight    =   6492
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   9552
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6492
   ScaleWidth      =   9552
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Stop Fetching"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2028
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6072
      Width           =   6372
   End
   Begin MSDataGridLib.DataGrid dtg 
      Height          =   4980
      Left            =   108
      TabIndex        =   3
      Top             =   972
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   8784
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16776927
      HeadLines       =   1
      RowHeight       =   19
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   -1  'True
         AllowSizing     =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit "
      Height          =   336
      Left            =   8484
      TabIndex        =   2
      Top             =   6072
      Width           =   972
   End
   Begin VB.CommandButton cmdFetch 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Fetch Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6072
      Width           =   1764
   End
   Begin VB.Label lblDemo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   864
      Left            =   108
      TabIndex        =   4
      Top             =   48
      Width           =   9372
   End
End
Attribute VB_Name = "frmAsyncFetch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim blstatus As Boolean
Dim nRecords As Long

Dim conn As ADODB.Connection

Private Sub cmdCancel_Click()
   blstatus = True
End Sub

Private Sub cmdExit_Click()
    Set rs = Nothing
    Set conn = Nothing
    Unload Me
    End
End Sub

Private Sub cmdFetch_Click()

    Set dtg.DataSource = Nothing
    blstatus = False
    nRecords = 0
    
    Set conn = New ADODB.Connection
    cmdCancel.Caption = "&Stop Fetching (Connecting to Database...)"
    
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\test.mdb"
    conn.CursorLocation = adUseClient 'required
            
    Set rs = New ADODB.Recordset
    cmdCancel.Caption = "&Stop Fetching (Executing SQL....)"
    
    rs.CursorLocation = adUseClient 'required
        
    rs.Properties(89).Value = 1 'required (initial fetch size)
    rs.Properties(88).Value = 100 'required (background fetch size)
            
    'Fetch the query asynchronously
    rs.Open "select * from test order by fld1", conn, adOpenKeyset, adLockOptimistic, adAsyncFetchNonBlocking
            
    Set rs.ActiveConnection = Nothing
    conn.Close
    Set conn = Nothing
   
    cmdCancel.SetFocus
End Sub

'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'stop retreival when escape pressed
    If KeyCode = vbKeyEscape Then
        blstatus = True
    Else
        blstatus = False
    End If
End Sub

Private Sub Form_Load()
    lblDemo.Caption = "This is a demo to show Asynchronous fetching of records." & _
                    "An .mdb table (Test.mdb) has been used having more than 18000 records." & vbCrLf & _
                    vbCrLf & "PRESS 'ESCAPE' OR PRESS THE 'STOP FETCHING' BUTTON TO STOP FETCHING THE RECORDS."
    Me.KeyPreview = True
End Sub

Private Sub rs_FetchComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    MsgBox "Total records Fetched: " & rs.RecordCount, vbInformation, "Records Fetched"
    Set dtg.DataSource = rs
    dtg.SetFocus
End Sub

Private Sub rs_FetchProgress(ByVal Progress As Long, ByVal MaxProgress As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If blstatus = True Then
        rs.Cancel
        nRecords = Progress
        cmdCancel.Caption = "Records Fetched: " & nRecords & ""
    Else
        cmdCancel.Caption = "&Stop Fetching (Records Fetched: " & Progress & ")" 'gives the progress of fetching
    End If
End Sub
