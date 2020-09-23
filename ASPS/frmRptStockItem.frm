VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRptStockItem 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Status"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4215
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Please select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3975
      Begin VB.OptionButton optAll 
         BackColor       =   &H00C0FFC0&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   615
      End
      Begin MSDataListLib.DataCombo dcboClassificationAnd 
         Bindings        =   "frmRptStockItem.frx":0000
         Height          =   360
         Left            =   1800
         TabIndex        =   8
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Type"
         Text            =   "classification"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.OptionButton optClassKind 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Classification:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton optClassification 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Classification:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo dcboKind 
         Bindings        =   "frmRptStockItem.frx":0020
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Category"
         Text            =   "kind"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcboClassification 
         Bindings        =   "frmRptStockItem.frx":0036
         Height          =   360
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Type"
         Text            =   "classification"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kind:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   9
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "and"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1800
         TabIndex        =   7
         Top             =   1560
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdPreviewItem 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoKind 
      Height          =   375
      Left            =   120
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoKind"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoClassification 
      Height          =   375
      Left            =   120
      Top             =   3360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoClassification"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Left            =   -1440
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   495
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   -240
      Width           =   5295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current status of materials."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   832
      TabIndex        =   11
      Top             =   0
      Width           =   2550
   End
End
Attribute VB_Name = "frmRptStockItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loadclass As Boolean
Dim LoadKind As Boolean

Private Sub adoClassification_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoKind_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewItem.FontBold = False
cmdClose.FontBold = True
End Sub

Private Sub cmdPreviewItem_Click()
On Error GoTo ErrHandler

If optClassification.Value = True Then
    Me.MousePointer = vbHourglass
    If dcboClassification.Text = "" Then
        MsgBox "Please select a Classification."
        Me.MousePointer = vbDefault
        Exit Sub
    End If

    If envAmadeus.rscmdRptStockItem_Grouping.State = adStateOpen Then
        envAmadeus.rscmdRptStockItem_Grouping.Close
        Unload rptStockItem
    End If
    
    envAmadeus.cmdRptStockItem_Grouping adoClassification.Recordset.Fields("StockType_ID"), 0, 0
    
    If envAmadeus.rscmdRptStockItem_Grouping.RecordCount = 0 Then
        envAmadeus.rscmdRptStockItem_Grouping.Close
        MsgBox "Cannot find record."
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    rptStockItem.Show
    rptStockItem.WindowState = vbMaximized
    Me.MousePointer = vbDefault
ElseIf optClassKind.Value = True Then
    Me.MousePointer = vbHourglass
    If dcboClassificationAnd.Text = "" Or dcboKind.Text = "" Then
        MsgBox "Please select a Classification or Kind."
        Me.MousePointer = vbDefault
        Exit Sub
    End If

    If envAmadeus.rscmdRptStockItem_Grouping.State = adStateOpen Then
        envAmadeus.rscmdRptStockItem_Grouping.Close
        Unload rptStockItem
    End If
    
    envAmadeus.cmdRptStockItem_Grouping 0, adoClassification.Recordset.Fields("StockType_ID"), _
        adoKind.Recordset.Fields("StockTypeCategory_ID")
    
    If envAmadeus.rscmdRptStockItem_Grouping.RecordCount = 0 Then
        envAmadeus.rscmdRptStockItem_Grouping.Close
        MsgBox "Cannot find record."
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    rptStockItem.Show
    rptStockItem.WindowState = vbMaximized
    Me.MousePointer = vbDefault
ElseIf optAll.Value = True Then
    Me.MousePointer = vbHourglass
    rptInventorySummary.Show
    rptInventorySummary.WindowState = vbMaximized
    Me.MousePointer = vbDefault
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdPreviewItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewItem.FontBold = True
cmdClose.FontBold = False
End Sub

Private Sub dcboClassification_Change()
If loadclass = False Then Exit Sub
If adoClassification.Recordset.Bookmark = 0 Then Exit Sub
adoClassification.Recordset.Bookmark = dcboClassification.SelectedItem
End Sub

Private Sub dcboClassification_GotFocus()
loadclass = True
End Sub

Private Sub dcboClassificationAnd_Change()
If loadclass = False Then Exit Sub
If adoClassification.Recordset.Bookmark = 0 Then Exit Sub
adoClassification.Recordset.Bookmark = dcboClassificationAnd.SelectedItem
adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from tblstocktypecategory where StockType_ID=" & adoClassification.Recordset.Fields("StockType_ID")
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.ReFill
End Sub

Private Sub dcboClassificationAnd_GotFocus()
loadclass = True
End Sub

Private Sub dcboKind_Change()
If LoadKind = False Then Exit Sub
If adoKind.Recordset.Bookmark = 0 Then Exit Sub
adoKind.Recordset.Bookmark = dcboKind.SelectedItem
End Sub

Private Sub dcboKind_GotFocus()
LoadKind = True
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

Call CenterForm(frmRptStockItem, MDIForm1)
loadclass = False
LoadKind = False

Call ConnectDB(adoClassification)
adoClassification.CommandType = adCmdTable
adoClassification.RecordSource = "tblstocktype"
adoClassification.Refresh

Call ConnectDB(adoKind)
adoKind.CommandType = adCmdTable
adoKind.RecordSource = "tblstocktypecategory"
adoKind.Refresh

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdPreviewItem.FontBold = False
cmdClose.FontBold = False
End Sub
