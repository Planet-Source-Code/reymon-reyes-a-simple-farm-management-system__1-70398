VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRequestDetails 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Request Details"
   ClientHeight    =   7560
   ClientLeft      =   5415
   ClientTop       =   1740
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   7695
   Begin VB.TextBox Text1 
      DataField       =   "Request_ID"
      DataSource      =   "adoRequestDetail"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   7920
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoRequestDetail 
      Height          =   375
      Left            =   1800
      Top             =   7560
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "adoRequestDetail"
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
   Begin VB.TextBox txtRequestQuantity 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   5
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CommandButton cmdViewAll 
      Caption         =   "&View All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "View all items"
      Top             =   6840
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoKind 
      Height          =   375
      Left            =   -240
      Top             =   8280
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
      Left            =   -240
      Top             =   7920
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
   Begin MSDataListLib.DataCombo dcboKind 
      Bindings        =   "frmRequestDetails.frx":0000
      Height          =   360
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ListField       =   "Category"
      Text            =   ""
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
      Bindings        =   "frmRequestDetails.frx":0016
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ListField       =   "Type"
      Text            =   ""
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
   Begin VB.CommandButton cmdFindItem 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin MSAdodcLib.Adodc adoInventoryItems 
      Height          =   375
      Left            =   -240
      Top             =   7560
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
      Caption         =   "adoInventoryItems"
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
   Begin VB.CommandButton cmdCLose 
      Caption         =   "&Finish"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   9
      ToolTipText     =   "Close this form"
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddToRequest 
      Caption         =   "&Add To Request"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      ToolTipText     =   "Add the selected item to request"
      Top             =   6840
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid grdCurrentStocks 
      Bindings        =   "frmRequestDetails.frx":0036
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8454143
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Stock_ID"
         Caption         =   "Stock_ID"
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
         DataField       =   "Type"
         Caption         =   "Classification"
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
      BeginProperty Column02 
         DataField       =   "Category"
         Caption         =   "Kind"
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
      BeginProperty Column03 
         DataField       =   "Unit"
         Caption         =   "Unit"
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
      BeginProperty Column04 
         DataField       =   "QtyRemaining"
         Caption         =   "QtyRemaining"
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
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1500.095
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCriteria 
      DataField       =   "RequestDetail_ID"
      DataSource      =   "requestDetail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity:"
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
      Left            =   480
      TabIndex        =   11
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kind"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Classification"
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
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1260
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   735
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   -240
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   600
      Shape           =   3  'Circle
      Top             =   -120
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   840
      Top             =   -120
      Width           =   4095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   975
      Left            =   1200
      Shape           =   2  'Oval
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   7440
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   855
      Index           =   1
      Left            =   -360
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   2775
   End
End
Attribute VB_Name = "frmRequestDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loadStatus As Boolean
Private Sub adoClassification_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoKind_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoInventoryItems_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub adoRequestDetail_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
fCancelDisplay = True
End Sub

Private Sub cmdAddToRequest_Click()
Dim tempQty As Integer

On Error GoTo ErrHandler

If txtRequestQuantity.Text = "" Or Not IsNumeric(txtRequestQuantity.Text) Then
    MsgBox "Invalid Quantity."
    txtRequestQuantity.SetFocus
    Exit Sub
End If

If Val(txtRequestQuantity.Text) <= 0 Then
    MsgBox "Invalid Quantity."
    txtRequestQuantity.SetFocus
    Exit Sub
End If
If Val(txtRequestQuantity.Text) > adoInventoryItems.Recordset.Fields("QtyRemaining") Then
    MsgBox "The requested quantity you entered is greater that the " & vbCr & _
      "remaining quantity of the item you selected." & vbCr & "You can either enter a " & _
      "value greater than or equal to the remaining quantity.", vbInformation
    txtRequestQuantity.SetFocus
    Exit Sub
End If

If Not frmRequest.adoRequestItems.Recordset.EOF Then
frmRequest.adoRequestItems.Refresh
frmRequest.adoRequestItems.Recordset.MoveFirst
frmRequest.adoRequestItems.Recordset.Find "Stock_ID=" & adoInventoryItems.Recordset.Fields("Stock_ID")
End If

If frmRequest.adoRequestItems.Recordset.EOF Then
    Me.MousePointer = vbHourglass
    adoRequestDetail.Recordset.AddNew
    adoRequestDetail.Recordset.Fields("Rrequest_ID") = frmRequest.adoRequest.Recordset.Fields("Request_ID")
    adoRequestDetail.Recordset.Fields("Stock_ID") = adoInventoryItems.Recordset.Fields("Stock_ID")
    adoRequestDetail.Recordset.Fields("Quantity") = Val(txtRequestQuantity.Text)
    adoRequestDetail.Recordset.Update
    adoRequestDetail.Refresh
    frmRequest.adoRequestItems.Refresh
    txtRequestQuantity.Text = ""
    adoInventoryItems.Refresh
    Me.MousePointer = vbDefault
    MsgBox "Item successfully added."

ElseIf adoInventoryItems.Recordset.Fields("Stock_ID") = frmRequest.adoRequestItems.Recordset.Fields("Stock_ID") Then
    Me.MousePointer = vbHourglass
    adoRequestDetail.Refresh
    adoRequestDetail.Recordset.Find "RequestDetail_ID=" & frmRequest.adoRequestItems.Recordset.Fields("RequestDetail_ID")
    tempQty = adoRequestDetail.Recordset.Fields("Quantity")
    adoRequestDetail.Recordset.Fields("Quantity") = tempQty + Val(txtRequestQuantity.Text)
    adoRequestDetail.Recordset.Update
    adoRequestDetail.Refresh
    frmRequest.adoRequestItems.Refresh
    txtRequestQuantity.Text = ""
    adoInventoryItems.Refresh
    Me.MousePointer = vbDefault
    MsgBox "Item successfully added."
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdAddToRequest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddToRequest.FontBold = True
cmdViewAll.FontBold = False
cmdClose.FontBold = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddToRequest.FontBold = False
cmdViewAll.FontBold = False
cmdClose.FontBold = True
End Sub

Private Sub cmdFindItem_Click()
On Error GoTo ErrHandler

adoInventoryItems.CommandType = adCmdText
adoInventoryItems.RecordSource = "select * from qryInventory where Category like '" & txtCriteria.Text & "%' and QtyRemaining<>0"
adoInventoryItems.Refresh

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub cmdFindItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdFindItem.FontBold = True
End Sub

Private Sub cmdViewAll_Click()
On Error GoTo ErrHandler

Me.MousePointer = vbHourglass
adoInventoryItems.CommandType = adCmdText
adoInventoryItems.RecordSource = "select * from qryInventory where QtyRemaining<>0"
adoInventoryItems.Refresh
Me.MousePointer = vbDefault

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub
    
Private Sub Command1_Click()
'MsgBox adoInventoryItems.Recordset.Fields("Stock_ID")
'MsgBox frmRequest.adoRequest.Recordset.Fields("Request_ID")
End Sub

Private Sub cmdViewAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddToRequest.FontBold = False
cmdViewAll.FontBold = True
cmdClose.FontBold = False
End Sub

Private Sub dcboClassification_Change()
On Error GoTo ErrHandler

If loadStatus = False Then Exit Sub
adoClassification.Recordset.Bookmark = dcboClassification.SelectedItem
adoKind.CommandType = adCmdText
adoKind.RecordSource = "select * from tblstocktypecategory where StockType_ID=" & adoClassification.Recordset.Fields("StockType_ID") 'dcboClassification.SelectedItem
adoKind.Refresh
Set dcboKind.RowSource = adoKind.Recordset
dcboKind.ListField = "Category"
dcboKind.Text = ""
dcboKind.ReFill
adoInventoryItems.CommandType = adCmdText
adoInventoryItems.RecordSource = "select * from qryInventory where QtyRemaining<>0 and StockType_ID =" & adoClassification.Recordset.Fields("StockType_ID") 'dcboClassification.Text & "'"
adoInventoryItems.Refresh

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub dcboClassification_Click(Area As Integer)
loadStatus = True
End Sub

Private Sub dcboClassification_GotFocus()
loadStatus = True
End Sub

Private Sub dcboClassification_LostFocus()
loadStatus = False
End Sub

Private Sub dcboKind_Change()
On Error GoTo ErrHandler

If dcboKind.Text = "" Then Exit Sub
If dcboClassification.Text = "" Then 'Exit Sub
    adoKind.Recordset.Bookmark = dcboKind.SelectedItem
    adoInventoryItems.CommandType = adCmdText
    adoInventoryItems.RecordSource = "select * from qryInventory where StockTypeCategory_ID =" & adoKind.Recordset.Fields("StockTypeCategory_ID") 'dcboKind.Text & "'"
    adoInventoryItems.Refresh
Else
    adoKind.Recordset.Bookmark = dcboKind.SelectedItem
    adoInventoryItems.CommandType = adCmdText
    adoInventoryItems.RecordSource = "select * from qryInventory where QtyRemaining<>0 and StockType_ID = " & adoClassification.Recordset.Fields("StockType_ID") _
        & " and StockTypeCategory_ID =" & adoKind.Recordset.Fields("StockTypeCategory_ID")
    adoInventoryItems.Refresh
End If

Exit Sub
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub dcboKind_GotFocus()
loadStatus = True
End Sub

Private Sub dcboKind_LostFocus()
loadStatus = False
End Sub

Private Sub Form_Activate()
txtCriteria.SetFocus
End Sub

Private Sub Form_Deactivate()
Me.ZOrder (vbBringToFront)
Beep
Me.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

Call CenterForm(frmRequestDetails, MDIForm1)
Call ConnectDB(adoInventoryItems)
adoInventoryItems.CommandType = adCmdText
adoInventoryItems.RecordSource = "select * from qryInventory where QtyRemaining<>0"
adoInventoryItems.Refresh
Call ConnectDB(adoClassification)
adoClassification.CommandType = adCmdTable
adoClassification.RecordSource = "tblstocktype"
adoClassification.Refresh
Call ConnectDB(adoKind)
adoKind.CommandType = adCmdTable
adoKind.RecordSource = "tblstocktypecategory"
adoKind.Refresh
Call ConnectDB(adoRequestDetail)
adoRequestDetail.CommandType = adCmdTable
adoRequestDetail.RecordSource = "tblRequestDetail"
adoRequestDetail.Refresh

Exit Sub
ErrHandler:
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdAddToRequest.FontBold = False
cmdViewAll.FontBold = False
cmdClose.FontBold = False
cmdFindItem.FontBold = False
End Sub

Private Sub grdCurrentStocks_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub

Private Sub txtCriteria_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
    Call cmdFindItem_Click
End If
End Sub

Private Sub txtRequestQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 110 Or KeyCode = 190 Then SendKeys "{Backspace}"
End Sub

Private Sub txtRequestQuantity_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyDelete Then Exit Sub
If KeyAscii = vbKeyBack Then Exit Sub
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    KeyAscii = 0
    Exit Sub
End If
End Sub
