VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_RecentReceiptsFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recent Receipts"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "POS_RecentReceiptsFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "POS_RecentReceiptsFrm.frx":000C
      Left            =   1440
      List            =   "POS_RecentReceiptsFrm.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5880
      Width           =   4695
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox txtOrderNumber 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      TabIndex        =   4
      Top             =   6840
      Width           =   4695
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1440
      TabIndex        =   3
      Top             =   6360
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F2: View Details"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   2160
      Picture         =   "POS_RecentReceiptsFrm.frx":0047
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC:Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   4200
      Picture         =   "POS_RecentReceiptsFrm.frx":064A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   1935
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "F1:Print"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   120
      Picture         =   "POS_RecentReceiptsFrm.frx":29D9
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "POSSaleId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "OR #"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   465
      Left            =   1440
      TabIndex        =   1
      Top             =   4920
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96075777
      CurrentDate     =   41686
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   465
      Left            =   1440
      TabIndex        =   2
      Top             =   5400
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96075777
      CurrentDate     =   41686
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "TYPE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   5925
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6120
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "OR #"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   6840
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CUSTOMER"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DATE TO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATE FROM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   1185
   End
End
Attribute VB_Name = "POS_RecentReceiptsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = con
    cmd.CommandText = "SYSAuditTrail_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId) '1 DEFAULT
    cmd.Parameters.Append cmd.CreateParameter("@Module", adVarChar, adParamInput, 250, "POS")
    cmd.Parameters.Append cmd.CreateParameter("@Action", adVarChar, adParamInput, 250, "REPRINT")
    cmd.Execute
    con.Close
    
    'Save Audit Trail
    SavePOSAuditTrail UserId, WorkstationId, lvList.SelectedItem.text, "REPRINT OR#: " & lvList.SelectedItem.SubItems(1)
    
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    
    If cmbType.ListIndex = 0 Then
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt.rpt")
    ElseIf cmbType.ListIndex = 1 Then
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt_Account.rpt")
    ElseIf cmbType.ListIndex = 2 Then 'StockOut
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt_StockOut.rpt")
    ElseIf cmbType.ListIndex = 3 Then 'Return
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt_Return.rpt")
    End If
'    crxRpt.RecordSelectionFormula = "{POS_Sales.POS_SalesId}= " & lvList.SelectedItem.text & ""

    crxRpt.DiscardSavedData
    crxRpt.EnableParameterPrompting = False
    crxRpt.ParameterFields(1).AddCurrentValue "***THIS IS A REPRINT***"
    If cmbType.ListIndex = 0 Then
        crxRpt.ParameterFields.GetItemByName("@POS_SalesId").AddCurrentValue Val(lvList.SelectedItem.text)
    ElseIf cmbType.ListIndex = 1 Then
        crxRpt.ParameterFields.GetItemByName("@SalesOrderId").AddCurrentValue Val(lvList.SelectedItem.text)
        crxRpt.OpenSubreport("POS_ReceiptAccount_Payments.rpt").ParameterFields.GetItemByName("@SalesOrderId").AddCurrentValue Val(lvList.SelectedItem.text)
    ElseIf cmbType.ListIndex = 2 Then
        crxRpt.ParameterFields.GetItemByName("@POS_SalesId").AddCurrentValue Val(lvList.SelectedItem.text)
    ElseIf cmbType.ListIndex = 3 Then
        crxRpt.ParameterFields.GetItemByName("@POS_SalesId").AddCurrentValue Val(lvList.SelectedItem.text)
    End If
    Call ResetRptDB(crxRpt)
    crxRpt.PrintOut False
End Sub

Private Sub btnSearch_Click()
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    
    If cmbType.ListIndex = 0 Then
        cmd.CommandText = "POS_RecentReceipts_Search"
    ElseIf cmbType.ListIndex = 1 Then
        cmd.CommandText = "POS_RecentReceiptsAccounts_Search"
    ElseIf cmbType.ListIndex = 2 Then 'StockOut
        cmd.CommandText = "POS_RecentReceiptsStockOut_Search"
    ElseIf cmbType.ListIndex = 3 Then 'Return
        cmd.CommandText = "POS_RecentReceiptsReturn_Search"
    End If
    
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , dtFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , dtTo.value)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, txtname.text)
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtOrderNumber.text)
    Set rec = cmd.Execute
    lvList.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If cmbType.ListIndex = 0 Then
                Set item = lvList.ListItems.add(, , rec!POS_SalesId)
                    item.SubItems(1) = rec!pos_ordernumber
            ElseIf cmbType.ListIndex = 1 Then
                Set item = lvList.ListItems.add(, , rec!SalesOrderId)
                    item.SubItems(1) = rec!OrderNumber
            ElseIf cmbType.ListIndex = 2 Then
                Set item = lvList.ListItems.add(, , rec!POS_SalesId)
                    item.SubItems(1) = rec!pos_ordernumber
            ElseIf cmbType.ListIndex = 3 Then
                Set item = lvList.ListItems.add(, , rec!POS_SalesId)
                    item.SubItems(1) = rec!pos_ordernumber
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnPrint_Click
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    lvList.ColumnHeaders(2).width = lvList.width * 0.95
    cmbType.ListIndex = 0
    
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("RecentReceipts")
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvList.ListItems.add(, , rec!POS_SalesId)
                item.SubItems(1) = rec!pos_ordernumber
            rec.MoveNext
        Loop
    End If
    
    dtFrom.value = Format(Now, "MM/DD/YY")
    dtTo.value = Format(Now, "MM/DD/YY")
End Sub
