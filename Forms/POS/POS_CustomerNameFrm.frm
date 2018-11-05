VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form POS_CustomerNameFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "POS_CustomerNameFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSets 
      Caption         =   "F2: SETS"
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
      Left            =   2280
      Picture         =   "POS_CustomerNameFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton btnNewCustomer 
      Caption         =   "F1: Add Customer"
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
      Picture         =   "POS_CustomerNameFrm.frx":279B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton btnReturn 
      Caption         =   "F4"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtCustomer 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   6255
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC: Cancel"
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
      Left            =   5520
      Picture         =   "POS_CustomerNameFrm.frx":2DB3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Accept"
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
      Left            =   3840
      Picture         =   "POS_CustomerNameFrm.frx":5142
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CustomerId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_CustomerNameFrm.frx":7516
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customer:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5415
      Left            =   120
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "POS_CustomerNameFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub LoadCustomersOnPOS()
    Dim item As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Customer")
    If Not rec.EOF Then
        lvList.ListItems.Clear
        Do Until rec.EOF
            Set item = lvList.ListItems.add(, , rec!CustomerId)
                item.SubItems(1) = rec!CustomerCode
                item.SubItems(2) = rec!Name
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub btnAccept_Click()
    If lvList.ListItems.Count > 0 Then
        POS_CashierFrm.lblCustomer.Caption = "| CUSTOMER: " & lvList.SelectedItem.SubItems(2)
        POS_CashierFrm.POSCustomerId = lvList.SelectedItem.text
        Unload Me
    Else
        Dim x As Variant
        x = MsgBox(txtCustomer.text & " is not registered. Would you like to register this customer?", vbQuestion + vbYesNo, "Customer not found.")
        If x = vbYes Then
            POS_AddCustomerFrm.txtName.text = txtCustomer.text
            POS_AddCustomerFrm.txtName.SelStart = Len(POS_AddCustomerFrm.txtName.text)
            POS_AddCustomerFrm.Show (1)
        Else
            txtCustomer.SelStart = 0
            txtCustomer.SelLength = Len(txtCustomer.text)
        End If
    End If
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnNewCustomer_Click()
    POS_AddCustomerFrm.txtName.text = txtCustomer.text
    POS_AddCustomerFrm.txtName.SelStart = Len(POS_AddCustomerFrm.txtName.text)
    POS_AddCustomerFrm.Show (1)
End Sub

Private Sub btnReturn_Click()
    txtCustomer.SelStart = 0
    txtCustomer.SelLength = Len(txtCustomer.text)
    txtCustomer.SetFocus
End Sub

Private Sub Command1_Click()

End Sub

Private Sub btnSets_Click()
    BASE_SetSelectionFrm.Show (1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        btnAccept_Click
    Case vbKeyEscape
        btnCancel_Click
    Case vbKeyF4
        btnReturn_Click
    Case vbKeyF1
        btnNewCustomer_Click
    Case vbKeyF2
        btnSets_Click
End Select
End Sub

Private Sub Form_Load()
    lvList.ColumnHeaders(2).width = (lvList.width * 0.22)  'Customer
    lvList.ColumnHeaders(3).width = (lvList.width * 0.73)
    LoadCustomersOnPOS
End Sub

Public Sub txtCustomer_Change()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Customer_Search"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtCustomer.text)
    cmd.Parameters.Append cmd.CreateParameter("@CustomerCode", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.Append cmd.CreateParameter("@isActive", adInteger, adParamInput, , Null)

    Set rec = cmd.Execute
    lvList.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                Set item = lvList.ListItems.add(, , rec!CustomerId)
                    item.SubItems(1) = rec!CustomerCode
                    item.SubItems(2) = rec!Name
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvList.ListItems.Count > 0 Then
                lvList.SetFocus
            End If
    End Select
End Sub
