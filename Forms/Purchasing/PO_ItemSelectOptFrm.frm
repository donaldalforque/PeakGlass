VERSION 5.00
Begin VB.Form PO_ItemSelectOptFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cmbUnit 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   1
         Text            =   "1"
         Top             =   120
         Width           =   2535
      End
      Begin VB.CommandButton btnOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   810
      End
   End
End
Attribute VB_Name = "PO_ItemSelectOptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isFormSearch As Boolean
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    'Check if unit is empty
    If cmbUnit.text = "" Then
        MsgBox "Please select a unit of measure.", vbCritical
        cmbUnit.SetFocus
        Exit Sub
    End If

    Dim isExisting As Boolean
    isExisting = False

    If isModify = False Then
        Dim item As MSComctlLib.ListItem
        
        With PO_PurchaseOrderFrm
            For Each item In .lvItems.ListItems
                If item.SubItems(9) = .lvItemList.SelectedItem.text And item.SubItems(5) = cmbUnit.text Then
                    item.SubItems(4) = FormatNumber(Val(Replace(item.SubItems(4), ",", "")) + txtQuantity.text, 2, vbTrue, vbFalse)
                    item.SubItems(6) = txtCost.text
                    'item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
                    isExisting = True
                    Unload Me
                    .lvItemList.Visible = False
                    .txtItemSearch.SetFocus
                    .CountTotal
                    Exit Sub
                End If
            Next
        
            Set item = .lvItems.ListItems.add(, , "")
                item.SubItems(1) = ""
                item.SubItems(2) = .lvItemList.SelectedItem.SubItems(1) 'ItemCode
                item.SubItems(3) = .lvItemList.SelectedItem.SubItems(2) 'Name
                item.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
                item.SubItems(5) = cmbUnit.text '.lvItemList.SelectedItem.SubItems(4)
                item.SubItems(6) = FormatNumber(txtCost.text, 2, vbTrue, vbFalse)
                'item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
                item.SubItems(9) = .lvItemList.SelectedItem.text
                item.SubItems(12) = cmbUnit.ItemData(cmbUnit.ListIndex)
                Unload Me
                .lvItemList.Visible = False
                .txtItemSearch.SetFocus
                '.txtCode.SetFocus
                .CountTotal
        End With
    Else
        With PO_PurchaseOrderFrm
            .lvItems.SelectedItem.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(6) = FormatNumber(txtCost.text, 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(5) = cmbUnit.text
            Unload Me
            .lvItemList.Visible = False
            .txtItemSearch.SetFocus
            '.txtCode.SetFocus
            .CountTotal
        End With
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnOk_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
     'Get Uom Related
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_UomConversion_Get"
    
    If isModify = True Then
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , PO_PurchaseOrderFrm.lvItems.SelectedItem.SubItems(9))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , PO_PurchaseOrderFrm.lvItemList.SelectedItem.text)
    End If
    Set rec = cmd.Execute
    'lvUom.ListItems.Clear
    cmbUnit.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            'Set item = lvUom.ListItems.add(, , rec!UomConversionId)
                'item.SubItems(1) = rec!Uom
                cmbUnit.AddItem rec!Uom
                If IsNull(rec!toqty) = True Then
                    'item.SubItems(2) = "0.00"
                    cmbUnit.ItemData(cmbUnit.NewIndex) = 1
                Else
                    cmbUnit.ItemData(cmbUnit.NewIndex) = rec!toqty
                End If
'                If IsNull(rec!toqty) = True Then
'                    item.SubItems(3) = 1
'                Else
'                    item.SubItems(3) = FormatNumber(rec!toqty, 2, vbTrue, vbFalse)
'                End If
                rec.MoveNext
        Loop
    End If
    con.Close
    
    On Error Resume Next
    cmbUnit.ListIndex = 0
End Sub

Private Sub txtCost_Change()
    If IsNumeric(txtCost.text) = False Then
        txtCost.text = "1"
        selectText txtCost
    End If
End Sub

Private Sub txtCost_GotFocus()
    selectText txtCost
End Sub

Private Sub txtQuantity_Change()
    If IsNumeric(txtQuantity.text) = False Then
        txtQuantity.text = "1"
        selectText txtQuantity
    End If
End Sub

Private Sub txtQuantity_GotFocus()
    selectText txtQuantity
End Sub
