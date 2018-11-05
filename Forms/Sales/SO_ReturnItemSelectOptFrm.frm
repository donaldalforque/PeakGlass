VERSION 5.00
Begin VB.Form SO_ReturnItemSelectOptFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CheckBox chkUnitCost 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Use Unit Cost"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4095
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
         Top             =   1560
         Width           =   1215
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
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cmbLocation 
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
         Left            =   -9999
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtPrice 
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
         Left            =   1560
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   840
         Width           =   2415
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
         Left            =   1560
         TabIndex        =   0
         Text            =   "1"
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Location"
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
         Left            =   -99999
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         TabIndex        =   10
         Top             =   840
         Width           =   465
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
         TabIndex        =   9
         Top             =   120
         Width           =   810
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
   End
End
Attribute VB_Name = "SO_ReturnItemSelectOptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isFormSearch As Boolean
Private Sub btnCancel_Click()
    Unload Me
    Set SO_ItemSelectOptionFrm = Nothing
End Sub
Private Sub GetPrice()
    On Error Resume Next
    txtPrice.text = FormatNumber(SO_SalesReturnFrm.lvItemList.SelectedItem.SubItems(3), 2, vbTrue, vbFalse)
    'GET ACTUAL PRICE
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_UomConversion_Get"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(SO_SalesReturnFrm.lvItemList.SelectedItem.text))
    Set rec = cmd.Execute
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!Uom = cmbUnit.text Then
                txtPrice.text = FormatNumber(rec!price, 2, vbTrue, vbFalse)
                Exit Sub
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
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
    Dim Available As Double
    Dim ReserveId As String
    
    With SO_SalesReturnFrm
        Dim item As MSComctlLib.ListItem
        If isModify = True Then
            'GetInventorySettings
'            If AllowNegativeInventory = False Then
'                'CHECK AVAILABLE QUANTITY
'                Available = CheckAvailableQuantity(.lvItems.SelectedItem.SubItems(9))
'
'                If Available + (Val(Replace(.lvItems.SelectedItem.SubItems(4), ",", "")) * Val(Replace(.lvItems.SelectedItem.SubItems(13), ",", ""))) < (Val(Replace(txtQuantity.text, ",", "")) * cmbUnit.ItemData(cmbUnit.ListIndex)) Then  'less the current
'                    MsgBox "Insufficient quantity.", vbCritical, "Error!"
'                    Exit Sub
'                Else
'                    'UDPATE RESERVES
'                    ReserveId = ReserveProduct(.lvItems.SelectedItem.SubItems(14), .lvItems.SelectedItem.text, (cmbUnit.ItemData(cmbUnit.ListIndex) * Val(Replace(txtQuantity.text, ",", ""))), UserId, False)
'                End If
'            End If
            
            .lvItems.SelectedItem.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(5) = cmbUnit.text
            .lvItems.SelectedItem.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
            .lvItems.SelectedItem.SubItems(13) = cmbUnit.ItemData(cmbUnit.ListIndex) 'actual inventory
            .lvItems.SelectedItem.SubItems(14) = ReserveId
            .CountTotal
            isModify = False
            Unload Me
            Exit Sub
        End If
        
        For Each item In .lvItems.ListItems
            If item.SubItems(9) = .lvItemList.SelectedItem.text And item.SubItems(5) = cmbUnit.text Then
                'GetInventorySettings
'                If AllowNegativeInventory = False Then
'                    'CHECK AVAILABLE QUANTITY
'                    Available = CheckAvailableQuantity(.lvItems.SelectedItem.SubItems(9))
'
'                    If Available + (Val(Replace(item.SubItems(4), ",", "")) * Val(Replace(item.SubItems(13), ",", ""))) < _
'                    (NVAL(item.SubItems(4)) + NVAL(txtQuantity.text)) * cmbUnit.ItemData(cmbUnit.ListIndex) Then
'                        MsgBox "Insufficient quantity.", vbCritical, "Insufficient Quantity"
'                        Exit Sub
'                    Else
'                        'UDPATE RESERVES
'                        ReserveId = ReserveProduct(item.SubItems(14), .lvItemList.SelectedItem.text, cmbUnit.ItemData(cmbUnit.ListIndex) * Val(Replace(txtQuantity.text, ",", "")) + (Val(Replace(item.SubItems(4), ",", "")) * Val(Replace(item.SubItems(13), ",", ""))), UserId, False)
'                    End If
'                End If
            
                item.SubItems(4) = FormatNumber(Val(Replace(item.SubItems(4), ",", "")) + txtQuantity.text, 2, vbTrue, vbFalse)
                item.SubItems(6) = txtPrice.text
                'item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
                item.SubItems(14) = ReserveId
                isExisting = True
                Exit For
            End If
        Next
  
        If isExisting = False Then
            If isFormSearch = False Then
                'GetInventorySettings
'                If AllowNegativeInventory = False Then
'                    'CHECK AVAILABLE QUANTITY
'                    Available = CheckAvailableQuantity(.lvItemList.SelectedItem.text)
'
'                    If Available < Val(Replace(txtQuantity.text, ",", "")) * cmbUnit.ItemData(cmbUnit.ListIndex) Then
'                        MsgBox "Insufficient quantity.", vbCritical, "Insufficient Quantity"
'                        Exit Sub
'                    Else
'                        'INSERT RESERVES
'                        ReserveId = ReserveProduct(0, .lvItemList.SelectedItem.text, cmbUnit.ItemData(cmbUnit.ListIndex) * Val(Replace(txtQuantity.text, ",", "")), UserId, False)
'                    End If
'                End If
                
                Set item = .lvItems.ListItems.add(, , "")
                item.SubItems(1) = ""
                item.SubItems(2) = .lvItemList.SelectedItem.SubItems(1) 'ItemCode
                item.SubItems(3) = .lvItemList.SelectedItem.SubItems(2) 'Name
                item.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
                item.SubItems(5) = cmbUnit.text '.lvItemList.SelectedItem.SubItems(4)
                item.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
                item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
                item.SubItems(9) = .lvItemList.SelectedItem.text
                item.SubItems(13) = cmbUnit.ItemData(cmbUnit.ListIndex)
                item.SubItems(14) = ReserveId
            Else
'                Set item = .lvItems.ListItems.add(, , "")
'                item.SubItems(1) = ""
'                item.SubItems(2) = INV_ProductSearch.lvItemList.SelectedItem.SubItems(1) 'ItemCode
'                item.SubItems(3) = INV_ProductSearch.lvItemList.SelectedItem.SubItems(2) 'Name
'                item.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
'                item.SubItems(5) = cmbUnit.text '.lvItemList.SelectedItem.SubItems(4)
'                item.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
'                item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
'                item.SubItems(9) = INV_ProductSearch.lvItemList.SelectedItem.text
            End If
        End If
            
        Unload Me
        Unload INV_ProductSearch
        .lvItemList.Visible = False
        .txtItemSearch.SetFocus
        '.txtCode.SetFocus
        .lvItems.ListItems(.lvItems.ListItems.Count).Selected = True
        .lvItems.ListItems(.lvItems.ListItems.Count).EnsureVisible
        .CountTotal
    End With
End Sub

Private Sub chkUnitCost_Click()
    If chkUnitCost.value = vbChecked Then
        txtPrice.text = FormatNumber(SO_SalesReturnFrm.lvItemList.SelectedItem.SubItems(5), 2, vbTrue, vbFalse)
    Else
        txtPrice.text = FormatNumber(SO_SalesReturnFrm.lvItemList.SelectedItem.SubItems(3), 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub cmbUnit_Click()
    GetPrice
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            btnCancel_Click
        Case vbKeyReturn
            btnOk_Click
    End Select
End Sub

Private Sub Form_Load()
    txtQuantity.text = "1.00"
    selectText txtQuantity
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Location")
    cmbLocation.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                cmbLocation.AddItem rec!Location
                cmbLocation.ItemData(cmbLocation.NewIndex) = rec!LocationId
                
            End If
            rec.MoveNext
        Loop
    End If
    cmbLocation.ListIndex = 1
    
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
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , SO_SalesReturnFrm.lvItems.SelectedItem.SubItems(9))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , SO_SalesReturnFrm.lvItemList.SelectedItem.text)
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
                    cmbUnit.ItemData(cmbUnit.NewIndex) = 0
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
    GetPrice
End Sub

Private Sub Form_Unload(Cancel As Integer)
    isModify = False
End Sub

Private Sub txtPrice_Change()
    If IsNumeric(txtPrice.text) = False Then txtPrice.text = "0.00"
End Sub

Private Sub txtPrice_GotFocus()
    selectText txtPrice
End Sub

Private Sub txtQuantity_Change()
    If IsNumeric(txtQuantity.text) = False Then
        txtQuantity.text = "1"
    End If
End Sub

Private Sub txtQuantity_GotFocus()
    selectText txtQuantity
End Sub


