VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form INV_UomPricingFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "INV_UomPricingFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbUom 
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
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5400
      Width           =   3495
   End
   Begin VB.CheckBox chkShow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4050
      TabIndex        =   1
      Top             =   1830
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox txtToQTY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   3840
      MaxLength       =   50
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvConversion 
      Height          =   3135
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "LocationId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Unit"
         Object.Width           =   6253
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   1588
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_UomPricingFrm.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_UomPricingFrm.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_UomPricingFrm.frx":D0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_UomPricingFrm.frx":13932
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "INV_UomPricingFrm.frx":1A194
      Top             =   680
      Width           =   480
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Sell your products with flexibility, add prices for each type of product units."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   240
      TabIndex        =   5
      Top             =   1250
      Width           =   4815
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Prices"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5415
      Left            =   120
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "INV_UomPricingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim ConversionId As Long
Public Sub Initialize()
    'cmbUom.text = ""
    ConversionId = 0
    cmbUom.SetFocus
End Sub
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "ConversionLoad"
            Dim con As New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_UomConversion_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewProductFrm.ProductId)
            Set rec = cmd.Execute
            lvConversion.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                     Set item = lvConversion.ListItems.add(, , "")
                            item.SubItems(1) = rec!UomConversionId
                            item.SubItems(2) = rec!Uom
                            If IsNull(rec!price) Then
                                item.SubItems(3) = "0.00"
                            Else
                                item.SubItems(3) = FormatNumber(rec!price, 2, vbTrue, vbFalse)
                            End If
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "Uom"
            Set rec = Global_Data("Uom")
            cmbUom.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbUom.AddItem rec!Uom
                        cmbUom.ItemData(cmbUom.NewIndex) = rec!UomId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbUom.ListIndex = 0
    End Select
End Sub

Private Sub chkShow_Click()
    Dim item As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Tax")
    lvConversion.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If chkShow.value = 1 Then
                Set item = lvConversion.ListItems.add(, , "")
                    item.SubItems(1) = rec!ConversionId
                    item.SubItems(2) = rec!TaxName
                    item.SubItems(3) = rec!Percentage
                If rec!isActive = "True" Then item.Checked = True
                lvConversion.ColumnHeaders(1).width = lvConversion.width * 0.06
                lvConversion.ColumnHeaders(3).width = lvConversion.width * 0.44
                lvConversion.ColumnHeaders(4).width = lvConversion.width * 0.44
            Else
                If rec!isActive = "True" Then
                    Set item = lvConversion.ListItems.add(, , "")
                        item.SubItems(1) = rec!ConversionId
                        item.SubItems(2) = rec!TaxName
                        item.SubItems(3) = rec!Percentage
                    If rec!isActive = "True" Then item.Checked = True
                    lvConversion.ColumnHeaders(1).width = lvConversion.width * 0
                    lvConversion.ColumnHeaders(3).width = lvConversion.width * 0.47
                    lvConversion.ColumnHeaders(4).width = lvConversion.width * 0.47
                End If
            End If
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Load()
    lvConversion.ColumnHeaders(1).width = lvConversion.width * 0
    lvConversion.ColumnHeaders(3).width = lvConversion.width * 0.47
    lvConversion.ColumnHeaders(4).width = lvConversion.width * 0.47
    Populate "Uom"
    Populate "ConversionLoad"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ConversionId = 0
End Sub

Private Sub lvConversion_ItemClick(ByVal item As MSComctlLib.ListItem)
    ConversionId = item.SubItems(1)
    On Error Resume Next
    cmbUom.text = item.SubItems(2)
    txtToQTY.text = item.SubItems(3)
    cmbUom.SetFocus
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo ErrorHandler:
    Select Case Button.Index
        Case 1 'NEW
            Initialize
        Case 2 'Save
            Dim item As MSComctlLib.ListItem
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            
           
            
            con.ConnectionString = ConnString
            con.Open
            
'            'Check for Deactivate/Activated Lists
'            For Each item In lvConversion.ListItems
'                Set cmd = New ADODB.Command
'                cmd.ActiveConnection = con
'                cmd.CommandType = adCmdStoredProc
'                cmd.CommandText = "BASE_Tax_Update"
'                cmd.Parameters.Append cmd.CreateParameter("@ConversionId", adInteger, adParamInputOutput, , item.SubItems(1))
'                cmd.Parameters.Append cmd.CreateParameter("@TaxName", adVarChar, adParamInput, 250, item.SubItems(2))
'                cmd.Parameters.Append cmd.CreateParameter("@Percentage", adInteger, adParamInput, , item.SubItems(3))
'                cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , item.Checked)
'                cmd.Execute
'            Next
            
            If Trim(cmbUom.text) = "" Then
                Exit Sub
            End If
        
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@UomConversionId", adInteger, adParamInputOutput, , ConversionId)
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewProductFrm.ProductId)
            cmd.Parameters.Append cmd.CreateParameter("@UomId", adInteger, adParamInput, , cmbUom.ItemData(cmbUom.ListIndex))
            If ConversionId = 0 Then
                cmd.Parameters.Append cmd.CreateParameter("@ToQty", adDecimal, adParamInput, , 1)
                                  cmd.Parameters("@ToQty").NumericScale = 2
                                  cmd.Parameters("@ToQty").Precision = 18
            Else
                cmd.Parameters.Append cmd.CreateParameter("@ToQty", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@ToQty").NumericScale = 2
                                  cmd.Parameters("@ToQty").Precision = 18
            End If
            cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , Val(Replace(txtToQTY.text, ",", "")))
                                  cmd.Parameters("@Price").NumericScale = 2
                                  cmd.Parameters("@Price").Precision = 18
            
            If ConversionId = 0 Then
                cmd.CommandText = "INV_UomConversion_Insert"
                cmd.Execute
                ConversionId = cmd.Parameters("@UomConversionId")
'                If ConversionId <> 0 Then
'                    Set item = lvConversion.ListItems.add(, , "")
'                        item.SubItems(1) = ConversionId
'                        item.SubItems(2) = cmbUom.text
'                        item.SubItems(3) = FormatNumber(Val(txtToQTY.text), 2, vbTrue, vbFalse)
'    '                    item.Checked = True
'                        item.Selected = True
'                        item.EnsureVisible
'                End If
                Populate "ConversionLoad"
            Else
                cmd.CommandText = "INV_UomConversion_Update"
                cmd.Execute
                For Each item In lvConversion.ListItems
                    If item.SubItems(1) = ConversionId Then
                        item.SubItems(2) = cmbUom.text
                        item.SubItems(3) = FormatNumber(Val(txtToQTY.text), 2, vbTrue, vbFalse)
                        item.Selected = True
                        item.EnsureVisible
                    End If
                Next
            End If
            con.Close
        Case 4 'Delete
            Dim x As Variant
            x = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo)
            If x = vbYes Then
                If lvConversion.ListItems.Count > 0 Then
                    'Dim item As MSComctlLib.ListItem
                    Set con = New ADODB.Connection
                    Set cmd = New ADODB.Command
                    
                    con.ConnectionString = ConnString
                    con.Open
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "INV_UomConversion_Delete"
                    cmd.Parameters.Append cmd.CreateParameter("@UomConversionId", adInteger, adParamInput, , lvConversion.SelectedItem.SubItems(1))
                    cmd.Execute
                    con.Close
                    
                    lvConversion.ListItems.Remove (lvConversion.SelectedItem.Index)
                    Initialize
                End If
            End If
    End Select
    Exit Sub
'ErrorHandler:
'    If IsNumeric(Err.Description) = True Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
'    Else
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
'    End If
'    GLOBAL_MessageFrm.Show (1)
End Sub

Private Sub txtToQty_Change()
    If IsNumeric(txtToQTY.text) = False Then
        txtToQTY.text = 0
    End If
End Sub

Private Sub cmbUom_GotFocus()
'    selectText cmbUom
End Sub





