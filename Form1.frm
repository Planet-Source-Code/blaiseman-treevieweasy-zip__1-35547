VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Private Sub Form_Initialize()
'The next line creates a connect string, just like the one you created in the previous RDO example. In both cases, you are using ODBC’s "non-DSN" connection strategy to save time and increase performance:
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\tree.mdb"
    cn.Open
End Sub

Private Sub Form_Load()

    loadCategory
End Sub

Public Sub ExecuteSQL(strSQL As String)
    'This sub is for Insert, Update, and Delete SQL strings
    cn.Execute (strSQL)
End Sub

Public Function getData(strSQL As String) As ADODB.Recordset
    'instantiate recordset object and open
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, cn
    
    Set getData = rs
End Function

Private Sub Form_Terminate()
    Set rs = Nothing
    Set rs1 = Nothing
    Set cn = Nothing
End Sub

Public Function GetCategories() As Variant
    Dim strSQL As String
    strSQL = "SELECT * FROM category"
        Set rs = getData(strSQL)
        GetCategories = rs.GetRows()
End Function

Public Function GetNames() As Variant
    Dim strSQL As String
    strSQL = "SELECT * FROM name"
        Set rs1 = getData(strSQL)
        GetNames = rs1.GetRows()
End Function

Private Sub loadCategory()
    Dim arrCategory As Variant, x As Integer
    Dim arrName As Variant, y As Integer
        arrCategory = GetCategories
        arrName = GetNames
        
    Dim nodx As Node
    Set nodx = TreeView1.Nodes.Add(, , "root", "Category")
    
    For x = 0 To UBound(arrCategory, 2)
        Set nodx = TreeView1.Nodes.Add("root", tvwChild, arrCategory(1, x), arrCategory(1, x))
    Next x
    
    For y = 0 To UBound(arrName, 2)
        Set nodx = TreeView1.Nodes.Add(arrName(2, y), tvwChild, arrName(1, y), arrName(1, y))
    Next y
    
End Sub

Private Sub TreeView1_DblClick()
On Error Resume Next
    MsgBox TreeView1.SelectedItem.Text & " is a " & TreeView1.SelectedItem.Parent.Text
End Sub
