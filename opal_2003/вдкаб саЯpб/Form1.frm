VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "opal just relax"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "print"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "search"
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid mf1 
      Height          =   2535
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   285
      BackColor       =   -2147483624
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "name"
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************
       'I love u opal

'*******************************
Private Sub Command1_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String

'*******************************
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=;data source=" & App.Path & "\db1.mdb"
cn.Open

mf1.Rows = 1
mf1.Cols = 4
'*******************************

sql = "SELECT * FROM table1 Where "

If Text1.Text <> "" Then
sql = sql & "sname LIKE '" & Text1.Text & "' AND "
End If



sql = Left(sql, Len(sql) - 4)
rs.Open sql, cn
 While Not rs.EOF
     
      mf1.AddItem mf1.Rows - 1
      mf1.TextMatrix(mf1.Rows - 1, 1) = rs("sname")
       mf1.TextMatrix(mf1.Rows - 1, 2) = rs("fname")
       mf1.TextMatrix(mf1.Rows - 1, 3) = rs("gname")
       
       rs.MoveNext
    Wend

If rs.EOF = True Then
rs.Close
End If
sql = "delete * FROM table2 "
rs.Open sql, cn, adOpenDynamic, adLockOptimistic
sql = "select * FROM table2 "
rs.Open sql, cn, adOpenDynamic, adLockOptimistic

For i = 1 To Val((mf1.Rows) - 1)
rs.AddNew
rs!tno = mf1.TextMatrix(i, 0)
rs!sname = mf1.TextMatrix(i, 1)
rs!fname = mf1.TextMatrix(i, 2)
rs!gname = mf1.TextMatrix(i, 3)
rs.Update
Next i


Exit Sub
End Sub
Private Sub Command2_Click()
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Set cn = New ADODB.Connection
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\" & "db1.mdb;" & "Jet OLEDB:Database Password=;"
Set rs = New ADODB.Recordset
Set rs = cn.Execute("select * from table2")
Set DataReport1.DataSource = rs
DataReport1.Show 1
cn.Close
End Sub

Private Sub Form_Load()
mf1.ColAlignment(0) = 4
mf1.ColWidth(0) = 700
mf1.TextMatrix(0, 0) = "no"
mf1.ColAlignment(1) = 4
mf1.TextMatrix(0, 1) = "name"
mf1.ColAlignment(2) = 4
mf1.TextMatrix(0, 2) = "father"
mf1.ColAlignment(3) = 4
mf1.TextMatrix(0, 3) = "grand"

End Sub
