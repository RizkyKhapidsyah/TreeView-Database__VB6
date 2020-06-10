VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Demo TreeView Database"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   540
      Width           =   1995
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2640
      TabIndex        =   1
      Top             =   1260
      Width           =   1995
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   2040
      Width           =   1995
   End
   Begin MSComctlLib.TreeView TVW 
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   7435
      _Version        =   393217
      Indentation     =   176
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   255
      Left            =   2700
      TabIndex        =   6
      Top             =   300
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   255
      Left            =   2700
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Atasan ID"
      Height          =   255
      Left            =   2700
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim Rst As New ADODB.Recordset

Dim Rst1 As New ADODB.Recordset
Dim Rst2 As New ADODB.Recordset
Dim Rst3 As New ADODB.Recordset
Dim Rst4 As New ADODB.Recordset
Dim Rst5 As New ADODB.Recordset

Sub FillTree()
    TVW.Nodes.Clear
    
    TVW.Nodes.Add(, , "M", "MASTER").Expanded = True 'level 0
    
    Rst1.Open "select * from tblStruktur where AtasanID ='M' order by Nama", Con
    
    Do Until Rst1.EOF 'karakter "k" ditambahkan supaya tidak error
        TVW.Nodes.Add("M", tvwChild, "k" & Rst1!ID, Rst1!Nama).Expanded = True 'level 1

 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Rst2.State = 1 Then Rst2.Close
        Rst2.Open "select * from tblStruktur where AtasanID ='" & Rst1!ID & "' order by Nama", Con
        
        Do Until Rst2.EOF
            TVW.Nodes.Add "k" & Rst1!ID, tvwChild, "k" & Rst2!ID, Rst2!Nama 'level 2
                        
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If Rst3.State = 1 Then Rst3.Close
            Rst3.Open "select * from tblStruktur where AtasanID ='" & Rst2!ID & "' order by Nama", Con
    
            Do Until Rst3.EOF
                TVW.Nodes.Add "k" & Rst2!ID, tvwChild, "k" & Rst3!ID, Rst3!Nama 'level 3
                
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Rst4.State = 1 Then Rst4.Close
                Rst4.Open "select * from tblStruktur where AtasanID ='" & Rst3!ID & "' order by Nama", Con
    
                Do Until Rst4.EOF
                    TVW.Nodes.Add "k" & Rst3!ID, tvwChild, "k" & Rst4!ID, Rst4!Nama 'level 4
                    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If Rst5.State = 1 Then Rst5.Close
                    Rst5.Open "select * from tblStruktur where AtasanID ='" & Rst4!ID & "' order by Nama", Con
                    
                    Do Until Rst5.EOF
                        TVW.Nodes.Add "k" & Rst4!ID, tvwChild, "k" & Rst5!ID, Rst5!Nama 'level 5
                        Rst5.MoveNext
                    Loop
                    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Rst4.MoveNext
                Loop
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Rst3.MoveNext
            Loop
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Rst2.MoveNext
        Loop
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Rst1.MoveNext
    Loop
    
    'close connection
    Set Rst1 = Nothing
    Set Rst2 = Nothing
    Set Rst3 = Nothing
    Set Rst4 = Nothing
    Set Rst5 = Nothing
End Sub

Private Sub cmdKeluar_Click()
    End
End Sub

Private Sub Form_Load()
    Con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\struktur.mdb;"

    FillTree
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Con.Close
End Sub

Private Sub TVW_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim DataID As String
    
    DataID = Right(Node.Key, Len(Node.Key) - 1) 'menghapus karakter "k"
    
    Rst.Open "select * from tblStruktur where ID ='" & DataID & "'", Con, adOpenKeyset
    Text1 = Rst!ID
    Text2 = Rst!Nama
    Text3 = Rst!AtasanID
    Set Rst = Nothing
End Sub

