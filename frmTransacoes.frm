VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTransacoes 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Transações Financeiras"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUltimo 
      Caption         =   "Último"
      Height          =   255
      Left            =   13920
      TabIndex        =   26
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdProximo 
      Caption         =   "Próximo"
      Height          =   255
      Left            =   12480
      TabIndex        =   25
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "Anterior"
      Height          =   255
      Left            =   11040
      TabIndex        =   24
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdinicio 
      Caption         =   "Início"
      Height          =   255
      Left            =   9600
      TabIndex        =   23
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame fralocaliza 
      Caption         =   "Localizar"
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   15135
      Begin VB.TextBox txtlocaliza 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   22
         Top             =   160
         Width           =   4815
      End
      Begin VB.CommandButton cmdlocaliza 
         Caption         =   "Localizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11280
         TabIndex        =   21
         Top             =   160
         Width           =   3735
      End
      Begin VB.Label lbllocaliza 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   160
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin MSDataGridLib.DataGrid dtgTransacoes 
      Height          =   5295
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   9340
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Transações"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTransacoes 
      Caption         =   "Transação"
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   15135
      Begin VB.CommandButton cmdexportarexcel 
         Appearance      =   0  'Flat
         Caption         =   "Exportar Xls"
         Height          =   255
         Left            =   11880
         TabIndex        =   18
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdsair 
         Appearance      =   0  'Flat
         Caption         =   "Sair"
         Height          =   255
         Left            =   13200
         TabIndex        =   17
         Top             =   3480
         Width           =   735
      End
      Begin VB.ComboBox cmbstatus 
         DataField       =   "Status_Transacao"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmTransacoes.frx":0000
         Left            =   1920
         List            =   "frmTransacoes.frx":000D
         TabIndex        =   4
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox txtnumtransacao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Numero_Cartao"
         Enabled         =   0   'False
         Height          =   285
         Left            =   11040
         TabIndex        =   14
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton cmdExcluir 
         Appearance      =   0  'Flat
         Caption         =   "Exccluir"
         Height          =   255
         Left            =   10800
         TabIndex        =   7
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmdAlterar 
         Appearance      =   0  'Flat
         Caption         =   "Salvar"
         Height          =   255
         Left            =   9840
         TabIndex        =   6
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton cmdIncluir 
         Appearance      =   0  'Flat
         Caption         =   "Novo"
         Height          =   255
         Left            =   8640
         TabIndex        =   5
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtDescricao 
         Appearance      =   0  'Flat
         DataField       =   "Descricao"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1920
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2400
         Width           =   12015
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Data_Transacao"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   1800
         Width           =   2775
      End
      Begin VB.TextBox txtValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Valor_Transacao"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox txtNumcartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "Numero_Cartao"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   16
         TabIndex        =   0
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblnumtransacao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Número da Transação"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9000
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lblDescricao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Descrição"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Data"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblValor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Valor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblNumcartao 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Número do Cartão"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Label lblpagina 
      Caption         =   "Paginação"
      Height          =   255
      Left            =   7920
      TabIndex        =   27
      Top             =   6000
      Width           =   1455
   End
End
Attribute VB_Name = "frmTransacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conn As ADODB.Connection
Private WithEvents rs  As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private WithEvents subRst As ADODB.Recordset
Attribute subRst.VB_VarHelpID = -1
Const PAGE_SIZE = 20
Dim reg As Integer
Dim novaPagina As Long
Private intIdtransacao As Long

Private Sub cmbstatus_LostFocus()
   '
End Sub

Private Sub cmdAnterior_Click()
Navega (1)
End Sub

Private Sub cmdexportarexcel_Click()
Dim dataini As String
Dim datafim As String
rs.Filter = ""
dataini = "01/" & Format(DateAdd("m", -1, Now), "mm") & "/" & Year(DateAdd("m", -1, Now))
datafim = DateAdd("m", 1, dataini) - 1
rs.MoveFirst
rs.Filter = "Data_Transacao >= " & dataini & " and Data_Transacao <= " & datafim
If Not rs.EOF Then
    Call ExportRecordsetToExcel(rs, "C:\Gilberto\transacoes_exportados.xlsx")
End If
CarregarTransacoes
End Sub

Private Sub cmdinicio_Click()
Navega (0)
End Sub

Private Sub cmdlocaliza_Click()
    CarregarTransacoes
    Navega (0)
End Sub

Private Sub cmdProximo_Click()
Navega (2)
End Sub

Private Sub cmdsair_Click()
    Unload Me
End Sub

Private Sub Navega(Index As Integer)



With rs
Select Case Index
  Case 0
     novaPagina = 1
     reg = PAGE_SIZE
     
  Case 1
     novaPagina = novaPagina - 1
     reg = reg - PAGE_SIZE
     If reg < PAGE_SIZE Then
       reg = PAGE_SIZE
     End If
     
  Case 2
     novaPagina = .AbsolutePage + 1
     reg = reg + PAGE_SIZE
     If reg > rs.RecordCount Then
        reg = rs.RecordCount
     End If
     
  Case 3
     novaPagina = .PageCount
     reg = rs.RecordCount
     
End Select
lblpagina.Caption = PAGE_SIZE & " por Página"
If novaPagina < 1 Or novaPagina > .PageCount Then
   Exit Sub
End If

 .AbsolutePage = novaPagina
 Set dtgTransacoes.DataSource = PaginarRecordset(rs)
Configuracomponentes
End With
End Sub



Private Sub cmdUltimo_Click()
Navega (3)
End Sub

Private Sub dtgTransacoes_HeadClick(ByVal ColIndex As Integer)
    txtlocaliza.Text = ""
    fralocaliza.Visible = True
'    lbllocaliza.Visible = True
    cmdlocaliza.Visible = True
    txtlocaliza.Visible = True
    lbllocaliza.Caption = dtgTransacoes.Columns(ColIndex).Caption
    cmdlocaliza.Caption = "Localizar " & dtgTransacoes.Columns(ColIndex).Caption
'    lbllocaliza.Visible = False
'    fralocaliza.Visible = False
End Sub

Private Sub Form_Load()
    On Error GoTo trata_erro
    ' Inicializar conexão
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-IM0LA0V\SQLEXPRESS;Initial Catalog=GTX;Integrated Security=SSPI;"
    conn.Open
    Set rs = New ADODB.Recordset
    Set subRst = New ADODB.Recordset
    rs.Open "select *  from vw_ConsolidadaTransacoes where 1 = 999 ", conn, adOpenStatic, adLockReadOnly
    Set dtgTransacoes.DataSource = PaginarRecordset(rs)
    ' Carregar dados na grade
'    CarregarTransacoes
    
Exit Sub
trata_erro:
    MsgBox "Erro ao carregar formulário  " & Err.Description & " erro número " & Err.Number
End Sub
Private Sub Configuracomponentes()
If dtgTransacoes.Columns.Count > 2 Then
    dtgTransacoes.Columns("Valor_Transacao").NumberFormat = "0.00"
    dtgTransacoes.Columns("Valor_Transacao").Alignment = 3
    txtValor.Text = Format(txtValor.Text, "0.00")
    txtData.Text = Format(txtData.Text, "dd/mm/yyyy")
    dtgTransacoes.Columns(0).Visible = False
'    txtlocaliza.Visible = False
'    cmdlocaliza.Visible = False
'    fralocaliza.Visible = False
End If
End Sub
Private Sub CarregarTransacoes()
    On Error GoTo trata_erro
    Dim i As Long
    Dim strFiltro As String
    Set rs = New ADODB.Recordset
    Habilita_controles (True)
    rs.PageSize = PAGE_SIZE
    If txtlocaliza.Text <> "" And lbllocaliza.Caption <> "" Then
        If lbllocaliza.Caption <> "Valor_Transacao" And lbllocaliza.Caption <> "Data_Transacao" Then
            strFiltro = "" & lbllocaliza.Caption & " like  '%" & CStr(txtlocaliza.Text) & "%'"
         Else
             If lbllocaliza.Caption = "Valor_Transacao" Then
                strFiltro = "" & lbllocaliza.Caption & "  =" & Replace(txtlocaliza.Text, ",", ".")
             Else
                strFiltro = "" & lbllocaliza.Caption & " =  '" & CStr(txtlocaliza.Text) & "'"
            End If
        End If
    Else
        If intIdtransacao > 0 Then
            strFiltro = " Id_Transacao = " & intIdtransacao
        Else
            strFiltro = " 1 = 1"
        End If
    End If
    rs.Open "select *  from vw_ConsolidadaTransacoes where " & strFiltro, conn, adOpenStatic, adLockReadOnly
    Set subRst = New ADODB.Recordset
    Set dtgTransacoes.DataSource = PaginarRecordset(rs)
'    Set dtgTransacoes.DataSource = rs
    If rs.RecordCount = 0 Then
        Habilita_controles (False)
    End If
    
    
    Habilita_comandos
    Configuracomponentes
   
    Exit Sub
trata_erro:
    MsgBox "Erro ao carregar transações  " & Err.Description & " erro número " & Err.Number
End Sub

Private Sub cmdIncluir_Click()
    On Error GoTo trata_erro
    If cmdIncluir.Caption = "Novo" Then
        Call Limpa_tela
        Habilita_controles (True)
        txtNumcartao.SetFocus
        cmdAlterar.Enabled = True
        cmdIncluir.Caption = "Cancelar"
    Else
        cmdIncluir.Caption = "Novo"
        CarregarTransacoes
    End If
    Exit Sub
trata_erro:
    MsgBox "Erro ao incluir transação " & Err.Description & " erro número " & Err.Number
End Sub
Private Sub Limpa_tela()
        On Error GoTo trata_erro
        txtnumtransacao.Text = ""
        txtNumcartao.Text = ""
        txtValor.Text = ""
        txtData.Text = ""
        txtDescricao.Text = ""
        cmbstatus.Text = ""
        Exit Sub
trata_erro:
    MsgBox "Erro ao limpa tela  " & Err.Description & " erro número " & Err.Number
End Sub
Private Sub Habilita_controles(bolHabilita As Boolean)
    On Error GoTo trata_erro
    Dim i As Long
    For i = 0 To Me.Controls.Count - 1
        If TypeName(Me.Controls(i)) = "ComboBox" Or TypeName(Me.Controls(i)) = "TextBox" Then
            Me.Controls(i).Enabled = bolHabilita
        End If
    Next i
    cmdIncluir.Enabled = True
    txtlocaliza.Enabled = True
    txtnumtransacao.Enabled = False
    Exit Sub
trata_erro:
    MsgBox "Erro ao habilitar controles  " & Err.Description & " erro número " & Err.Number
End Sub
Private Function Valida_dados() As Boolean
        On Error GoTo trata_erro
        If Not IsNumeric(txtNumcartao.Text) Then
            MsgBox ("Número de cartão não numérico")
            txtNumcartao.SetFocus
            Exit Function
        End If
        If Len(txtNumcartao.Text) <> 16 Then
            MsgBox ("Número de cartão tem que ter 16 dígitos")
            txtNumcartao.SetFocus
            Exit Function
        End If
        If Not IsNumeric(txtValor.Text) Then
            MsgBox ("Valor não numérico")
            txtValor.SetFocus
            Exit Function
        End If
        If Not IsDate(txtData.Text) Then
            MsgBox ("Data inválida")
            txtData.SetFocus
            Exit Function
        End If
        If (txtDescricao.Text) = "" Then
            MsgBox ("Descrição inválida")
            txtDescricao.SetFocus
            Exit Function
        End If
        If (cmbstatus.Text) = "" Then
            MsgBox ("Staus inválido")
            cmbstatus.SetFocus
            Exit Function
        End If
       
        Valida_dados = True
Exit Function
trata_erro:
    MsgBox "Erro ao validar dados  " & Err.Description & " erro número " & Err.Number
End Function
Private Sub Salvar()
    On Error GoTo trata_erro
    Dim sql As String
    
    Dim bolincluiu As Boolean
   
    If Valida_dados Then
        If txtnumtransacao.Text = "" Then
            sql = "exec sp_InserirTransacao " & txtNumcartao.Text & "," & Replace(txtValor.Text, ",", ".") & ",'" & txtData.Text & "','" & txtDescricao.Text & "','" & cmbstatus.Text & "', ''"
            bolincluiu = True
         Else
            sql = "Exec sp_AlterarTransacao " & subRst("Id_Transacao").Value & "," & txtNumcartao.Text & "," & Replace(txtValor.Text, ",", ".") & ",'" & txtData.Text & "','" & txtDescricao.Text & "','" & cmbstatus.Text & "'"
        End If
        conn.Execute sql
        If subRst.BOF And subRst.EOF Then
            CarregarTransacoes
            Exit Sub
        End If
        intIdtransacao = subRst("Id_Transacao").Value
        CarregarTransacoes

        MsgBox "Transação salava com sucesso!"
        intIdtransacao = 0
        cmdIncluir.Caption = "Novo"
    End If
    Exit Sub
trata_erro:
    MsgBox "Erro ao salvar transaçao  " & Err.Description & " erro número " & Err.Number
End Sub
Private Sub cmdAlterar_Click()
    On Error GoTo trata_erro
    Call Salvar
    Exit Sub
trata_erro:
    MsgBox "Erro ao Salvar transação " & Err.Description & " erro número " & Err.Number
End Sub

Private Sub cmdExcluir_Click()
    On Error GoTo trata_erro
    If subRst.EOF Then Exit Sub
    Dim sql As String
    If InputBox("Deseja realmente excluir a transação selecionada", , "Sim") = "Sim" Then
        sql = "Exec sp_ExcluirTransacao " & subRst("Id_Transacao").Value
        conn.Execute sql
        CarregarTransacoes
        MsgBox "Transação excluída com sucesso!"
    End If
    If rs.EOF Then Limpa_tela
    Exit Sub
trata_erro:
    MsgBox "Erro ao excluir transação " & Err.Description & " erro número " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Fechar conexão
    On Error GoTo trata_erro
    If Not rs Is Nothing Then rs.Close
    Exit Sub
trata_erro:
    MsgBox "Erro ao sair do formulário  " & Err.Description & " erro número " & Err.Number
End Sub
Private Sub Habilita_comandos()
If rs.RecordCount = 0 Then
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdexportarexcel.Enabled = False
Else
    cmdExcluir.Enabled = True
    cmdAlterar.Enabled = True
    cmdexportarexcel.Enabled = True
End If
End Sub


Private Sub rs_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 ' Quando uma linha for selecionada, preencher os campos do formulário
    If Not rs.EOF Then
        txtnumtransacao.Text = rs!Id_Transacao
        txtNumcartao.Text = rs!Numero_Cartao
        txtValor.Text = rs!Valor_Transacao
        txtData.Text = rs!Data_Transacao
        txtDescricao.Text = rs!Descricao
        cmbstatus.Text = rs!Status_Transacao
        Habilita_alteracao_aprovada
    End If
    
End Sub
Private Sub Habilita_alteracao_aprovada()
        If cmbstatus.Text = "Aprovada" Then
            Habilita_controles (False)
        
        Else
            Habilita_controles (True)
            
        End If
End Sub

Public Sub ExportRecordsetToExcel(rs As ADODB.Recordset, ByVal CaminhoArquivo As String)
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    Dim i As Long
    Dim row As Long
    Dim fso As Object
    Dim pastaDestino As String

    On Error GoTo Erro

    ' Criar diretório, se não existir
    Set fso = CreateObject("Scripting.FileSystemObject")
    pastaDestino = fso.GetParentFolderName(CaminhoArquivo)
    If Not fso.FolderExists(pastaDestino) Then
        fso.CreateFolder (pastaDestino)
    End If

    ' Iniciar Excel
    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Sheets(1)

    ' Cabeçalhos formatados
    For i = 1 To rs.Fields.Count
        
        xlSheet.Cells(1, i).Value = rs.Fields(i - 1).Name
        xlSheet.Cells(1, i).Font.Bold = True
        If rs.Fields(i - 1).Name = "Valor_Transacao" Then
           xlSheet.Cells(1, i).HorizontalAlignment = 4
           xlSheet.Cells(1, i).NumberFormat = "0.00"
        End If
    Next i

    ' Dados do Recordset
    row = 2
    Do Until rs.EOF
        For i = 1 To rs.Fields.Count
            If rs.Fields(i - 1).Name = "Numero_Cartao" Then
                xlSheet.Cells(row, i).Value = "'" & CStr(rs.Fields(i - 1).Value)
            Else
            xlSheet.Cells(row, i).Value = CStr(rs.Fields(i - 1).Value)
            End If
            If rs.Fields(i - 1).Name = "Valor_Transacao" Then
                xlSheet.Cells(row, i).NumberFormat = "0.00"
                xlSheet.Cells(row, i).HorizontalAlignment = 4
            End If
        Next i
        rs.MoveNext
        row = row + 1
    Loop

    ' Ajuste de largura
    xlSheet.Columns.AutoFit

    ' Salvar
    xlWB.SaveAs CaminhoArquivo
    xlWB.Close False
    xlApp.Quit

    MsgBox "Planilha exportada com sucesso para: " & CaminhoArquivo, vbInformation

Saida:
    ' Liberar objetos
    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    Set fso = Nothing
    Exit Sub

Erro:
    MsgBox "Erro ao exportar para Excel: " & Err.Description, vbCritical
    Resume Saida
End Sub
Private Function PaginarRecordset(recset As ADODB.Recordset) As ADODB.Recordset

Dim x As Long
Dim fld As Fields
Dim origPage As Long
Dim i As Integer
origPage = IIf(recset.AbsolutePage > 0, recset.AbsolutePage, 1)

With subRst
  If .State = adStateOpen Then .Close

  'Cria campos
  For i = 0 To recset.Fields.Count - 1
     .Fields.Append recset.Fields(i).Name, recset.Fields(i).Type, recset.Fields(i).DefinedSize, recset.Fields(i).Attributes
  Next i

  'Inclui registros
  .Open
  For x = 1 To PAGE_SIZE
 
    If recset.EOF Then Exit For
      .AddNew
 
    
    For i = 0 To recset.Fields.Count - 1
        
            .Fields(i).Value = recset.Fields(i).Value
        
    Next i
    .Update
    recset.MoveNext

  Next x
If .RecordCount > 0 Then
    .MoveFirst
End If
If Not recset.EOF And Not recset.BOF Then
    recset.AbsolutePage = origPage
End If
End With

Set PaginarRecordset = subRst

End Function

Private Sub subRst_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
' Quando uma linha for selecionada, preencher os campos do formulário
    If Not subRst.EOF Then
        txtnumtransacao.Text = subRst!Id_Transacao
        txtNumcartao.Text = "" & subRst!Numero_Cartao
        txtValor.Text = "" & subRst!Valor_Transacao
        txtData.Text = "" & subRst!Data_Transacao
        txtDescricao.Text = "" & subRst!Descricao
        cmbstatus.Text = "" & subRst!Status_Transacao
        Habilita_alteracao_aprovada
        Configuracomponentes
    End If
    
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    ' Permite apenas números e Backspace
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtData_Change()
    Dim txt As String
    txt = txtData.Text

    ' Remove barras antigas
    txt = Replace(txt, "/", "")

    ' Aplica formatação "dd/mm/yyyy"
    If Len(txt) >= 2 Then
        txt = Left(txt, 2) & "/" & Mid(txt, 3)
    End If
    If Len(txt) >= 5 Then
        txt = Left(txt, 5) & "/" & Mid(txt, 6)
    End If

    ' Evita loop infinito alterando texto dentro do evento
    If txt <> txtData.Text Then
        txtData.Text = txt
        txtData.SelStart = Len(txt)
    End If
End Sub
