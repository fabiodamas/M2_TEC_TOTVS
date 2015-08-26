VERSION 5.00
Object = "{02F125F5-49EE-11D5-A561-0050BF395743}#1.0#0"; "OcxControl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmFiExportacaoTotvs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportação Totvs"
   ClientHeight    =   8025
   ClientLeft      =   2025
   ClientTop       =   1515
   ClientWidth     =   16845
   Icon            =   "FrmFiExportacaoTotvs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   16845
   Begin VB.Frame FraTipo 
      Caption         =   "Tipo de Exportação"
      Height          =   735
      Left            =   100
      TabIndex        =   9
      Top             =   0
      Width           =   2940
      Begin VB.OptionButton optTipo 
         Caption         =   "Pagamento"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   320
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Recebimento"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   320
         Width           =   1320
      End
   End
   Begin VB.Frame FraIntervalo 
      Caption         =   "Data de vencimento"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   2490
      Begin MSComDlg.CommonDialog Dia 
         Left            =   -495
         Top             =   -270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin OcxControl.Text TxtDataInicial 
         Height          =   330
         Left            =   90
         TabIndex        =   0
         Top             =   250
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         BorderStyle     =   0
         Mask            =   "99/99/9999"
         MaxLength       =   0
         Text            =   "08/12/2013"
         PromptChar      =   "_"
         CampoTipo       =   2
      End
      Begin OcxControl.Text TxtDataFinal 
         Height          =   330
         Left            =   1350
         TabIndex        =   1
         Top             =   250
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         BorderStyle     =   0
         Mask            =   "99/99/9999"
         MaxLength       =   0
         Text            =   "08/12/2013"
         PromptChar      =   "_"
         CampoTipo       =   2
      End
      Begin VB.Label LblIntervaloA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "à"
         Height          =   195
         Left            =   1160
         TabIndex        =   14
         Top             =   250
         Width           =   195
      End
   End
   Begin VB.CommandButton CmdMontar 
      Caption         =   "Montar"
      Height          =   675
      Left            =   14295
      Picture         =   "FrmFiExportacaoTotvs.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   75
      Width           =   735
   End
   Begin VB.CommandButton CmdLimpar 
      Caption         =   "&Limpar"
      Height          =   675
      Left            =   15960
      Picture         =   "FrmFiExportacaoTotvs.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Limpar"
      Top             =   75
      Width           =   735
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   15135
      Picture         =   "FrmFiExportacaoTotvs.frx":3A06
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   75
      Width           =   735
   End
   Begin VB.ListBox lstLancamentos 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      ItemData        =   "FrmFiExportacaoTotvs.frx":4228
      Left            =   12000
      List            =   "FrmFiExportacaoTotvs.frx":422A
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtLancamentos 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   100
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   960
      Width           =   16695
   End
   Begin VB.TextBox txtErros 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   100
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   5520
      Width           =   16695
   End
   Begin MSComDlg.CommonDialog dlgArquivo 
      Left            =   11520
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lancamentos:0 | Parcelas:0 | Erros:0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   100
      TabIndex        =   12
      Top             =   7680
      Width           =   16695
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lançamentos"
      Height          =   195
      Left            =   100
      TabIndex        =   11
      Top             =   730
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Erros"
      Height          =   195
      Left            =   100
      TabIndex        =   10
      Top             =   5280
      Width           =   360
   End
End
Attribute VB_Name = "FrmFiExportacaoTotvs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLimpar_Click()
    On Error GoTo Erro
   
    'Limpando dados do formulário
    TxtDataInicial.text = "__/__/____"
    TxtDataFinal.text = "__/__/____"
    
    lblTotal.Caption = "Lancamentos:0 | Parcelas:0 | Erros:0"

    
    TxtDataInicial.SetFocus
    txtErros.text = ""
    
    'Limpa os itens
    lstLancamentos.Clear
    'lstLancamentosVisualizacao.Clear
    txtLancamentos.text = ""
    
    cmdExportar.Enabled = False
   
    Exit Sub
Erro:
   TratarErro Me.Name, "Executado", Err.Number, Err.Description, Erl
End Sub

Private Sub CmdCancelar1_Click()
End Sub

Private Sub Exporta()
On Error GoTo Erro
    'Caminho do arquivo a ser exportado
    Dim Caminho As String
    
    'Variável para For
    Dim iContador As Integer

    'Definindo o nome do arquivo dependendo do tipo de exportação
    If (optTipo(0).value = True) Then
        dlgArquivo.FileName = "Pagamento-" & Format(Date, "DD-MM-YYYY") & ".TXT"
    Else
        dlgArquivo.FileName = "Recebimento-" & Format(Date, "DD-MM-YYYY") & ".TXT"
    End If
    
    'Configuração básicas para salvar
    dlgArquivo.InitDir = App.path
    dlgArquivo.Filter = "Text (*.txt)|*.txt"
    dlgArquivo.CancelError = True
    dlgArquivo.ShowSave
    
    'Caminho escolhido
    Caminho = dlgArquivo.FileName
    
    'Se o arquivo existir, apagamos
    If Dir(Caminho) = dlgArquivo.FileTitle Then
        Kill Caminho
    End If
    
    'Criamos o arquivo no caminho especificado
    Open Caminho For Output As #1
    
    'Gravando os registro baseado nos itens da lista
    For iContador = 0 To lstLancamentos.ListCount
        DoEvents
        Print #1, lstLancamentos.List(iContador)
    Next
    
    'Fechando o arquivo
    Close #1
    
    MsgBox "Arquivo '" & Caminho & "' exportado com sucesso.", vbInformation

Exit Sub
Erro:

    'Caso o usuário clicar no botão cancelar, não terão tenhuma ação
    Select Case Err
    Case 32755
        Exit Sub
    End Select

   Close #1
   TratarErro Me.Name, "Executado", Err.Number, Err.Description, Erl
End Sub
Private Sub CmdMontar_Click()
    On Error GoTo Erro
    
    'Se os dados são validos
    If ValidaDados Then
        'Preenchemos o ListBox com a lista de pagamentos e baixas
        MontaLancamentoPagamentoRecebimento
    End If
    
Exit Sub
Erro:
   TratarErro Me.Name, "Executado", Err.Number, Err.Description, Erl
End Sub

Private Sub CmdExportar_Click()
    On Error GoTo Erro
   
    If txtErros.text <> "" Then
        If MsgBox("Existem erros na exportação, deseja continuar?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
            Exit Sub
        End If
    End If
   
    'Procedimento para exportar
    Exporta

    Exit Sub
Erro:
   TratarErro Me.Name, "Executado", Err.Number, Err.Description, Erl
End Sub

Private Sub Form_Load()
   On Error GoTo Erro
   
   PosicionaForm Me, 1
   
   
   
   Exit Sub
Erro:
   TratarErro Me.Name, "Executado", Err.Number, Err.Description, Erl
End Sub
Private Function ValidaDados() As Boolean
    On Error GoTo Erro
    
    'Verificação da data inicial
    If Not IsDate(Format(TxtDataInicial.text, "YYYY/MM/DD")) Then
       MsgBox "Data Inicial inválida!", vbInformation
       TxtDataInicial.SetFocus
       Exit Function
    End If
    
    'Verificação da data final
    If Not IsDate(Format(TxtDataFinal.text, "YYYY/MM/DD")) Then
       MsgBox "Data Final inválida!", vbInformation
       TxtDataInicial.SetFocus
       Exit Function
    End If
    
    'Data final deve ser maior que data inicial
    If DateDiff("d", Format(TxtDataInicial.text, "YYYY/MM/DD"), Format(TxtDataFinal.text, "YYYY/MM/DD")) < 0 Then
       MsgBox "A data inicial deve ser menor ou igual que a data final da exportação.", vbInformation
       TxtDataInicial.SetFocus
       Exit Function
    End If
    
    ValidaDados = True
    
    Exit Function
Erro:
   TratarErro Me.Name, "Executado", Err.Number, Err.Description, Erl
End Function

Private Sub MontaLancamentoPagamentoRecebimento()
    'Armazena cada registro vindo da Stored Procedure
    Dim sRegistro As String
   
    'Verificação da existência de stored procedures no banco
    Dim sSql As String
    
    'Registro da quantidade de lançamentos, parcelas e erros
    Dim iContaLancamento As Integer
    Dim iContaParcela As Integer
    Dim iErros As Integer
    
    'Contador do For
    Dim iCont As Integer
    
    'Linha da exportação
    Dim sLinha As String
    
    'Armazenagem dos campos a serem exportados
    Dim sDocumento As String, sEspecie As String, sSerie As String, sParcela As String, sCodigoParcela As String, _
    sDataEmissao As String, sDataVencimento As String, sValor As String, sCodigoFornecedor As String, _
    sCodigoCliente As String
   
    'Usada armazenar a chamada da stored procedure
    Dim objCmd As ADODB.Command
    
    'Usado para armazenar o resultado da stored procedure
    Dim objRst As ADODB.Recordset
    
    'Definição do objeto
    Set objCmd = New ADODB.Command
    Set objRst = New ADODB.Recordset
    
    'Conexão
    objRst.ActiveConnection = SV
    
    'Zerando estatísticas
    iContaLancamento = 0
    iContaParcela = 0
    
    If optTipo(0).value = True Then 'pagamento
        'Procurando a stored procedure
        sSql = ""
        sSql = sSql & " IF OBJECT_ID('SP_CONTA_PAGAR') IS NULL "
        sSql = sSql & " BEGIN "
        sSql = sSql & "   SELECT 'FALTA_PROCEDURE' as msgerror  "
        sSql = sSql & " End "
        objRst.Source = sSql
        objRst.Open
        
        'Procura o nome do campo gerado
        For iCont = 0 To objRst.Fields.Count - 1
            If objRst.Fields(iCont).Name = "msgerror" Then
               MsgBox "Stored Procedure não encontrada na base de dados" & Chr(13) & "Contate o serviço de suporte"
               Exit Sub
            End If
        Next
    
        'Definição da stored procedure a ser executada
        objCmd.CommandText = "SP_CONTA_PAGAR"
    ElseIf optTipo(1).value = True Then 'recebimento
    
        'Procurando a stored procedure
        sSql = ""
        sSql = sSql & " IF OBJECT_ID('SP_CONTA_RECEBER') IS NULL "
        sSql = sSql & " BEGIN "
        sSql = sSql & "   SELECT 'FALTA_PROCEDURE' as msgerror  "
        sSql = sSql & " End "
        objRst.Source = sSql
        objRst.Open
        
        'Procura o nome do campo gerado
        For iCont = 0 To objRst.Fields.Count - 1
            If objRst.Fields(iCont).Name = "msgerror" Then
               MsgBox "Stored Procedure não encontrada na base de dados" & Chr(13) & "Contate o serviço de suporte"
               Exit Sub
            End If
        Next
    
        'Definição da stored procedure a ser executada
        objCmd.CommandText = "SP_CONTA_RECEBER"
        
    End If
    
    Set objCmd.ActiveConnection = SV
    
    'Limpa os itens e zera erros
    lstLancamentos.Clear
    txtLancamentos.text = ""
    txtErros.text = ""
    iErros = 0
    
    
    'Processando
    lblTotal.Caption = "Processando..."
    
    'Definição do tipo de chamado
    objCmd.CommandType = adCmdStoredProc
    
    SetAmpulheta
        
    'Criação dos parametros
    objCmd.Parameters.Append objCmd.CreateParameter("dataInicial", adVarChar, adParamInput, 10, Format(TxtDataInicial.text, "YYYY/MM/DD"))
    objCmd.Parameters.Append objCmd.CreateParameter("dataFinal", adVarChar, adParamInput, 10, Format(TxtDataFinal.text, "YYYY/MM/DD"))
    
    DoEvents
    
    'Execução e retorno dos dados
    Set objRst = objCmd.Execute
    
    If optTipo(0).value = True Then 'pagamento
        txtLancamentos.text = txtLancamentos.text & "Lançamento  Data        Valor" & Chr(13) & Chr(10)
        txtLancamentos.text = txtLancamentos.text & "    Fornecedor  Espec.  Serie Documento        Parcela Emissão    Vencimento Valor " & Chr(13) & Chr(10)
    Else
        txtLancamentos.text = txtLancamentos.text & "Lançamento  Data        Valor" & Chr(13) & Chr(10)
        txtLancamentos.text = txtLancamentos.text & "    Fornecedor Espec.  Serie    Documento  Parcela Vencimento Valor " & Chr(13) & Chr(10)
    End If
                                                     
    CmdMontar.Enabled = False
    
    Do While (Not objRst Is Nothing)
    
        If objRst.state = adStateClosed Then
            MsgBox "Não há lançamentos para o período selecionado!"
            Exit Do
        End If
        
        
        While Not objRst.EOF
            
           'Procura o nome do campo gerado
            For iCont = 0 To objRst.Fields.Count - 1
                sRegistro = objRst(iCont)
            Next
            
            If Mid(sRegistro, 1, 3) = "100" Then
                iContaLancamento = iContaLancamento + 1
            ElseIf Mid(sRegistro, 1, 3) = "200" Then
                iContaParcela = iContaParcela + 1
            End If
            
            If optTipo(0).value = True Then 'pagamento
                'Definição da stored procedure a ser executada
                If Mid(sRegistro, 1, 3) = "100" Then
                    sLinha = ""
                    sLinha = Trim((Mid(sRegistro, 7, 10))) + Space(10 - Len(Trim((Mid(sRegistro, 7, 10))))) & "  " & _
                    Mid(sRegistro, 17, 2) & "/" & Mid(sRegistro, 19, 2) & "/" & Mid(sRegistro, 21, 4) & "  " & _
                    CLng(Mid(sRegistro, 25, 12)) & "," & Format(CLng(Mid(sRegistro, 37, 2)), "00")
                    
                    'lstLancamentosVisualizacao.AddItem sLinha
                    txtLancamentos.text = txtLancamentos.text & sLinha & Chr(13) & Chr(10)
                ElseIf Mid(sRegistro, 1, 3) = "200" Then
                    sLinha = "    "
                    
                    sCodigoFornecedor = Trim((Mid(sRegistro, 4, 9))) + Space(9 - Len(Trim((Mid(sRegistro, 4, 9)))))
                    sDocumento = Trim((Mid(sRegistro, 428, 16))) + Space(16 - Len(Trim((Mid(sRegistro, 428, 16))))) & " "
                    sEspecie = Trim((Mid(sRegistro, 13, 3))) + Space(3 - Len(Trim((Mid(sRegistro, 13, 3))))) & " "
                    sSerie = Trim((Mid(sRegistro, 16, 3))) + Space(3 - Len(Trim((Mid(sRegistro, 16, 3))))) & "   "
                    sParcela = Trim((Mid(sRegistro, 29, 2))) + Space(2 - Len(Trim((Mid(sRegistro, 29, 2))))) & "      "
                    sDataEmissao = Mid(sRegistro, 31, 2) & "/" & Mid(sRegistro, 33, 2) & "/" & Mid(sRegistro, 35, 4) & " "
                    sDataVencimento = Mid(sRegistro, 39, 2) & "/" & Mid(sRegistro, 41, 2) & "/" & Mid(sRegistro, 43, 4) & " "
                    sValor = CLng(Mid(sRegistro, 65, 9)) & "," & Format(CLng(Mid(sRegistro, 74, 2)), "00")
                    
                    If Trim(sDocumento) = "" Then
                        sDocumento = "Documento não Informado."
                    Else
                        sDocumento = "Documento: '" & sDocumento & "'. "
                    End If
                    
                    'código do fornecedor
                    sLinha = sLinha & sCodigoFornecedor
                    
                    'Validação do código do fornecedor
                    If Not IsNumeric(sCodigoFornecedor) Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Código do Fornecedor inválido." & Chr(13) & Chr(10)
                    Else
                        If Val(sLinha) = 0 Then
                            iErros = iErros + 1
                            txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "- Código do Fornecedor Totvs não informado. Informe o código no Cadastro de Fornecedores." & Chr(13) & Chr(10)
                        End If
                    End If
                    
                    'Espécie
                    sLinha = sLinha & "   " & sEspecie
                    
                    'Validação da Espécie
                    If Trim(sEspecie) = "" Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Espécie não informada." & Chr(13) & Chr(10)
                    End If
                    
                    
                    'serie
                    sLinha = sLinha & "    " & sSerie
                    
                    'Validação da série
                    If Trim(sSerie) = "" Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Série não informada." & Chr(13) & Chr(10)
                    End If
                    
                    'documento
                    sLinha = sLinha & Trim((Mid(sRegistro, 428, 16))) + Space(16 - Len(Trim((Mid(sRegistro, 428, 16))))) & " "
                    
                    'Validação do documento
                    If Trim(Trim((Mid(sRegistro, 428, 16))) + Space(16 - Len(Trim((Mid(sRegistro, 428, 16)))))) = "" Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & sDocumento & Chr(13) & Chr(10)
                    End If
                    
                    'codigo parcela
                    sLinha = sLinha & sParcela
                    
                    'Validação do código da parcela
                    If Trim(sParcela) = "" Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Código da Parcela não informado." & Chr(13) & Chr(10)
                    End If
                    
                    ' data emissao
                    sLinha = sLinha & sDataEmissao
                    
                    'Validação da data de emissão
                    If Not IsDate(sDataEmissao) Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Data de emissão inválida." & Chr(13) & Chr(10)
                    End If
                    
                    ' data de vencimento
                    sLinha = sLinha & sDataVencimento
                    
                    'Validação da data de vencimento
                    If Not IsDate(sDataVencimento) Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Data de vencimento inválida." & Chr(13) & Chr(10)
                    End If
                    
                    'valor
                    sLinha = sLinha & sValor
                    
                    'Validação do valor
                    If Not IsNumeric(sValor) Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Valor do documento inválido." & Chr(13) & Chr(10)
                    End If
                
                    'lstLancamentosVisualizacao.AddItem sLinha
                    txtLancamentos.text = txtLancamentos.text & sLinha & Chr(13) & Chr(10)

                End If
                
            ElseIf optTipo(1).value = True Then 'recebimento
                'Definição da stored procedure a ser executada
                If Mid(sRegistro, 1, 3) = "100" Then
                    sLinha = ""
                    sLinha = Trim((Mid(sRegistro, 7, 10))) + Space(10 - Len(Trim((Mid(sRegistro, 7, 10))))) & "  " & _
                    Mid(sRegistro, 17, 2) & "/" & Mid(sRegistro, 19, 2) & "/" & Mid(sRegistro, 21, 4) & "  " & _
                    CLng(Mid(sRegistro, 26, 11)) & "," & Format(CLng(Mid(sRegistro, 37, 2)), "00")
                    
                    'lstLancamentosVisualizacao.AddItem sLinha
                    txtLancamentos.text = txtLancamentos.text & sLinha & Chr(13) & Chr(10)
                ElseIf Mid(sRegistro, 1, 3) = "200" Then
                    sLinha = "    "
                    
                    sCodigoCliente = Trim((Mid(sRegistro, 5, 9))) + Space(9 - Len(Trim((Mid(sRegistro, 5, 9)))))
                    sDocumento = Trim((Mid(sRegistro, 20, 10))) + Space(10 - Len(Trim((Mid(sRegistro, 20, 10))))) & " "
                    sEspecie = Trim((Mid(sRegistro, 14, 3))) + Space(3 - Len(Trim((Mid(sRegistro, 14, 3))))) & " "
                    sSerie = Trim((Mid(sRegistro, 17, 3))) + Space(3 - Len(Trim((Mid(sRegistro, 17, 3))))) & "          "
                    sParcela = Trim((Mid(sRegistro, 38, 2))) + Space(2 - Len(Trim((Mid(sRegistro, 38, 2))))) & "      "
                    sDataVencimento = Mid(sRegistro, 48, 2) & "/" & Mid(sRegistro, 50, 2) & "/" & Mid(sRegistro, 52, 4) & " "
                    sValor = CLng(Mid(sRegistro, 75, 9)) & "," & Format(CLng(Mid(sRegistro, 83, 2)), "00")
                    
                    If Trim(sDocumento) = "" Then
                        sDocumento = "Documento não Informado."
                    Else
                        sDocumento = "Documento: '" & sDocumento & "'. "
                    End If
                    
                    'código do cliente
                    sLinha = sLinha & sCodigoCliente
                    
                    'Validação do código do fornecedor
                    If Not IsNumeric(sCodigoCliente) Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Código do Cliente inválido." & Chr(13) & Chr(10)
                    Else
                        If Val(sLinha) = 0 Then
                            iErros = iErros + 1
                            txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Código do Cliente é obrigatório." & Chr(13) & Chr(10)
                        End If
                    End If
                    
                    'especie
                    sLinha = sLinha & "  " & sEspecie
                    'Validação da Espécie
                    If Trim(sEspecie) = "" Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Espécie não informada." & Chr(13) & Chr(10)
                    End If
                    
                    'serie
                    sLinha = sLinha & sSerie
                    
                    'Validação da série
                    'If Trim(sSerie) = "" Then
                    ' iErros = iErros + 1
                    '     txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Série não informada." & Chr(13) & Chr(10)
                    ' End If
                    
                    'documento
                    sLinha = sLinha & Trim((Mid(sRegistro, 20, 10))) + Space(10 - Len(Trim((Mid(sRegistro, 20, 10))))) & " "
                    
                    'Validação do documento
                    If Trim((Mid(sRegistro, 20, 10))) + Space(10 - Len(Trim((Mid(sRegistro, 20, 10))))) = "" Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & sDocumento & Chr(13) & Chr(10)
                    End If
                   
                    
                    'codigo parcela
                    sLinha = sLinha & sParcela
                    
                    'Validação do código da parcela
                    If Trim(sParcela) = "" Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Código da Parcela não informado." & Chr(13) & Chr(10)
                    End If
                    
                    ' data de vencimento
                    sLinha = sLinha & sDataVencimento
                    
                    'Validação da data de vencimento
                    If Not IsDate(sDataVencimento) Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Data de vencimento inválida." & Chr(13) & Chr(10)
                    End If
                    
                    'valor
                    sLinha = sLinha & sValor
                    
                    'Validação do valor
                    If Not IsNumeric(sValor) Then
                        iErros = iErros + 1
                        txtErros.text = txtErros.text & "Erro: " & iErros & ". " & sDocumento & "Valor do documento inválido." & Chr(13) & Chr(10)
                    End If
                
                    txtLancamentos.text = txtLancamentos.text & sLinha & Chr(13) & Chr(10)
                End If
                    
            End If
            
            'Adiciona o lançamento no listbox
            lstLancamentos.AddItem sRegistro
           
            'Proximo registro
            objRst.MoveNext
            
            sRegistro = ""
        Wend
        
        txtLancamentos.SelStart = Len(txtLancamentos.text)
        txtLancamentos.SelLength = 0
        
        txtErros.SelStart = Len(txtErros.text)
        txtErros.SelLength = 0
        
        txtLancamentos.Refresh
        txtErros.Refresh
        'A primeira vez é excutado para o conjunto de registro mestre
        'Logo após é executado para o conjunto de registro detalhe
        Set objRst = objRst.NextRecordset
    Loop
    
    CmdMontar.Enabled = True
    reSetAmpulheta
    
    If lstLancamentos.ListCount = 0 Then
        'Exibindo o total de lançamentos
        lblTotal.Caption = "Lancamentos:0 | Parcelas:0 | Erros:0"
    Else
        lblTotal.Caption = "Lancamentos:" & iContaLancamento & " | Parcelas:" & iContaParcela & " | Erros:" & iErros
    End If
    
    'Apaga o recordset
    Set objRst = Nothing
    
   
   'Caso encontrarmos lançamentos, permitiremos a exportação
    If lstLancamentos.ListCount > 0 Then
       cmdExportar.Enabled = True
    Else
       cmdExportar.Enabled = False
    End If
    
    Exit Sub
Erro:
   CmdMontar.Enabled = True
   reSetAmpulheta
   TratarErro Me.Name, "Executado", Err.Number, Err.Description, Erl
End Sub

