VERSION 5.00
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "SmartMenuXP.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm MDISave 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Save Hospitalar"
   ClientHeight    =   8205
   ClientLeft      =   1470
   ClientTop       =   1335
   ClientWidth     =   12435
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picRelatorio 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      Height          =   930
      Left            =   0
      Picture         =   "MDI.frx":0CCA
      ScaleHeight     =   870
      ScaleWidth      =   12375
      TabIndex        =   10
      Top             =   375
      Visible         =   0   'False
      Width           =   12435
   End
   Begin VB.PictureBox picBackdrop 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   1875
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   12375
      TabIndex        =   7
      Top             =   1305
      Visible         =   0   'False
      Width           =   12435
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   315
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   9
         Top             =   315
         Width           =   4095
      End
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7650
         Left            =   855
         Picture         =   "MDI.frx":D9BF
         ScaleHeight     =   7650
         ScaleWidth      =   12000
         TabIndex        =   8
         Top             =   135
         Width           =   12000
      End
   End
   Begin VB.PictureBox picBarra 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   12435
      TabIndex        =   0
      Top             =   7920
      Width           =   12435
      Begin VB.Frame FraAviso 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   11100
         TabIndex        =   3
         Top             =   -45
         Visible         =   0   'False
         Width           =   825
         Begin VB.CommandButton CmdAviso 
            BackColor       =   &H00FEBD9A&
            Caption         =   "AVISO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   45
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5100
         Begin VB.CommandButton CmdLicenca 
            BackColor       =   &H0080FFFF&
            Caption         =   "Troque a senha da licença: faltam 5 dia(s)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   4830
         End
      End
      Begin ComctlLib.StatusBar StatusBar 
         Height          =   285
         Left            =   60
         TabIndex        =   5
         Top             =   0
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   503
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   2
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Alignment       =   1
               Object.Width           =   7055
               MinWidth        =   7055
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   6
               Alignment       =   1
               Object.Width           =   1764
               MinWidth        =   1764
               TextSave        =   "20/12/2013"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label LblF2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Caption         =   "Pressione F8 para a requisição"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7875
         TabIndex        =   6
         Top             =   45
         Visible         =   0   'False
         Width           =   3165
      End
   End
   Begin VB.Timer Tim_Licenca 
      Interval        =   3000
      Left            =   855
      Top             =   0
   End
   Begin VBSmartXPMenu.SmartMenuXP MenuXP 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   15982021
      CheckBackColor  =   0
      SelBackColor    =   12907007
      SeparatorColor  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OffsetLeft      =   3
      OffsetTop       =   0
      OffsetRight     =   3
      OffsetBottom    =   0
      Shadow          =   0   'False
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1290
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "MDISave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim QuantidadeHora As Long
Public nKey As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Sub CmdLicenca_Click()
   FrmLicenca.Show vbModal
End Sub

Private Sub MDIForm_Load()
On Error GoTo Erro
Dim Impe As String, Resultado As String
Dim Indice As Long
Dim fsot As New FileSystemObject
Dim fsot1
Dim VersaoExecutavelLocal As String
      'VerificaSenha ("aaa")
      
      MDISave.MenuXP.MenuItems.Add 0, "", smiNone, " ", , , , , True, True
      Modulo = 1
1:    SetAmpulheta
2:    If App.PrevInstance Then
3:       MsgBox "Já existe uma cópia do aplicativo sendo executado !", vbInformation
4:       End
5:       Exit Sub
6:    End If
      
      rptPath = App.path

7:    If OpenRegKey(HKEY_CURRENT_USER, "System\Save", nKey) Then
         
8:       RegQueryStringValue nKey, "Layout", LayouRegEdit
9:       RegQueryStringValue nKey, "Impressora", Impe
10:      RegQueryStringValue nKey, "ImpressoraFarmacia", ImpressoraFarmacia
11:      RegQueryStringValue nKey, "DriverImpressao", DriverImpressao

         RegQueryStringValue nKey, "CentroCustoPrescricaoLocal", CentroCustoPrescricaoLocal

         RecuperaPadraoImpressora
         
12:      If Trim(ImpressoraFarmacia) = "" Then
13:         RegSetValueEx nKey, "ImpressoraFarmacia", 0, 1, ByVal "", 0
14:      End If
15:      RegQueryStringValue nKey, "Fonte5", Fonte5
16:      RegQueryStringValue nKey, "Fonte6", Fonte6
17:      RegQueryStringValue nKey, "Fonte10", Fonte10
18:      RegQueryStringValue nKey, "Fonte12", Fonte12
19:      RegQueryStringValue nKey, "Fonte17", Fonte17
20:      RegQueryStringValue nKey, "Fonte20", Fonte20
         
21:      CarregaVetorImpressora
22:      For Indice = 1 To TotalImpressoraMatricial
            Resultado = ""
23:         RegQueryStringValue nKey, VetImpressoraMatricial(Indice).nome, Resultado
24:         VetImpressoraMatricial(Indice).Caminho = Resultado
25:      Next
         
         RegQueryStringValue nKey, "PadraoImpressaoEtiquetaPrescricaoIndividual", PadraoImpressaoEtiquetaPrescricaoIndividual
         RegQueryStringValue nKey, "PadraoImpressaoEtiquetaPrescricaoEmbutidas", PadraoImpressaoEtiquetaPrescricaoEmbutidas
         PortaImpressaoEtiquetaPrescricaoIndividual = VetImpressoraMatricial(8).Caminho
         PortaImpressaoEtiquetaPrescricaoEmbutidas = VetImpressoraMatricial(9).Caminho
         
         RegQueryStringValue nKey, "PortaImpressaoRecepcaoEtiquetaInternos", PortaImpressaoRecepcaoEtiquetaInternos
         
         PortaImpressaoLaudoETIQUETA = VetImpressoraMatricial(5).Caminho
         RegQueryStringValue nKey, "PadraoImpressoraLaudoETIQUETA", PadraoImpressoraLaudoETIQUETA
         PortaImpressaoLaudoREQUISICAO = VetImpressoraMatricial(6).Caminho
         RegQueryStringValue nKey, "PadraoImpressoraLaudoREQUISICAO", PadraoImpressoraLaudoREQUISICAO
         PortaImpressaoLaudoEXAME = VetImpressoraMatricial(7).Caminho
         RegQueryStringValue nKey, "PadraoImpressoraLaudoEXAME", PadraoImpressoraLaudoEXAME
         
         RegQueryStringValue nKey, "PortaImpressaoProdutoFarmacia", PortaImpressaoProdutoFarmacia
         PortaImpressaoProdutoFarmacia = VetImpressoraMatricial(10).Caminho
         
         'PortaImpressaoProdutoFarmacia
         
         RegQueryStringValue nKey, "IPInterfaceamento", IPInterfaceamento
         If Trim(IPInterfaceamento) <> "" Then
9875:       Interfaceamento_Conecta "SISTEX", "sistex"
         End If
         
26:      ImpressoraMatricial = Val(Impe)
27:   Else
28:      RegCreateKey HKEY_CURRENT_USER, "System\Save", nKey
29:      RegSetValueEx nKey, "Impressora", 0&, 1, ByVal "1", 1
30:      RegSetValueEx nKey, "ImpressoraFarmacia", 0&, 1, ByVal "1", 1
31:      RegQueryStringValue nKey, "Layout", LayouRegEdit
32:   End If
      
33:   SenhaBanco = RecuperaSenha(Val(LayouRegEdit))
      
      UsuarioBanco = "SA"
      If LayouRegEdit = 46 Then UsuarioBanco = "SAVE"
      'SenhaBanco = "123456"
      
      
34:   Set Banco = rdoEnvironments(0).OpenConnection("", rdDriverNoPrompt, False, "DSN=SaveSql;UID=" & UsuarioBanco & ";PWD=" & SenhaBanco & ";")
      
35:   Banco.QueryTimeout = 3600
      
      'SV.ConnectionTimeout = 90
      'SV.Open "DSN=SaveSQL", "sa", SenhaBanco
      
      CarregaParametroContabilidade
      'SETA A CIDADE ONDE O HOSPITAL ESTÁ
      '----------------------------------------------
36:   CidadeSede = 3500709
      '----------------------------------------------
      
      'RAPHAEL 19/03/2013 09:52 PEGA DATA/HORA DO SERVER
      sql = "SELECT CONVERT(VARCHAR(11),GETDATE(),110) AS DATA, CONVERT(VARCHAR(11),GETDATE(),114) AS HORA, GETDATE() AS DATAHORA"
      Set TblDataHora = Banco.OpenResultset(sql, 3)
      
37:   LimpaImagens
      
38:   SetAmpulheta
      
      Set fsot = CreateObject("Scripting.FileSystemObject")
      If fsot.FileExists(App.path & "\Save.exe") Then
         Set fsot1 = fsot.GetFile(App.path & "\Save.exe")
         VersaoExecutavelLocal = Format(fsot1.DateLastModified, "DD/MM HH:NN")
      Else
         VersaoExecutavelLocal = "00/00 00:00"
      End If
39:   MDISave.Caption = "Gestão Hospitalar (" & Versao & " - " & VersaoExecutavelLocal & ") - M2 Tecnologia (c) " '-" & LayouRegEdit
40:   MDISave.StatusBar.Panels(1).text = "Save: Gestão Hospitalar - " & Versao & " - M2 Tecnologia (c)"


42:   If Dir(App.path & "\CID10.hlp") <> "" Then App.HelpFile = App.path & "\CID10.hlp"
      
43:   frmSplash.Show 1
44:   Saida = False
45:   Show
      
'41:   If TipoHospital = 0 Then If Dir(App.path & "\ImgFundo.jpg") <> "" Then Me.Picture = LoadPicture(App.path & "\ImgFundo.jpg")
'      If TipoHospital = 1 Then If Dir(App.path & "\ImgFundo_TIPO1.jpg") <> "" Then Me.Picture = LoadPicture(App.path & "\ImgFundo_TIPO1.jpg")
'      If TipoHospital = 2 Then If Dir(App.path & "\ImgFundo_TIPO2.jpg") <> "" Then Me.Picture = LoadPicture(App.path & "\ImgFundo_TIPO2.jpg")

'41:   If TipoHospital = 0 Then If Dir(App.path & "\ImgFundo.jpg") <> "" Then picOriginal.Picture = LoadPicture(App.path & "\ImgFundo.jpg")
'      If TipoHospital = 1 Then If Dir(App.path & "\ImgFundo_TIPO1.jpg") <> "" Then picOriginal.Picture = LoadPicture(App.path & "\ImgFundo_TIPO1.jpg")
'      If TipoHospital = 2 Then If Dir(App.path & "\ImgFundo_TIPO2.jpg") <> "" Then picOriginal.Picture = LoadPicture(App.path & "\ImgFundo_TIPO2.jpg")




46:   DoEvents
'47:   StatusMenu False
      
48:   frmLogin.Show
      
      'Funcionalidade interna
      'frmLogin.loginAutomatico "SA", "FOREVER"
      
49:   reSetAmpulheta
      
      
      'If TipoHospital = 0 Then
         MDISave.BackColor = CorHospital
         MDISave.MenuXP.BackColor = CorHospital
      'Else
      '   MDISave.BackColor = &HC0FFC0     '&HFFFFFF
      '   MDISave.MenuXP.BackColor = &HC0FFC0     '&HFFFFFF
      'End If
      
Exit Sub
Erro:
   If Erl = 9875 Then
      MsgBox " Conexão com  o Banco de dados do Interfaceamento inválida. Verifique", vbCritical
      End
   Else
      TratarErro "MDI", "MDIForm_Load", Err.Number, Err.Description, Erl
   End If
   'TratarErro "MDI", "MDIForm_Load", Err.Number, Err.Description, Erl
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Saida = True
End Sub

Private Sub MDIForm_Resize()
   AjustaImagemFundo
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
   Dim ind As Long
   
   'Rotina necessário para atualizar o Atualiza.exe
   If Dir(App.path & "\ArquivoTemporario.exe") <> "" Then
      If Dir(App.path & "\Atualiza.exe") <> "" Then
         FinalizaAplicativo "Atualiza"
         Kill App.path & "\Atualiza.exe"
      End If
      Name App.path & "\ArquivoTemporario.exe" As App.path & "\Atualiza.exe"
   End If
   
   For ind = 0 To Forms.Count - 1
      If UCase(Forms(ind).Name) = "FRMATENDIMENTOENFERMAGEM" Then
         Unload FrmAtendimentoEnfermagem
      End If
   
      If UCase(Forms(ind).Name) = "FRMATENDIMENTOPACIENTE" Then
         Unload FrmAtendimentoPaciente
      End If
   
      If UCase(Forms(ind).Name) = "FRMPRESCRICAOMEDICAELETRONICA" Then
         Unload FrmPrescricaoMedicaEletronica
      End If
   Next
   
   FinalizaAplicativo "save_est"
   FinalizaAplicativo "save_patri"
   FinalizaAplicativo "save_TISS"
   FinalizaAplicativo "save_custo"
   
End Sub

Private Sub MenuXP_Click(ByVal id As Long)
On Error GoTo Erro
   Dim NomeMenu As String
   Dim idMenu As Long
   
   Dim UtilizaFinanceiroAntigo As Long
   
   'Variavel que define se o cliente utiliza o financeiro novo ou o antigo
   UtilizaFinanceiroAntigo = Val(RecuperaCampo(1, "1", "UTILIZAFINANCEIROANTIGO", "PARAMETRO"))

   NomeMenu = MenuXP.MenuItems.key(id)
   
   If MenuNovo Then
      NomeMenu = RecuperaNomeAntigoMenu(NomeMenu)
   End If
   
   idMenu = recuperaIdMenu(NomeMenu)
   
   Select Case NomeMenu
   '=============================
   'Estoque
   '=============================
   Case "mnuEst"
      If Dir(App.path & "\Estoque\Save_Est.exe", vbArchive) <> "" Then
         Shell Chr(34) & App.path & "\Estoque\Save_Est.exe" & Chr(34) & " " & Format(CodigoUsuario, "0000000000") & SenhaUsuario, vbMaximizedFocus
      End If
   '=============================
   'Patrimonio
   '=============================
   Case "mnuPtr"
      If Dir(App.path & "\Patrimonio\Save_Patri.exe", vbArchive) <> "" Then
         Shell Chr(34) & App.path & "\Patrimonio\Save_Patri.exe" & Chr(34) & " " & Format(CodigoUsuario, "0000000000") & SenhaUsuario, vbMaximizedFocus
      End If
   '=============================
   'Parametros
   '=============================
   '-----------------------------
   'Cadastro
   '-----------------------------
   'Geral
   Case "mnuPar_Cad_Ger"
      FrmCadastroGeral.Show
'   Case "mnuPar_Cad_Ger_Pac"
'      frmPacientes.Show
'   Case "mnuPar_Cad_Ger_Med"
'      frmMedico.Show
'   Case "mnuPar_Cad_Ger_Enf"
'      frmEnfermeira.Show
'   Case "mnuPar_Cad_Ger_Lei"
'      frmEnfermagem.Show
'   Case "mnuPar_Cad_Ger_Tec"
'      FrmTecnico.Show
'   Case "mnuPar_Cad_Ger_For"
'      frmFornecedores.Show
'   Case "mnuPar_Cad_Ger_Ter"
'      frmTerceiros.Show
'   Case "mnuPar_Cad_Ger_Soc"
'      FrmCad_Socio.Show
'   Case "mnuPar_Cad_Ger_Con"
'      frmConvenios.Show
'   Case "mnuPar_Cad_Ger_Ccu"
'      frmCentroCusto.Show
'   Case "mnuPar_Cad_Ger_CNa"
'      frmCidades.Show
'   Case "mnuPar_Cad_Ger_UEn"
'      FrmUnidadeEncaminhamento.Show
'   Case "mnuPar_Cad_Ger_Ati"
'      FrmAtividade.Show
'   Case "mnuPar_Cad_Ger_Fam"
'      frmFamilia.Show
'   Case "mnuPar_Cad_Ger_Reg"
'      frmRegimento.Show
'   Case "mnuPar_Cad_Ger_Pen"
'      frmPendencia.Show
'   Case "mnuPar_Cad_Ger_Mar"
'      FrmMarca.Show
'   Case "mnuPar_Cad_Ger_Per"
'      frmPertence.Show
'   Case "mnuPar_Cad_Ger_Ser"
'      FrmSer_CadServico.Show
'   Case "mnuPar_Cad_Ger_GrC"
'      FrmFIGrupo.Show

   'Faturamento
   Case "mnuPar_Cad_Fat_MAu"
      If Layout = 1 Then
         frmJustificativaAuditoria.Show
      End If
   Case "mnuPar_Cad_Fat_CIn"
      frmCaraterInternacao.Show
   Case "mnuPar_Cad_Fat_Mot"
      frmMotivoAlta.Show
   Case "mnuPar_Cad_Fat_Atu"
      frmAtuacao.Show
   'Case "mnuPar_Cad_Fat_Tia"
   '   frmTipoAto.Show
   Case "mnuPar_Cad_Fat_Esp"
      FrmEspecialidadesCirurgica.Show
   Case "mnuPar_Cad_Fat_CPr"
      FrmCategoriaProcedimento.Show
   Case "mnuPar_Cad_Fat_Tab"
      FrmTabelas.Show
   Case "mnuPar_Cad_Fat_CCo"
        Unload frmConvenios
      frmConvenios.ConsultaConvenio = True
      frmConvenios.Show
   'Relatório
   Case "mnuPar_Cad_Rel"
      frmRelCadastros.Show
   
   'Conv. Próprio
   Case "mnuPar_Cad_Cpr_Ben"
      FrmConv_CadBeneficiario.Show
   Case "mnuPar_Cad_Cpr_EGr"
      FrmConv_CadEmpresa.Show
   
   Case "MnuPar_Cad_Pla"
      frmEscalaPlantao.Show
      
   '-----------------------------
   'Configurações
   '-----------------------------
   Case "mnuPar_Par"
      frmParametros.Show
   Case "mnuPar_Par_Global"
      frmParametros.Show
   Case "mnuPar_Par_Email"
      frmServidorEmail.Show
   '-----------------------------
   'Segurança
   '-----------------------------
   Case "mnuPar_Seg_Cnv"
      FrmTravamentoConta.Show
   Case "mnuPar_Seg_Usu"
      FrmSeguranca1.Show
   Case "mnuPar_Seg_Pre"
      FrmPresSeguranca.Show
   Case "mnuPar_Seg_Per"
      FrmPerfilUsuario.Show
   Case "mnuPar_FIP"
      FrmFiParametro.Show
   
   '-----------------------------
   'Utilitarios
   '-----------------------------
   Case "mnuPar_Utl_Efd"
      FrmEliminaDuplicidade.Show vbModal
   Case "mnuPar_Utl_Reg"
      frmRelRegimento.Show
   Case "mnuPar_Utl_Pen"
      frmRelPendencia.Show
      'frmPendencia.Show
   Case "mnuPar_Utl_CIM"
      FrmImpressora.Show
   'Exportações
   Case "mnuPar_Utl_Exp_AIH"
      'If Layout = 4 Then
      '   frmExportacaoSUS.Show
      'Else
         frmExportacaoSUSUnificada.Show
      'End If
   Case "mnuPar_Utl_Exp_APC"
      FrmExportacaoAPAC.Show
   Case "mnuPar_Utl_Exp_UNI"
      frmExportarFaturaUnimed.Show
   Case "mnuPar_Utl_Exp_APS"
      frmExportarApas.Show
   Case "mnuPar_Utl_Exp_COT"
      frmExportacaoCotacao.Show
   Case "mnuPar_Utl_Exp_DME"
      FrmFiExportacaoDmed.Show
   Case "mnuPar_Utl_Exp_CPD"
      frmCPDH.Show
   Case "mnuPar_Utl_IFC"
      frmImportacaoFicha.Show
   Case "mnuPar_Utl_IFN"
      If Layout = 27 Then
         Form2.Show
      Else
         frmImportacaoFornecedores.Show
      End If
      
   'tela de log
   Case "mnuInf_Log"
      frmRelLog.Show
   'Relação Materiais Agendados
   Case "mnuInf_MatEquip"
      frmRelMaterais.Show
   'Motivos Auditoria
   Case "mnuPar_Cad_Fat_Jus"
      frmJustificativaAuditoria.Show
   Case "mnuPar_Utl_LOG"
      FrmLOG.Show
   
   Case "mnuPar_Utl_ICH"
      frmIntegracaoCIH1.Show
   Case "mnuPar_Utl_Etq"
      If Layout >= 10 Then
         frmSameEtiqueta.Show
      Else
         FrmEmissaoEtiquetaFaturamento.Show
      End If
   Case "mnuPar_Utl_Imp"
      'Importação do Laboratório de Agudos - Sistema Esmeralda
      frmImportacaoLaboratorio.Show
   Case "mnuPar_Utl_Exp_Nef"
      'Exportação do Laboratório de fernadopolis para a nefrodata
      FrmLaudo_Interfaceamento.Show
   Case "mnuPar_Utl_Exp_HEM"
      'Exportação do Laboratório de fernadopolis para a nefrodata
      frmExportacaoHemocentro.Show
   
   Case "mnuPar_RHC_CFU"
      FrmRH_CadFuncionario.Show
   Case "mnuPar_RHC_LRF"
      FrmRH_Refeicao.Show
   
   
   
   
   '=============================
   'Recepção
   '=============================
   Case "mnuRec_Int"
      If Layout >= 2 Then
         Unload frmRegistroInternacao1
         frmRegistroInternacao1.Show
      Else
         frmRegistroInternacao.Show
      End If
   Case "mnuRec_Amb"
      'If (Layout <> 10 And Layout <> 11 And Layout <> 9 And Layout <> 8 And Layout <> 7 And Layout <> 1 And Layout <> 4) And PossuiLicenca = False Then
      '   frmRegistroEntradaAmbPA.Cadastro = 2
      '   frmRegistroEntradaAmbPA.Caption = "Registro Ambulatorial"
      '   frmRegistroEntradaAmbPA.lblTipoRegistro.Caption = "Ambulatorial"
         'frmRegistroEntradaAmbPA.lblTipoRegistro1.ForeColor = &HFFFF&
      '   frmRegistroEntradaAmbPA.lblTipoRegistro1.Caption = "Ambulatorial"
         'frmRegistroEntradaAmbPA.TxtSUS.Enabled = False
         'frmRegistroEntradaAmbPA.TxtNomeSUS.Enabled = False
         'frmRegistroEntradaAmbPA.txtDias.Enabled = False
         'frmRegistroEntradaAmbPA.CmbTipoAcidente.Enabled = False
      '   frmRegistroEntradaAmbPA.cmdCancelar.Value = True
      '   frmRegistroEntradaAmbPA.Show
      'Else
         frmRegistroEntradaAmb.Cadastro = 2
         frmRegistroEntradaAmb.Caption = "Registro Ambulatorial"
         'frmRegistroEntradaAmb.lblTipoRegistro.Caption = "Ambulatorial"
         frmRegistroEntradaAmb.lblTipoRegistro1.Caption = "Ambulatorial"
         frmRegistroEntradaAmb.cmdCancelar.value = True
         frmRegistroEntradaAmb.Show
      'End If
   Case "mnuRec_Urg"
       'If (Layout <> 10 And Layout <> 11 And Layout <> 9 And Layout <> 8 And Layout <> 7 And Layout <> 1 And Layout <> 4) And PossuiLicenca = False Then
       '  frmRegistroEntradaAmbPA.Cadastro = 3
       '  frmRegistroEntradaAmbPA.Caption = "Registro de Urgência"
       '  frmRegistroEntradaAmbPA.lblTipoRegistro.Caption = "Urgência"
         'frmRegistroEntradaAmbPA.lblTipoRegistro1.ForeColor = &H4000&
       '  frmRegistroEntradaAmbPA.lblTipoRegistro1.Caption = "Urgência"
         'frmRegistroEntradaAmbPA.TxtSUS.Enabled = False
         'frmRegistroEntradaAmbPA.TxtNomeSUS.Enabled = False
         'frmRegistroEntradaAmbPA.txtDias.Enabled = False
         'frmRegistroEntradaAmbPA.CmbTipoAcidente.Enabled = True
       '  frmRegistroEntradaAmbPA.cmdCancelar.Value = True
       '  frmRegistroEntradaAmbPA.Show
      'Else
         frmRegistroEntradaAmb.Cadastro = 3
         frmRegistroEntradaAmb.Caption = "Registro de Urgência"
         'frmRegistroEntradaAmb.lblTipoRegistro.Caption = "Urgência"
         frmRegistroEntradaAmb.lblTipoRegistro1.Caption = "Urgência"
         frmRegistroEntradaAmb.cmdCancelar.value = True
         frmRegistroEntradaAmb.Show
      'End If
   Case "mnuRec_Con"
      frmRegistroConsulta.Caption = "Registro de Consultas"
      frmRegistroConsulta.Show
   Case "mnuRec_Reg"
      FrmConsultaRegistro1.Show
   Case "mnuRec_Ces"
      FrmControleVisita.Show
   Case "mnuRec_Cpe"
      frmControlePertence.Show
   Case "mnuRec_Cin"
      frmSer_LancamentoServico.Show
   Case "mnuRec_Gco"
      frmGestaoConta.Show
   Case "mnuRec_Ate"
      FrmAgendaTelefonica.Show
   Case "mnuRec_Coc"
      FrmOcorrencias.Show
   Case "mnuRec_TMe"
      FrmTransfernciaMedica.Show
   Case "mnuRec_LDi"
      FrmLancamentoDiaria.Show
      
   Case "mnuRec_Ain"
      frmAvaliacaoInterno.Show
   Case "mnuRec_POP"
      FrmPesquisa.Show
   Case "mnuRec_Les"
      frmListaEspera.Show
   Case "mnuRec_CAm"
      frmControleAmbulancia.Show
      
   '-----------------------------
   'Relatórios
   '-----------------------------
   Case "mnuRec_Rel_Dec"
      frmRelDeclaracao.Show
   Case "mnuRec_Rel_Ate"
      frmRelAtendimentoAmb.Show
   Case "mnuRec_Rel_Ips"
      FrmRelInternadoPS.Show
   Case "mnuRec_Rel_Toc"
      frmRelTaxaOcupacao.Show
   Case "mnuRec_Rel_Moh"
      frmRelMapaOcupacaoHospitalar.Show
   Case "mnuRec_Rel_Rgl"
      frmRelRegistroGeral.Show
   Case "mnuRec_Rel_Eat"
      FrmRelEmpresasAtendida.Show
   Case "mnuRec_Rel_Dag"
      FrmRelDadosGuia.Show
   Case "MnuRec_Rel_CAm"
      FrmRelResgateAmbulancia.Show
   Case "MnuRec_Rel_Obs"
      frmRelObservacao.Show
   Case "MnuRec_Rel_AML"
      FrmRelAtendimentoMalaDireta.Show
   Case "MnuRec_Rel_CVi"
      FrmRelControleVisita.Show
   Case "MnuRec_Rel_RMA"
      frmRelMicroArea.Show
   Case "MnuRec_Rel_TEs"
      FrmRelTempoAtendimento.Show
   Case "mnuRec_Rel_Acr"
      frmRelAtendimentoCor.Show
   Case "MnuRec_Rel_TAH"
      frmRelTempoAtendimentoPorAla.Show
   Case "MnuRec_Rel_Dia"
      frmRelDiarias.Show
   '-----------------------------
   'Controle de Prontuários
   '-----------------------------
   Case "mnuRec_Cpr_LUs"
      FrmSameCad_LocalArmazenamento.Show
   Case "mnuRec_Cpr_TAr"
      FrmSameCad_TipoArmazenamento.Show
   Case "mnuRec_Cpr_TPr"
      FrmSameTransferenciaProntuario.Show
      
   '=============================
   'SAME
   '=============================
   '-----------------------------
   'Arquivamento
   '-----------------------------
   Case "mnuSam_Arq_Ens"
      FrmSame.Show
   Case "mnuSam_Arq_Prp"
      frmRelSamePendencias.Show
   Case "mnuSam_Arq_Mes"
      frmRelSameEntradaSaida.Show
   Case "mnuSam_Arq_CAr"
      FrmSame_ConsultaProntuario.Show
   
   '-----------------------------
   'Estatistica
   '-----------------------------
   Case "MnuSam_Est_Cen"
      frmRelSameCenso.Show
   Case "mnuSam_Est_Pac"
      FrmRelHistoricoPaciente.Show
   Case "mnuSam_Est_AFA"
      frmRelAtendimentoFaixaEtaria.Show
   Case "mnuSam_Est_ATT"
      frmRelAtendimentoTipoTratamento.Show
   Case "mnuSam_Est_Cnv"
      frmRelAtendimentoConvenio.Show
   Case "mnuSam_Est_Cid"
      frmRelCidadeMedico1.Show
   Case "mnuSam_Est_Med"
      frmRelAtendimento.Show
   Case "mnuSam_Est_Prd"
      frmRelAtendimentoProcedimento.Show
   Case "mnuSam_Est_Esp"
      frmRelInternacaoEspMedica.Show
   Case "mnuSam_Est_Enc"
      frmRelAtendimentoOrigemEncaminhamento.Show
   Case "mnuSam_Est_Alt"
      frmRelAlta.Show
   Case "mnuSam_Est_Nas"
      frmRelNascidos.Show
   'Case "mnuSam_Est_Obi"
      'frmObitos.Show
   Case "mnuSam_Est_Tra"
      frmRelTransferencia.Show
   Case "mnuSam_Est_Pfm"
      frmRelPacientesFemininoMasculino.Show
   Case "mnuSam_Est_ROb"
      frmRelObitos.Show
   Case "mnuSam_Est_RDR"
      frmRelDrs.Show
   Case "mnuSam_Est_RIn"
      FrmRelIndicadores.Show
   Case "mnuSam_Est_RMP"
      If Layout = 44 Then
         frmRelNascidos.Show
      Else
         FrmRelacaoMP.Show
      End If
   Case "mnuSam_Est_Pre"
      FrmRelPreAlta.Show
   Case "mnuSam_Est_Ger"
      frmRelEstatisticaGeral.Show
   Case "mnuSam_Est_Anu"
      FrmEstatisticasGerenciais.Show
   Case "mnuSam_Est_Pro"
      FrmEstatisticasGerenciaisPA.Show
   Case "mnuSam_Est_Tin"
      frmRelTipoInternacao.Show
   
   '-----------------------------
   'CONSULTA DO SAME ANTERIOR A 2006 DE CAMPINAS
   '-----------------------------
   Case "mnuSam_Con"
      frmSAMEConsulta.Show
   
   '=============================
   'Enfermagem
   '=============================
   Case "mnuEnf_Alp"
      abrirTela frmAltaPaciente, idMenu, CodigoUsuario
   Case "mnuEnf_Pal"
      FrmPreAlta.Show
   Case "mnuEnf_Tle"
      frmTransferenciaLeito.Show
   Case "mnuEnf_Bin"
      frmBoletimInformativo.Show
   Case "mnuEnf_MñA"
      frmMovimentoNaoAutorizado.Show
   Case "mnuEnf_Etq"
      frmEtiquetaMatMed.Show
'   Case "mnuEnf_Etq"
'      If Layout = 8 Then
'         FrmEmissaoEtiquetaPrescricao.Show
'      ElseIf Layout = 1 Or Layout = 9 Then
'         FrmEtiquetaProdutoLayout1.Show
'      Else
'         FrmEtiquetaProduto.Show
'      End If
   '-----------------------------
   'CCIH
   '-----------------------------
   Case "mnuEnf_CCH_Fic"
      FrmCCIHFicha.Show
   Case "mnuEnf_CCH_Ser"
      FrmCCIHServico.Show
   Case "mnuEnf_CCH_Cci"
      frmRelRegistroInternacao.Show
   Case "mnuEnf_CCH_Bac"
      FrmBacteria.Show
   Case "mnuEnf_CCH_Cor"
      FrmBacteriaAntimiCrobiano.Show
   Case "mnuEnf_CCH_Ass"
      frmCir_CirurgiaCategoria.Show
   Case "mnuEnf_CCH_UNe_Cri"
      FrmCCIH_Criterio.Show
   Case "mnuEnf_CCH_UNe_ENe"
      frmCCIH_RelEstatistica.Show
   Case "mnuEnf_CCH_UNe_Fic"
      FrmCCIH_Ficha.Show
   
   '-----------------------------
   'Cirurgico
   '-----------------------------
   Case "mnuenf_cir_CCi"
      FrmCIR_CadCirurgia.Show
   Case "mnuenf_cir_Pre"
      If Layout = 10 Or Layout = 29 Or Layout = 43 Or Layout = 53 Then
         FrmCir_Dispensacao.Show
      Else
         Dim CentroCirurgico_BOX As Integer
         CentroCirurgico_BOX = Val(RecuperaCampo(1, "1", "UTILIZACENTROCIRURGICO_BOX", "PARAMETRO"))
         If CentroCirurgico_BOX = 1 Then
            FrmCir_Dispensacao.UTILIZACENTROCIRURGICO_BOX = 1
            FrmCir_Dispensacao.Show
         Else
            FrmCir_LancamentoAgendamentoCirurgico.Show
         End If
      End If
   Case "mnuenf_cir_Lmm"
      If Layout >= 10 Or (Layout = 1 And LoteValidade = 1) Then
         FrmPrescricaoEletronicaPeriodo.Show
         FrmPrescricaoEletronicaPeriodo.TelaCentroCirurgico = True
         FrmPrescricaoEletronicaPeriodo.Caption = "Lançamento de Gasto do C.Cirúrgico": FrmPrescricaoEletronicaPeriodo.Caption = ""
      Else
         frmCir_ConferenciaPreAgendamento.Show
      End If
   Case "mnuenf_cir_Ldc"
      abrirTela frmCirurgia, idMenu, CodigoUsuario
   Case "mnuenf_cir_lob"
      If Layout = 10 Then
         FrmCir_DadosObstetrico2.Show
      Else
         FrmCir_DadosObstetrico.Show
      End If
   Case "mnuenf_cir_Par"
      FrmDadosObstetricoParto.Show
   Case "mnuenf_cir_IMe"
      FrmDadosMedico.Show
   Case "mnuenf_cir_Agc"
      If Layout >= 10 Or Layout = 8 Or Layout = 1 Then
         frmCir_AgendamentoCirurgias.Show
      Else
         frmCir_AgendamentoCirurgico.Show
      End If
   Case "mnuenf_cir_Cad"
      frmCir_CadGeral.Show
   'Relatório
   Case "mnuenf_cir_Rel_Mac"
      If Layout >= 10 Then
         frmRelAgendamentoCirurgico.Show
      Else
         frmCir_RelMapaCirurgico.Show
      End If

   Case "mnuenf_cir_Rel_Esp"
      FrmCir_RelMaterialAgendamentoCirurgico.Show
   Case "mnuEnf_cir_Rel_Tax"
      frmRelTaxaOcupacao.Show
   Case "mnuEnf_Cir_Rel_Rci"
      frmCir_RelDadoCirurgica.Show
   Case "mnuEnf_cir_Rel_PAR"
      frmCir_RelacaoNascidos.Show
      
   '-----------------------------
   'Agencia Transfusional
   '-----------------------------
   Case "mnuEnf_Age_Gru"
      FrmHemo_GrupoSanguineo.Show
   Case "mnuEnf_Age_Rec"
      FrmHemo_Receptor.Show
   
   Case "mnuEnf_Ate"
      FrmAtendimentoEnfermagem.Show
   Case "mnuEnf_PKU"
      abrirTela FrmDadosPKU, idMenu, CodigoUsuario
      'FrmDadosPKU.Show
   Case "mnuEnf_TPE"
      abrirTela frmTestePezinhoExportacao, idMenu, CodigoUsuario
      'frmTestePezinhoExportacao.Show
   Case "mnuEnf_Lim"
      FrmMan_LimpezaLeito.Show
   Case "mnuEnf_ACo"
      frmIntegracaoConvenio.Show
      
   '=============================
   'Médico
   '=============================
   Case "mnuMed_Atm"
      '2013/01/07 Samuel/Visita Dracena
      'Comentado por inutilidade
      'If Layout = 29 Or Layout = 30 Or Layout = 32 Or Layout = 100 Or Layout = 10 Or Layout = 46 Or Layout = 12 Or Layout = 50 Or Layout = 19 Or Layout = 53 Then
         FrmAtendimentoPaciente.Show
      'Else
      '   frmAtendimentoMedico.Show
      'End If
   Case "mnuMed_Pme", "mnuMed_Pel"
      If Layout = 4 Then
         frmPrescricaoEletronica2.Show
      ElseIf (Layout = 21 Or Layout = 24 Or Layout = 33 Or Layout = 38 Or Layout = 42 Or Layout = 49 Or Layout = 52) And NomeMenu = "mnuMed_Pme" Then
         FrmPrescricaoMedicaEletronica21.TelaCentroCirurgico = False
         FrmPrescricaoMedicaEletronica21.Show
      ElseIf (Layout = 29 Or Layout = 32 Or Layout = 46 Or Layout = 19 Or Layout = 53 Or Layout = 47 Or Layout = 51 Or Layout = 54) And NomeMenu = "mnuMed_Pme" Then
         FrmPrescricaoMedicaEletronica.TelaCentroCirurgico = False
         FrmPrescricaoMedicaEletronica.Show
      ElseIf Layout >= 9 Or (Layout = 1 And LoteValidade = 1) Or Layout = 8 Then
         FrmPrescricaoEletronicaPeriodo.TelaCentroCirurgico = False
         FrmPrescricaoEletronicaPeriodo.Show
      Else
         frmPrescricaoEletronica.Show
      End If
   Case "mnuMed_EPr"
      If Layout = 1 Then
         frmRelEtiquetaMedicacaoPaciente.Show
      Else
         If Layout = 12 Or Layout = 10 Or Layout = 11 Or Layout = 27 Or Layout = 23 Or Layout >= 40 Then
            FrmPrescricaoMedicaEletronica.TelaCentroCirurgico = False
            FrmPrescricaoMedicaEletronica.Show
         ElseIf Layout = 39 Then
            FrmPrescricaoMedicaEletronica21.TelaCentroCirurgico = False
            FrmPrescricaoMedicaEletronica21.Show
         Else
            FrmEmissaoEtiquetaPrescricao.Show
         End If
      End If
   Case "mnuMed_Age"
      If Layout <> 27 Then
         frmAgendamentoConsulta.Show
      Else
         frmAgendamentoConsulta_Proced.Show
      End If
   '-----------------------------
   'Relatórios
   '-----------------------------
   Case "MnuMed_Rel_RAg"
      FrmRelAgendamentoConsulta.Show
   Case "mnuMed_Rel_Cre"
      FrmRelMedicosCredenciados.Show
   Case "MnuMed_Rel_RIP"
      frmPres_RelUnidade.Show
      
   '-----------------------------
   'Laudos
   '-----------------------------
   Case "mnuMed_Lau_Rpa"
      If Layout = 12 Or Layout = 14 Or Layout = 1 Then
         FrmLaudo_PacienteNovo.TelaFat = False
         FrmLaudo_PacienteNovo.Show
      Else
         FrmLaudo_Paciente.TelaFat = False
         FrmLaudo_Paciente.Show
      End If
   Case "mnuMed_Lau_Cmo"
      FrmLaudo_Modelos.Show
   Case "mnuMed_Lau_Dla"
      If Layout = 9 Then
         'FrmLaudo_Lancamento1.Show
      Else
         FrmLaudo_Lancamento.Show
      End If
   Case "mnuMed_Lau_ELa"
      FrmLaudo_Entrega.Show
   'Relatórios
   Case "MnuMed_Lau_Rel_Rla"
      FrmLaudo_RelConsulta.Show
   Case "mnuMed_Lau_Rel_ERe"
      FrmLaudo_RelExames.Show
   Case "MnuMed_Lau_Rel_ENo"
      FrmLaudo_RelNota.Show
   Case "mnuMed_Lau_Rel_Ate"
      FrmLaudo_RelAtendimentos.Show
   Case "mnuMed_Lau_Int"
      FrmLaudo_VerificaInterfaceamento.Show
   Case "mnuMed_Lau_CEA"
      FrmLaudo_ConsultaExame.Show
   
   '-----------------------------
   'UTI
   '-----------------------------
   Case "mnuMed_UTI_CCU"
      frmUti2.Show
   Case "mnuMed_UTI_Cig"
      frmUtiCadastro.Show
      
   Case "mnuMed_UTI_Cig"
      frmUtiCadastro.Show
   
   Case "mnuMed_UTI_Dis"
      FrmUTINeoDispensacao.Show
   
   '-----------------------------
   'LABORATÓRIO
   '-----------------------------
   Case "mnuMed_Lab"
      If Dir(App.path & "\Laboratorio\Save_Lab.exe", vbArchive) <> "" Then
         Shell Chr(34) & App.path & "\Laboratorio\Save_Lab.exe" & Chr(34) & " " & Format(CodigoUsuario, "0000000000") & SenhaUsuario, vbMaximizedFocus
      End If
         
   Case "mnuMed_Per"
      FrmCad_PresPergunta.Show
      
   Case "mnuMed_Cad"
      frmCad_Prescricao.Show
      
   Case "mnuMed_PRe"
      If Layout = 21 Or Layout = 33 Or Layout = 38 Or Layout = 39 Or Layout = 42 Or Layout = 49 Or Layout = 52 Then FrmPres_RepetirPrescricao.Show
      
   Case "mnuMed_APr"
      FrmPres_Autorizacao.Show
      
   Case "mnuMed_Agp"
      frmAgendamentoConsulta_Proced.Show

   '=============================
   'Faturamento
   '=============================
   Case "mnuFat_Lch"
      If Layout >= 7 Or Layout = 4 Then
         frmLancamentoConta2.TelaLaudo = False
         frmLancamentoConta2.TelaLanServicoProfissional = False
         frmLancamentoConta2.Show
      'ElseIf Layout = 4 Then
      '   frmLancamentoProcedimento.Show
      Else
         'PARA A NATHALIA RODRIGUDS DO FATURAMENTO, ABRE A TELA NOVA
         'If CodigoUsuario = 236 And Layout = 1 Then
            frmLancamentoConta2.TelaLaudo = False
            frmLancamentoConta2.TelaLanServicoProfissional = False
            frmLancamentoConta2.Show
         'Else
         '   frmLancamentoConta.Show
         'End If
      End If
   Case "mnuFat_Lsp"
      'If Layout >= 7 Or Layout = 1 Then
      If Layout >= 7 Then
         If Layout = 12 Then
            FrmLaudo_PacienteNovo.TelaFat = True
            FrmLaudo_PacienteNovo.TelaCantina = False
            FrmLaudo_PacienteNovo.Show
         Else
            FrmLaudo_Paciente.TelaFat = True
            FrmLaudo_Paciente.TelaCantina = False
            FrmLaudo_Paciente.Show
         End If
      Else
         frmLancamentoSADT1.Show
      End If
   Case "mnuFat_Lgc"
      frmDadosGuia.TelaFaturamento = False
      frmDadosGuia.Show
   Case "mnuFat_Gui"
      'não vai mais usar o form... tem que tirar da tabela menu
      frmGuia.Show
   '-----------------------------
   'Gerenciamento de Glosas
   '-----------------------------
   Case "mnuFat_Glo"
      FrmGlo_Menu.Show
'   Case "mnuFat_Glo_Lac"
'      frmGlo_LancamentoGlosa.Show
'   Case "mnuFat_Glo_Mot"
'      frmGlo_MotivoGlosa.Show
'   Case "mnuFat_Glo_TDf"
'      FrmGlo_TipoDiferenca.Show
'   Case "mnuFat_Glo_TPg"
'      FrmGlo_PagamentoGlobal.Show
'   'Relatórios
'   Case "mnuFat_Glo_Rel_Aud"
'      FrmGlo_RelAuditoria.Show
'   Case "mnuFat_Glo_Rel_Emi"
'      frmGlo_RelEmissaoGlosa.Show
'   Case "mnuFat_Glo_Rel_RDG"
'      frmGlo_RelGlosaDiferencas.Show
'   Case "mnuFat_Glo_Rel_Pro"
'      FrmGlo_DiferencaProntuario.Show
'   Case "mnuFat_Glo_Rel_Sam"
'      frmGlo_RelGlosaSame.Show
'   Case "mnuFat_Glo_Rel_Oco"
'      frmGlo_RelGlosaErroInterno.Show
'   Case "mnuFat_Glo_Rel_PGl"
'      FrmGlo_LancamentoPagamentoGlobal.Show
   Case "mnuFat_Cpe"
      FrmControlePendencia.TelaTransferencia = False
      FrmControlePendencia.Show
   Case "mnuFat_Bpe"
      
      frmBaixaPendencia.Show
   Case "mnuFat_Pmp"
      frmControleProntuarios.Show
   Case "mnuFat_Amm"
      'Feito no cliente 20/04/2012
'      If Layout = 10 Then
'         FrmTransferenciaParticular.Show
      If Layout = 21 Then
         FrmTransferenciaPrescricao.Show
      Else
         frmTransferenciaMatMed.Show
      End If
   Case "mnuFat_Tga"
      FrmTransferenciaGasto.Show
   Case "mnuFat_Clm"
      If Layout >= 9 Then
         FrmAuditoriaMatMed.Show
      Else
         frmAcertaContaPaciente.Show
      End If
   Case "mnuFat_Cdp"
      If Layout = 8 Then
         FrmLancamentoDadosCIH.Show
      Else
         frmAlteraInternacaoAltaConvenioPaciente.Show
      End If
   Case "mnuFat_Cch"
      frmImpressaoContaPaciente.Show
   Case "mnuFat_Grs"
      frmGradeTabelas.Show
      
   Case "mnuFat_Fch"
      If FechamentoFatura = 0 Then
         frmFechamentoContaAG.Show
      End If
   Case "mnuFat_Rch"
      If Layout = 21 Or Layout = 11 Or Layout = 12 Or Layout = 30 Or Layout = 39 Or Layout >= 42 Then
         FrmFat_TransferenciaIntExt.Show
      Else
         frmAberturaContaPaciente.Show
      End If
   Case "mnuFat_Tco"
      frmTravaConta1.Show
   Case "MnuFat_CAI"
      FrmInsumoAlteracao.Show
   Case "mnuFat_Fft"
      FrmFechamentoFaturamento.Show
   '-----------------------------
   'SUS
   '-----------------------------
   Case "mnuFat_Sus"
      FrmSUS_Menu.Show
'   'INTERNO
'   Case "mnuFat_SUS_Int_Las"
'      frmDadosAIH1.TelaInternacao = False
'      frmDadosAIH1.Show
'   Case "mnuFat_SUS_Int_RPF"
'      FrmLancamentoListaUAC.TipoRelacao = 1
'      FrmLancamentoListaUAC.Show
'   Case "mnuFat_Sus_Int_GAI"
'      frmSUS_RelGerenciamentoAIH.Show
'   Case "mnuFat_Sus_Int_UTI"
'      frmRelUTISUS.Show
'   Case "mnuFat_Sus_Int_Afa"
'      frmRelAIHFaturado.Show
'   Case "mnuFat_Sus_Int_Tat"
'      If Layout = 4 Then
'         frmRelTipoAto.Show
'      Else
'         'If Layout = 8 Then
'         '   frmRelFatSUSHospitalLAY8.TipoConvenio = 3
'         '   frmRelFatSUSHospitalLAY8.Show
'         'Else
'            frmRelFatSUSHospital.TipoConvenio = 3
'            frmRelFatSUSHospital.Show
'         'End If
'      End If
'   Case "mnuFat_Sus_Int_Pme"
'      If Layout = 4 Then
'         frmRelPlanilhaMedicos.Show
'      Else
'         frmRelFatSUSMedico.Show
'      End If
'   Case "mnuFat_Sus_Int_LMe"
'      If Layout = 8 Then
'         FrmLancamentoListaUAC.TipoRelacao = 0
'         FrmLancamentoListaUAC.Show
'      Else
'         FrmRelLaudoAIH.Show
'      End If
'   Case "mnuFat_Sus_Int_Pro"
'      FrmSUSProcessamento.Show
'   Case "mnuFat_Sus_Int_Imp"
'      frmSUS_Importacao.Show
'   Case "mnuFat_SUS_Int_SAI"
'      frmRelPacienteAlemLimite.Show
'
'   'AMBULATORIAL
'   Case "mnuFat_SUS_Amb_Apc"
'      FrmAPAC.Show
'   Case "mnuFat_Sus_Amb_Cla"
'      If Layout = 4 Then
'         frmRelConferenciaLancamentos.Show
'      Else
'         frmSUSRelConferenciaLancamento.TipoConvenio = 3
'         frmSUSRelConferenciaLancamento.Show
'      End If
'   Case "mnuFat_Sus_Amb_Clo"
'      If Layout = 4 Then
'         frmRelCapaLote.Show
'      Else
'         frmRelCapaLote1.Show
'      End If
'   Case "mnuFat_Sus_Amb_SCA"
'      frmSUS_CotaAmbulatorial.Show
'   Case "mnuFat_SUS_Amb_FPO"
'      FrmRelFPO.Show
'   Case "mnuFat_SUS_Amb_Crp"
'      frmCoRelacaoProcedimentosAtividades.Show
'   Case "mnuFat_SUS_Amb_CNO"
'      frmSUS_Amb_DivideLote.Show
      
   'IAMSPE
   Case "mnuFat_Iam"
      FrmIamspe_Menu.Show 'frmRelIAMSPE.Show
   'Relatórios
   Case "mnuFat_Rel_Fco"
      If FaturaDetalhada = 0 Then
         frmExtratoConta.Show
      Else
         frmRelFaturaContaHospitalar.Show
      End If
   Case "mnuFat_Rel_Gcg"
      frmRelGuiaConvenioGeral.Show
   Case "mnuFat_Rel_FRe"
      FrmRelFaturamentoPaciente.Show
   Case "mnuFat_Rel_Cen"
      FrmRelComprovanteEntrega1.Show
   Case "mnuFat_Rel_Cfa"
      frmRelComprovanteFaturamento.Show
   Case "mnuFat_Rel_Nsp"
      If Layout >= 10 Or Layout = 1 Then
         frmRelNotaServicoProfissional.Show
      Else
         frmRelNotasServicos.Show
      End If

   Case "mnuFat_Rel_Ccp"
      frmRelConferenciaContaPaciente.Show
   Case "mnuFat_Rel_Lco"
      If FaturaDetalhada = 0 Then
         frmRelListagemConvenio.Show
      Else
         frmRelControleEntrega.Show
      End If
   Case "mnuFat_Rel_Rfi" ' A ENDOSCOPIA DA BENE UTILIZA ISSO
      frmRelFichaIdentificacao.Show
   Case "mnuFat_Rel_Dpr"
      frmRelDivergencia.Show
   Case "mnuFat_Rel_EEL"
      FrmEmissaoExtrato.Show
    Case "mnuFat_Rel_FCC"
        FrmRelConsumoFaturadoCC.Show
    Case "mnuFat_Rel_FEM"
        frmRelFaturasEmitidas.Show
   
   '=============================
   'Informativo
   '=============================
   Case "mnuInf_Pfa"
      frmRelPacientesFaturados.Show
   Case "mnuInf_Pñf"
      frmRelPacientesNaoFaturados.Show
   Case "mnuInf_Pci"
      frmRelPacienteCID.Show
   Case "mnuInf_Gpr"
      If Layout = 3 Then
         frmRelGastoProcedimento.Show
      Else
         frmRelGastosProcedimentos.Show
      End If
   Case "mnuInf_Eir"
      FrmRelInsumoExamesPacientesFaturados.Show
   '------------------------
   'Honorários Médicos
   '------------------------
   Case "mnuInf_Hnr_Cre"
      frmRelCirurgia.Show
   Case "mnuInf_Hnr_Mph"
      frmRelParticMedicosProducaoHospital.Show
   Case "mnuInf_Hnr_Phh"
      frmRelMedicosNaoCooperadosPagamentoHospital.Show
   Case "mnuInf_Hnr_Hpf"
      frmRelHonorarioMedico1.Show
   Case "mnuInf_Hnr_Hmp"
      frmRelHonorarioMedicoProducao1.Show
   Case "mnuInf_Hnr_SAT"
      frmRelProducaoSADT.Show
   Case "mnuInf_Hnr_Acp"
      frmRelParto.Show

   Case "mnuInf_Fgl"
      If Layout = 4 Then
         frmRelMovimentoGlobal1.Show
      Else
         frmRelMovimentoGlobal.Show
      End If
   Case "mnuInf_Clp"
      frmRelLucroPerda.Show
   'Case "mnuInf_Civ"
   '   frmRelInsumosVendaCusto.Show
   Case "mnuInf_Pac"
      frmRelPacoteLucroPerda.Show
   Case "mnuInf_Cnt"
      frmRelControleInternacao.Show
   Case "mnuInf_Sit"
      FrmRelSituacaoFaturamento.Show
   Case "MnuInf_MMC"
      frmRelMatMedConvenioContrato.Show
   Case "MnuInf_Ger"
      FrmTelaGerencial.Show
   Case "mnuInf_PPC"
      frmRelConvenioProduto.Show
   Case "mnuInf_PIC"
      frmRelConvenioInsumo.Show
      
   '=============================
   'Financeiro
   '=============================
   '----------------------------
   'Contas à Pagar
   '----------------------------
   Case "mnuFin_Cop_Ter"
      frmFICredor.Show
   Case "mnuFin_Cop_Ban"
      FrmBancos.Show
   'Case "mnuFin_Cop_Tdo"
   '   frmTipoDocumento.Show
   Case "mnuFin_Cop_Lcp"
      'If Layout = 8 Or Layout = 9 Or Layout = 7 Then
      'If UtilizaFinanceiroAntigo = 0 Then
         frmFILancamentoPagar1.TelaFinanceiro = True
         frmFILancamentoPagar1.pedido = 0
         frmFILancamentoPagar1.Show
      'Else
      '   frmFILancamentoPagar.Pedido = 0
      '   frmFILancamentoPagar.Show
      'End If
   Case "mnuFin_Cop_Cpb"
      'If Layout = 8 Or Layout = 9 Or Layout = 7 Then
      'If UtilizaFinanceiroAntigo = 0 Then
         FrmFIBaixaPagamento1.Show
      'Else
      '   FrmFIBaixaPagamento.Show
      'End If
   Case "MnuFin_Cop_Bch"
      'If Layout = 8 Or Layout = 1 Or Layout = 9 Or Layout = 7 Or Layout = 11 Then
      'If UtilizaFinanceiroAntigo = 0 Then
         FrmFIPagamentoCheque1.Show
      'ElseIf Layout <> 9 Then
      '   FrmFIPagamentoCheque.Show
      'End If
   Case "mnuFin_Cop_PgE"
      FrmFiExportacaoBanco.Show
   Case "MnuFin_Cop_Adi"
      FrmFiAdiantamento.Show
   Case "mnuFin_Cop_EmC"
      FrmFiEmissaoChequeAvulso.Show
   Case "mnuFin_Cop_Eme"
      FrmFIEmissaoCheque.Show
   Case "mnuFin_Cop_ERC"
      FrmFIEmissaoReciboCheque.Show
   Case "mnuFin_Cop_PPa"
      FrmFIPrePagamento.Show
   Case "mnuFin_Cop_IIm"
      FrmFiLancamentoFinanceiroImposto.Show
   Case "mnuFin_Cop_Exp"
      FrmFiExportacaoNotaPrefeitura.Show

   'Relatórios
   Case "mnuFin_Cop_Rel_Cpp"
      frmRelContaVencimento.Show
   Case "mnuFin_Cop_Rel_Cpv"
      frmRelFICheque.Show
   Case "MnuFin_Cop_Rel_Cev"
      FrmFiDevolucaoCompensacao.Show
   Case "MnuFin_Cop_Rel_LGr"
      frmRelFiLancamentoGrupo.Show
   Case "MnuFin_Cop_Rel_CIm"
      frmRelContaImposto.Show
   Case "MnuFin_Cop_Rel_CFo"
      frmFiRelConferenciaFornecedor.Show
      
   '----------------------------
   'Contas à Receber
   '----------------------------
   Case "mnuFin_Cor_Out"
      'frmOutros.Show
      FrmCad_Outro.Show
   Case "mnuFin_Cor_Lcr"
      'If Layout <> 9 And Layout <> 7 And Layout <> 1 And Layout <> 11 And Layout <> 10 Then
      'If UtilizaFinanceiroAntigo = 1 Then
      '   frmFILancamentoReceber.Show
      'Else
         frmFILancamentoReceber1.Show
      'End If
   Case "mnuFin_Cor_Bpa"
      'If Layout = 9 Or Layout = 7 Or Layout = 1 Or Layout = 11 Or Layout = 10 Then
      'If UtilizaFinanceiroAntigo = 0 Then
         FrmFIBaixaRecebimento1.Show
      'Else
      '   FrmFIBaixaRecebimento.Show
      'End If
   Case "mnuFin_Cor_Bmp"
      FrmFiBaixaRecebimentoMultipla.Show
   
   Case "mnuFin_Cor_EBa"
      FrmFiExclusaoBaixaRecebimento.Show
   
   'importação
   Case "mnuFin_Cor_Imp_SLS"
      FrmFiImportacaoContasReceber.Show
   Case "mnuFin_Cor_Imp_Ret"
      FrmFiImportacaoRetornoBaixa.Show
   Case "mnuFin_Cor_Imp_EIm"
      FrmFiExclusaoImportacaoContasReceber.Show
   Case "MnuFin_Cor_Adi"
      FrmFiAdiantamentoReceber.Show
   Case "MnuFin_Cor_INF"
      FrmFiImpressaoNotaReceber.Show
   
   'Relatórios
   Case "mnuFin_Cor_Rel_Crp"
      frmRelFIRecebimento.Show
   Case "mnuFin_Cor_Rel_Pde"
      frmRelParticularesDebito.Show
   Case "mnuFin_Cor_Rel_Rec"
      frmRelReciboPagamento.Show
   Case "mnuFin_Cor_Rel_CFa"
      FrmFiRelFaturaRecebimento.Show
   
   Case "mnuFin_Cor_Rel_DFR"
      FrmRelDiferencaContaReceber.Show
   Case "MnuFin_Cor_LRL"
      FrmFiLancamentoReceberLote.Show
   Case "MnuFin_Cor_CRX"
      FrmFIControleRecebimento.Show
   Case "MnuFin_Cor_CNF"
      FrmFiCancelaNotaFiscal.Show
   Case "MnuFin_Cor_Aue"
      FrmFilaudoCaixa.Show
   Case "MnuFin_Cor_Can"
      If Layout = 12 Then
         FrmLaudo_PacienteNovo.TelaCantina = True
         FrmLaudo_PacienteNovo.TelaFat = True
         FrmLaudo_PacienteNovo.Show
      Else
         FrmLaudo_Paciente.TelaCantina = True
         FrmLaudo_Paciente.TelaFat = True
         FrmLaudo_Paciente.Show
      End If
   Case "mnuFin_Trf"
      FrmFiTransferencia.Show
   Case "MnuFin_Ldi"
      frmFILancamentoDireto.Show
   Case "mnuFin_Fca"
      frmRelFluxoCaixa.Show
   Case "mnuFin_CFi"
      FrmFICompromissoFinanceiro.Show
      
   '----------------------------
   'Pagamento de Terceiros
   '----------------------------
   Case "MnuFin_PTe"
      FrmTer_Menu.Show
'   'Cadastro
'   Case "MnuFin_PTe_Cad_Emp"
'      FrmTer_Empresa.Show
'   Case "MnuFin_PTe_Cad_Eve"
'      FrmTer_Evento.Show
'   Case "MnuFin_PTe_Cad_Par"
'      FrmTer_Parametro.Show
'   Case "MnuFin_PTe_Cad_TMo"
'      FrmTer_TipoMovimento.Show
'   Case "MnuFin_PTe_Cad_TIm"
'      FrmTer_TerceiroImposto.Show
'   'Movimentação
'   Case "MnuFin_PTe_Mov_ARP"
'      FrmTer_DataRepasse.Show
'   Case "MnuFin_PTe_Mov_IFa"
'      FrmTer_IntergracaoFaturamento.Show
'   Case "MnuFin_PTe_Mov_IFi"
'      FrmTer_IntegracaoFinanceiro.Show
'   Case "MnuFin_PTe_Mov_IRH"
'      FrmTer_IntegracaoFolha.Show
'   Case "MnuFin_PTe_Mov_LEE"
'      FrmTer_LancamentoEvento.Show
'   Case "MnuFin_PTe_Mov_LIn"
'      FrmTer_LancamentoIndividual.Show
'   Case "MnuFin_PTe_Mov_LLo"
'      FrmTer_LancamentoLote.Show
'   Case "MnuFin_PTe_Mov_SPS"
'      FrmTer_ImportacaoConvenioSaoLuiz.Show
'
'   'Relatórios
'   Case "MnuFin_PTe_Rel_CPg"
'      frmTer_RelConferenciaRepasse.Show
'   Case "mnuFin_PTe_Rel_Daf"
'      frmDARF.Show
'   Case "mnuFin_PTe_Rel_EDi"
'      frmDirf.Show
'   Case "MnuFin_PTe_Rel_ELo"
'      FrmTer_ImpressaoExtratoLote.Show
'   Case "mnuFin_PTe_Rel_RRe"
'      frmTer_RelRendimentoRetencao.Show
'   Case "mnuFin_PTe_Rel_ReT"
'      frmTer_RelRecebimentoTerceiro.Show
   
   'Compensação
   Case "mnuFin_CPR"
      'If Layout = 11 Then
         FrmFICompensacao.Show
      'Else
      '   FrmFICompensacao1.Show
      'End If
   
   'tipo de baixa
   Case "mnuFin_TBa"
      FrmFITipoBaixa.Show
   
   Case "mnuFin_ALA"
      FrmFiAlteracaoDados.Show
   
   Case "mnuFin_SFi"
      frmFiRelSituacaoFinanceiro.Show
   
   Case "mnuFin_Cfl"
      If Layout >= 11 Or Layout = 4 Or Layout = 1 Or Layout = 8 Then
         FrmFiFluxo_TipoOperacao.Show
      Else
         FrmFiFluxo_TipoOperacaoBanco.Show
      End If
   Case "mnuFin_FPR"
      If Layout >= 11 Or Layout = 4 Or Layout = 1 Or Layout = 8 Then
         FrmFIFluxoCaixa.Show
      Else
         FrmFIFluxoCaixaBanco.Show
      End If
   Case "mnuFin_OxR"
      FrmFiOrcamento.Show

   Case "mnuFin_SFP"
      frmFiSituacaoPaciente.Show
   
   Case "mnuFin_Com"
      frmFiRelComparativoFinanceiro.Show
      
   Case "MnuFin_FBD"
      FrmFiFehcamentoDiario.Show
      
   Case "mnuFin_PCM"
      frmPMor_Cadastros.Show

   '=============================
   'Contabilidade
   '=============================
   Case "mnuCtb_PlC"
      FrmCont_PlanoConta.Show
   Case "mnuCtb_CFo"
      FrmCont_FornecedorConta.Show
   Case "mnuCtb_His"
      FrmCont_Historico.Show
   Case "mnuCtb_CPP"
      frmProdutoConta.Show
   Case "mnuCtb_LEx"
      FrmCont_LancamentoExterno.Show
   Case "mnuCtb_Exp"
      FrmCont_Exportacao.Show
   Case "mnuCtb_RMC"
      FrmCont_ReaberturaMensal.Show
   Case "mnuCtb_ERM"
      FrmFiExportacaoRM.Show
   
   '----------------------------
   'Relatórios
   '----------------------------
   Case "mnuCtb_Rel_Res"
      If Layout = 8 Then
         frmRelResumoContabil.Show
      Else
         FrmCont_FechamentoMensal.Show
      End If
   Case "mnuCtb_Rel_Cct"
      frmRelCont_Conferencia.Show
   Case "mnuCtb_Rel_Con"
      frmCont_RelConta.Show
   
   '=============================
   'TISS
   '=============================
   Case "mnuTIS"
      If Dir(App.path & "\TISS\Save_TISS.exe", vbArchive) <> "" Then
         Shell Chr(34) & App.path & "\TISS\Save_TISS.exe" & Chr(34) & " " & Format(CodigoUsuario, "0000000000") & SenhaUsuario, vbMaximizedFocus
      End If
   
   '=============================
   'custo
   '=============================
   Case "mnuCut"

      If Dir(App.path & "\custo\save_custo.exe", vbArchive) <> "" Then
         Shell Chr(34) & App.path & "\custo\save_custo.exe" & Chr(34) & " " & Format(CodigoUsuario, "0000000000") & SenhaUsuario, vbMaximizedFocus
      End If
      
   Case "mnuFin_ETT"
        FrmFiExportacaoTotvs.Show
   
   End Select
Exit Sub
Erro:
   TratarErro Me.Name, "Execucutado", Err.Number, Err.Description, Erl
End Sub

Public Sub Tim_Licenca_Timer()
On Error GoTo Erro
   Dim tbl As rdoResultset
   Dim sql As String
      
   
   If Tim_Licenca.Interval = 3000 And PossuiLicenca = False And ExecutouLicenca Then
      If LicencaParametro.SQLFree = 1 Then
         If Format(LicencaParametro.DataExecucaoFree, "DD/MM/YYYY") <> Format(Date, "DD/MM/YYYY") Then
            Banco.Execute " EXECUTE SP_JOBSAVE"
         End If
      End If
      
      If Not ValidaStatusLicenca Then End
      'COM ESSA QTDE O SISTEMA SÓ IRÁ VERIFICAR A NOVA DATA QUANDO ATINGIR A 12:00 DO
      'DIA SEGUINTE COM ISSO MESMO SE O USUÁRIO TROCAR A DATA DO COMPUTADOR
      'O SISTEMA CONTINUARÁ PEGANDO DO SERVIDOR
      
      QuantidadeMinutoRestante = DateDiff("n", DataServidor & " " & Format(HoraServidor, "hh:mm"), DateAdd("d", 1, DataServidor) & " 00:00")
      Tim_Licenca.Interval = 65535 '1 MINUTO
   Else
      If PossuiLicenca Then Tim_Licenca.Enabled = False
   End If
   
   QuantidadeHora = QuantidadeHora + 1
   
   If QuantidadeHora >= QuantidadeMinutoRestante And QuantidadeMinutoRestante <> 0 Then
      VerificaLicenca
      Tim_Licenca.Interval = 3000
      QuantidadeHora = 0
   ElseIf ExecutouLicenca And IsDate(HoraServidor) Then
      HoraServidor = DateAdd("n", 1, HoraServidor)
   End If
   
   FechamentoBalanco
 
Exit Sub
Erro:
   TratarErro Me.Name, "Execucutado", Err.Number, Err.Description, Erl
End Sub

Private Sub FechamentoBalanco()
   On Error GoTo Erro
   
  'FEZ O FECHAMENTO DO BALANCO AUTOMATICO NO ULTIMO DIA DO MêS
   
   If Trim(DataServidor) = "" Then
      DataServidor = Format(Date, "YYYY/MM/DD")
      HoraServidor = Format(Date, "HH:MM:SS")
   End If
   
   'salto não quer que faça automático
   'If Layout = 42 Then Exit Sub
   
   'Jaboticabal não quer faça automático
   If Layout = 30 Then Exit Sub
   
   If Day(DataServidor) = IIf(Layout = 36, 1, UltimoDiaMes(Month(DataServidor), Year(DataServidor))) Then
   'If Day(DataServidor) = "23" Then
      sql = ""
      sql = sql & " SELECT ULTIMA_DATA_FECHAMENTO_BALANCO FROM PARAMETRO "
      Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
      If tbl.EOF = False Then
         If Format(tbl!ULTIMA_DATA_FECHAMENTO_BALANCO, "YYYY/MM/DD") < Format(DataServidor, "YYYY/MM/DD") Then
            sql = ""
            sql = sql & " EXEC SP_PROCESSA_BALANCO "
            Banco.Execute sql
         End If
      End If
      tbl.Close
   Else
   'VERIFICA SE NÃO MODOU O MES E FICOU SEM FAZER
      sql = ""
      sql = sql & " SELECT ULTIMA_DATA_FECHAMENTO_BALANCO FROM PARAMETRO "
      Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
      If tbl.EOF = False Then
         If Month(tbl!ULTIMA_DATA_FECHAMENTO_BALANCO) <> Month(DataServidor) Then
            sql = ""
            sql = sql & " EXEC SP_PROCESSA_BALANCO "
            Banco.Execute sql
         End If
      End If
      tbl.Close
   
   End If
   
   Exit Sub
Erro:
   Resume Next
End Sub


Public Sub AjustaImagemFundo()
On Error GoTo Erro
   Dim client_rect As RECT
   Dim client_hwnd As Long
   
   If TipoHospital = 0 Then If Dir(App.path & "\ImgFundo.jpg") <> "" Then picOriginal.Picture = LoadPicture(App.path & "\ImgFundo.jpg")
   If TipoHospital = 1 Then If Dir(App.path & "\ImgFundo_TIPO1.jpg") <> "" Then picOriginal.Picture = LoadPicture(App.path & "\ImgFundo_TIPO1.jpg")
  If TipoHospital = 2 Then If Dir(App.path & "\ImgFundo_TIPO2.jpg") <> "" Then picOriginal.Picture = LoadPicture(App.path & "\ImgFundo_TIPO2.jpg")
   
   picStretched.Move 0, 0, ScaleWidth, ScaleHeight
   picStretched.PaintPicture picOriginal.Picture, 0, 0, picStretched.ScaleWidth, picStretched.ScaleHeight, 0, 0, picOriginal.ScaleWidth, picOriginal.ScaleHeight
   Picture = picStretched.Image
   client_hwnd = FindWindowEx(Me.hwnd, 0, "MDIClient", vbNullChar)
   GetClientRect client_hwnd, client_rect
   InvalidateRect client_hwnd, client_rect, 1
Exit Sub
Erro:
   TratarErro "ModuleMagic", "FormataVersao", Err.Number, Err.Description, Erl
End Sub

Private Function RecuperaNomeAntigoMenu(NomeNovo As String) As String
   On Error GoTo Erro
   
   Dim tblaux As rdoResultset
   Dim SqlAux As String
   
   SqlAux = ""
   SqlAux = SqlAux & " SELECT NOMESUBNOVO " & Chr(13)
   SqlAux = SqlAux & " FROM MENU " & Chr(13)
   SqlAux = SqlAux & " WHERE NOMESUBALTERADO = '" & NomeNovo & "'" & Chr(13)
   
   Set tblaux = Banco.OpenResultset(SqlAux, rdOpenStatic)
   
   If tblaux.EOF = False Then
      RecuperaNomeAntigoMenu = tblaux!NOMESUBNOVO
   Else
      RecuperaNomeAntigoMenu = NomeNovo
   End If
   
   Exit Function
Erro:
   TratarErro
End Function

