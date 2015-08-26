VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{02F125F5-49EE-11D5-A561-0050BF395743}#1.0#0"; "OcxControl.ocx"
Object = "{9DD6E228-C5FE-11D8-B43F-0002444CD4A3}#1.0#0"; "ControleVB.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFornecedores 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fornecedores"
   ClientHeight    =   7560
   ClientLeft      =   3015
   ClientTop       =   1575
   ClientWidth     =   11910
   Icon            =   "frmFornecedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11910
   Begin VB.Frame FraRelacaoForn 
      BackColor       =   &H80000005&
      Caption         =   "Relação Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   960
      TabIndex        =   111
      Top             =   4680
      Visible         =   0   'False
      Width           =   6585
      Begin VB.CommandButton cmdIncluiRelacaoForn 
         Caption         =   "OK"
         Height          =   300
         Left            =   5985
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Salvar"
         Top             =   225
         Width           =   465
      End
      Begin OcxControl.Combo UseCmbFornecedor 
         Height          =   315
         Left            =   630
         TabIndex        =   113
         Top             =   225
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   556
         BorderStyle     =   0
      End
      Begin OcxControl.Grid grdRelacao 
         Height          =   870
         Left            =   90
         TabIndex        =   115
         Top             =   585
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   1535
         RowSel          =   1
         Row             =   1
         MouseIcon       =   "frmFornecedores.frx":0CCA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColSel          =   1
         Col             =   1
         CellPicture     =   "frmFornecedores.frx":0CE6
         CellFontSize    =   8,25
         CellFontName    =   "MS Sans Serif"
         CellBackColor   =   0
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         Caption         =   "Forn.:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   112
         Top             =   315
         Width           =   405
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   13441
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   794
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Fabricantes"
      TabPicture(0)   =   "frmFornecedores.frx":0D02
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "stbPadrao1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Fornecedores"
      TabPicture(1)   =   "frmFornecedores.frx":0D1E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "stbPadrao"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lista"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.ListBox Lista 
         Height          =   450
         Left            =   1080
         TabIndex        =   77
         Top             =   3360
         Visible         =   0   'False
         Width           =   7680
      End
      Begin TabDlg.SSTab stbPadrao 
         Height          =   7080
         Left            =   0
         TabIndex        =   42
         Top             =   45
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   12488
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   8388608
         TabCaption(0)   =   "&Inclusão/Alteração"
         TabPicture(0)   =   "frmFornecedores.frx":0D3A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fmeDados"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Pesquisa/Exclusão"
         TabPicture(1)   =   "frmFornecedores.frx":0D56
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(1)=   "Label2"
         Tab(1).Control(2)=   "Grid"
         Tab(1).Control(3)=   "cmbCampos"
         Tab(1).Control(4)=   "txtConteudo"
         Tab(1).Control(5)=   "CmdEtiqueta"
         Tab(1).ControlCount=   6
         Begin VB.CommandButton CmdEtiqueta 
            Caption         =   "Etiqueta"
            Height          =   330
            Left            =   -64440
            TabIndex        =   48
            Top             =   6600
            Width           =   1095
         End
         Begin VB.Frame fmeDados 
            BackColor       =   &H00EEEEEE&
            Height          =   6585
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   11670
            Begin VB.Frame fraProdutos 
               BackColor       =   &H80000005&
               Caption         =   "Grupo de produtos associados"
               Height          =   3015
               Left            =   6360
               TabIndex        =   91
               Top             =   1200
               Visible         =   0   'False
               Width           =   5055
               Begin OcxControl.Grid grdProdutos 
                  Height          =   2655
                  Left            =   120
                  TabIndex        =   92
                  Top             =   240
                  Width           =   4845
                  _ExtentX        =   8546
                  _ExtentY        =   4683
                  RowSel          =   1
                  Row             =   1
                  MouseIcon       =   "frmFornecedores.frx":0D72
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FixedCols       =   0
                  Cols            =   3
                  CellPicture     =   "frmFornecedores.frx":0D8E
                  CellFontSize    =   8,25
                  CellFontName    =   "MS Sans Serif"
                  CellBackColor   =   0
               End
            End
            Begin OcxControl.Text txtISS 
               Height          =   540
               Left            =   6600
               TabIndex        =   94
               Top             =   4560
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   953
               PromptChar      =   "_"
               Caption         =   "% ISS"
               CampoDecimais   =   2
               CampoTipo       =   0
               CorFundo        =   15658734
            End
            Begin VB.CheckBox chkRelacaoFornecedor 
               Caption         =   "Relação Forn."
               Height          =   330
               Left            =   5880
               Style           =   1  'Graphical
               TabIndex        =   105
               Top             =   5760
               Width           =   1470
            End
            Begin VB.TextBox txtObservacao 
               Height          =   1275
               Left            =   7440
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   107
               Tag             =   " TObservacao"
               Top             =   4800
               Width           =   4155
            End
            Begin VB.CheckBox chkTemporario 
               BackColor       =   &H00EEEEEE&
               Caption         =   "Temporário"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   3480
               TabIndex        =   103
               Top             =   5880
               Width           =   1185
            End
            Begin VB.Frame FraDadosBanco 
               BackColor       =   &H00EEEEEE&
               Caption         =   "Dados do Banco"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   945
               Left            =   120
               TabIndex        =   95
               Top             =   5160
               Width           =   3255
               Begin ControleVB.Text txtAgencia 
                  Height          =   525
                  Left            =   750
                  TabIndex        =   97
                  Top             =   225
                  Width           =   840
                  _ExtentX        =   1482
                  _ExtentY        =   926
                  MaxLength       =   8
                  PromptChar      =   "_"
                  Caption         =   "Agência"
                  CorFundo        =   15658734
               End
               Begin ControleVB.Text txtContaCorrente 
                  Height          =   525
                  Left            =   1650
                  TabIndex        =   98
                  ToolTipText     =   "Conta Corrente completa sem Pontos e Traços"
                  Top             =   225
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   926
                  PromptChar      =   "_"
                  Caption         =   "Conta c/ Dígito"
                  CorFundo        =   15658734
               End
               Begin ControleVB.Text txtBanco 
                  Height          =   525
                  Left            =   45
                  TabIndex        =   96
                  Top             =   225
                  Width           =   660
                  _ExtentX        =   1164
                  _ExtentY        =   926
                  PromptChar      =   "_"
                  Caption         =   "Banco"
                  CorFundo        =   15658734
               End
            End
            Begin VB.CheckBox chkDesativado 
               BackColor       =   &H00EEEEEE&
               Caption         =   "Desativado"
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   4680
               TabIndex        =   104
               Tag             =   "0NDesativado"
               Top             =   5880
               Width           =   1185
            End
            Begin VB.Frame FraTipoFornecedor 
               BackColor       =   &H00EEEEEE&
               Caption         =   "Centro de Custo do Fornecedor"
               ForeColor       =   &H00800000&
               Height          =   510
               Left            =   3480
               TabIndex        =   99
               Top             =   5160
               Width           =   3150
               Begin VB.CheckBox ChkSND 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "S.N.D"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   2340
                  TabIndex        =   102
                  Top             =   240
                  Width           =   780
               End
               Begin VB.CheckBox chkFarmacia 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Farmácia"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   45
                  TabIndex        =   100
                  Top             =   240
                  Width           =   990
               End
               Begin VB.CheckBox chkAlmoxarifado 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Almoxarifado"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   1050
                  TabIndex        =   101
                  Top             =   240
                  Width           =   1320
               End
            End
            Begin VB.CommandButton cmdSalvar 
               Caption         =   "&Salvar"
               Height          =   345
               Left            =   8430
               Style           =   1  'Graphical
               TabIndex        =   108
               ToolTipText     =   "Salvar"
               Top             =   6105
               Width           =   1050
            End
            Begin VB.CommandButton cmdCancelar 
               Caption         =   "&Cancelar"
               Height          =   345
               Left            =   10545
               Style           =   1  'Graphical
               TabIndex        =   110
               ToolTipText     =   "Cancelar"
               Top             =   6105
               Width           =   1020
            End
            Begin VB.TextBox TxtCodigo 
               Height          =   285
               Left            =   120
               MaxLength       =   9
               TabIndex        =   52
               Tag             =   " NFornecedor"
               ToolTipText     =   "Código do Procedimento"
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtDescricao 
               Height          =   285
               Left            =   2280
               MaxLength       =   155
               TabIndex        =   53
               Tag             =   " TDescricao"
               ToolTipText     =   "Descrição do Procedimento"
               Top             =   480
               Width           =   9255
            End
            Begin VB.Frame FmeCampos 
               BackColor       =   &H00EEEEEE&
               BorderStyle     =   0  'None
               Height          =   3900
               Left            =   60
               TabIndex        =   54
               Top             =   660
               Width           =   11565
               Begin VB.TextBox txtCodigoTotvs 
                  Height          =   285
                  Left            =   7440
                  MaxLength       =   9
                  TabIndex        =   116
                  Tag             =   "NCodigoTotvs"
                  Top             =   3600
                  Visible         =   0   'False
                  Width           =   1815
               End
               Begin ControleVB.Text txtRespNegoc 
                  Height          =   525
                  Left            =   7440
                  TabIndex        =   84
                  Top             =   2640
                  Width           =   4095
                  _ExtentX        =   7223
                  _ExtentY        =   926
                  MaxLength       =   0
                  PromptChar      =   "_"
                  Caption         =   "Contato responsável pela negociação"
               End
               Begin VB.TextBox txtValorMinimo 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   88
                  Top             =   3600
                  Width           =   1455
               End
               Begin VB.TextBox txtPrazoEntrega 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   120
                  TabIndex        =   87
                  Top             =   3600
                  Width           =   1455
               End
               Begin VB.CheckBox chkProdutos 
                  Caption         =   "Produtos"
                  Height          =   375
                  Left            =   9720
                  Style           =   1  'Graphical
                  TabIndex        =   90
                  Top             =   3480
                  Width           =   1695
               End
               Begin ControleVB.Text txtCondPgto 
                  Height          =   525
                  Left            =   3240
                  TabIndex        =   89
                  Top             =   3360
                  Width           =   3735
                  _ExtentX        =   6588
                  _ExtentY        =   926
                  MaxLength       =   0
                  PromptChar      =   "_"
                  Caption         =   "Condição de Pagamento"
               End
               Begin OcxControl.Text txtFundacao 
                  Height          =   540
                  Left            =   8760
                  TabIndex        =   58
                  Top             =   240
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   953
                  Mask            =   "99/99/9999"
                  MaxLength       =   0
                  PromptChar      =   "_"
                  Caption         =   "Fundação"
                  CampoTipo       =   2
               End
               Begin VB.Frame Frame1 
                  Caption         =   "Tipo de Pessoa"
                  ForeColor       =   &H00800000&
                  Height          =   975
                  Left            =   9960
                  TabIndex        =   59
                  Top             =   360
                  Width           =   1575
                  Begin VB.OptionButton optPessoa 
                     Caption         =   "Jurídica"
                     ForeColor       =   &H00800000&
                     Height          =   255
                     Index           =   1
                     Left            =   240
                     TabIndex        =   61
                     Top             =   600
                     Value           =   -1  'True
                     Width           =   975
                  End
                  Begin VB.OptionButton optPessoa 
                     Caption         =   "Física"
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Index           =   0
                     Left            =   240
                     TabIndex        =   60
                     Top             =   360
                     Width           =   1035
                  End
               End
               Begin OcxControl.Text txtCGCCPF 
                  Height          =   330
                  Left            =   5520
                  TabIndex        =   67
                  Top             =   1080
                  Width           =   4305
                  _ExtentX        =   7594
                  _ExtentY        =   582
                  BorderStyle     =   0
                  MaxLength       =   20
                  PromptChar      =   "_"
               End
               Begin VB.TextBox txtNomeFantasia 
                  Height          =   285
                  Left            =   120
                  MaxLength       =   100
                  TabIndex        =   56
                  Tag             =   "0TNomeFantasia"
                  Text            =   " "
                  ToolTipText     =   "Bairro"
                  Top             =   470
                  Width           =   4785
               End
               Begin VB.TextBox txtFORendereco 
                  Height          =   315
                  Left            =   2280
                  MaxLength       =   155
                  TabIndex        =   71
                  Tag             =   "0Tendereco"
                  ToolTipText     =   "Endereço"
                  Top             =   1680
                  Width           =   6300
               End
               Begin VB.TextBox txtFORbairro 
                  Height          =   315
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   73
                  Tag             =   "0Tbairro"
                  Text            =   " "
                  ToolTipText     =   "Bairro"
                  Top             =   1680
                  Width           =   2745
               End
               Begin VB.TextBox TxtNomeCidade 
                  Height          =   315
                  Left            =   900
                  MaxLength       =   30
                  TabIndex        =   76
                  Tag             =   "0FORcidade"
                  Text            =   " "
                  ToolTipText     =   "Cidade"
                  Top             =   2280
                  Width           =   7635
               End
               Begin VB.TextBox txtFORInscricao 
                  Height          =   285
                  Left            =   2280
                  MaxLength       =   15
                  TabIndex        =   65
                  Tag             =   "0TInscricao"
                  ToolTipText     =   "Inscrição Estadual"
                  Top             =   1080
                  Width           =   3135
               End
               Begin VB.TextBox TxtFORContaContabil 
                  Height          =   285
                  Left            =   120
                  MaxLength       =   20
                  TabIndex        =   63
                  Tag             =   " TContaContabil"
                  ToolTipText     =   "Conta Contábil"
                  Top             =   1080
                  Width           =   2025
               End
               Begin VB.TextBox TxtCidade 
                  Height          =   315
                  Left            =   180
                  TabIndex        =   75
                  Tag             =   "0Ncidade"
                  Top             =   2280
                  Width           =   705
               End
               Begin VB.TextBox txtEmail 
                  Height          =   345
                  Left            =   4080
                  MaxLength       =   150
                  TabIndex        =   83
                  Tag             =   "0TEmail"
                  ToolTipText     =   "Descrição do Procedimento"
                  Top             =   2880
                  Width           =   3315
               End
               Begin MSMask.MaskEdBox mskFORtelefone 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   79
                  Tag             =   "0TTelefone9"
                  ToolTipText     =   "Telefone"
                  Top             =   2880
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   556
                  _Version        =   393216
                  AllowPrompt     =   -1  'True
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskFORcep 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   69
                  Tag             =   "0Tcep"
                  ToolTipText     =   "CEP"
                  Top             =   1680
                  Width           =   2025
                  _ExtentX        =   3572
                  _ExtentY        =   556
                  _Version        =   393216
                  MaxLength       =   9
                  Mask            =   "99999-999"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskFax 
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   81
                  Tag             =   "0TFAX9"
                  ToolTipText     =   "Telefone"
                  Top             =   2880
                  Width           =   1905
                  _ExtentX        =   3360
                  _ExtentY        =   556
                  _Version        =   393216
                  AllowPrompt     =   -1  'True
                  PromptChar      =   "_"
               End
               Begin OcxControl.Combo UseCmbAtividade1 
                  Height          =   525
                  Left            =   5040
                  TabIndex        =   57
                  Top             =   240
                  Width           =   3615
                  _ExtentX        =   6376
                  _ExtentY        =   926
                  Caption         =   "Atividade"
               End
               Begin VB.Label lblCodigoTotvs 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código Totvs"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   7440
                  TabIndex        =   117
                  Top             =   3360
                  Visible         =   0   'False
                  Width           =   945
               End
               Begin VB.Label Label16 
                  Caption         =   "Pedido mínimo (R$)"
                  ForeColor       =   &H00800000&
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   85
                  Top             =   3360
                  Width           =   2415
               End
               Begin VB.Label Label12 
                  Caption         =   "Prazo entrega (dias)"
                  ForeColor       =   &H00800000&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   86
                  Top             =   3360
                  Width           =   1575
               End
               Begin VB.Label Label13 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Nome Fantasia"
                  ForeColor       =   &H00800000&
                  Height          =   225
                  Left            =   120
                  TabIndex        =   55
                  Top             =   240
                  Width           =   1170
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Endereço"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   70
                  Top             =   1440
                  Width           =   690
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Bairro"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   3
                  Left            =   8760
                  TabIndex        =   72
                  Top             =   1485
                  Width           =   645
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Cidade"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   4
                  Left            =   120
                  TabIndex        =   74
                  Top             =   2040
                  Width           =   495
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "CEP"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   6
                  Left            =   120
                  TabIndex        =   68
                  Top             =   1440
                  Width           =   315
               End
               Begin VB.Label lblFornecedores 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Insc. Estadual"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   7
                  Left            =   2280
                  TabIndex        =   64
                  Top             =   840
                  Width           =   1140
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "CNPJ"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   8
                  Left            =   5520
                  TabIndex        =   66
                  Top             =   825
                  Width           =   405
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Fone"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   9
                  Left            =   180
                  TabIndex        =   78
                  Top             =   2625
                  Width           =   360
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Conta Contábil"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   135
                  TabIndex        =   62
                  Top             =   855
                  Width           =   1035
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Fax"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Index           =   0
                  Left            =   2025
                  TabIndex        =   80
                  Top             =   2685
                  Width           =   255
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "E-mail"
                  ForeColor       =   &H00800000&
                  Height          =   195
                  Left            =   4080
                  TabIndex        =   82
                  Top             =   2640
                  Width           =   420
               End
            End
            Begin VB.CommandButton cmdExcluir 
               Caption         =   "&Excluir"
               Enabled         =   0   'False
               Height          =   345
               Left            =   9495
               Style           =   1  'Graphical
               TabIndex        =   109
               ToolTipText     =   "Salvar"
               Top             =   6105
               Width           =   1050
            End
            Begin OcxControl.Combo UseCmbAtividadeISS 
               Height          =   525
               Left            =   210
               TabIndex        =   93
               Top             =   4560
               Width           =   6345
               _ExtentX        =   11192
               _ExtentY        =   926
               Caption         =   "Atividade ISS"
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00EEEEEE&
               Caption         =   "Observação"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   7440
               TabIndex        =   106
               Top             =   4560
               Width           =   870
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00EEEEEE&
               Caption         =   "Código"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00EEEEEE&
               Caption         =   "Descrição"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   2280
               TabIndex        =   51
               Top             =   240
               Width           =   720
            End
         End
         Begin VB.TextBox txtConteudo 
            Height          =   315
            Left            =   -71580
            TabIndex        =   46
            Top             =   630
            Width           =   8295
         End
         Begin VB.ComboBox cmbCampos 
            Height          =   315
            ItemData        =   "frmFornecedores.frx":0DAA
            Left            =   -74910
            List            =   "frmFornecedores.frx":0DC0
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   630
            Width           =   3285
         End
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   5610
            Left            =   -74940
            TabIndex        =   47
            ToolTipText     =   "Pressione DELETE para excluir!"
            Top             =   975
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   9895
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FormatString    =   $"frmFornecedores.frx":0E03
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Que possuam"
            Height          =   195
            Left            =   -71580
            TabIndex        =   44
            Top             =   420
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Campos"
            Height          =   195
            Left            =   -74880
            TabIndex        =   43
            Top             =   420
            Width           =   570
         End
      End
      Begin TabDlg.SSTab stbPadrao1 
         Height          =   7125
         Left            =   -75000
         TabIndex        =   1
         Top             =   0
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   12568
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   -2147483635
         TabCaption(0)   =   "&Inclusão/Alteração"
         TabPicture(0)   =   "frmFornecedores.frx":0E97
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fmeDados1"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "&Pesquisa/Exclusão"
         TabPicture(1)   =   "frmFornecedores.frx":0EB3
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Grid1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "txtConteudo1"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmbCampos1"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.Frame fmeDados1 
            BackColor       =   &H00EEEEEE&
            Height          =   3825
            Left            =   -74940
            TabIndex        =   2
            Top             =   330
            Width           =   8085
            Begin VB.CommandButton cmdSalvar1 
               BackColor       =   &H00FCCE9C&
               Caption         =   "&Salvar"
               Height          =   375
               Left            =   4740
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Salvar"
               Top             =   3240
               Width           =   1050
            End
            Begin VB.CommandButton cmdCancelar1 
               BackColor       =   &H00FCCE9C&
               Caption         =   "&Cancelar"
               Height          =   375
               Left            =   6855
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "Cancelar"
               Top             =   3240
               Width           =   1050
            End
            Begin VB.TextBox TxtCodigo1 
               Height          =   315
               Left            =   990
               MaxLength       =   9
               TabIndex        =   4
               Tag             =   " NFabricante"
               ToolTipText     =   "Código do Procedimento"
               Top             =   360
               Width           =   1245
            End
            Begin VB.TextBox txtDescricao1 
               Height          =   315
               Left            =   3240
               MaxLength       =   50
               TabIndex        =   6
               Tag             =   " TDescricao"
               ToolTipText     =   "Descrição do Procedimento"
               Top             =   360
               Width           =   4275
            End
            Begin VB.Frame FmeCampos1 
               BackColor       =   &H00EEEEEE&
               BorderStyle     =   0  'None
               Height          =   2475
               Left            =   60
               TabIndex        =   7
               Top             =   690
               Width           =   7995
               Begin VB.CheckBox ChkDesativado1 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Desativado"
                  ForeColor       =   &H8000000D&
                  Height          =   315
                  Left            =   6300
                  TabIndex        =   33
                  Tag             =   "0NDesativado"
                  Top             =   2100
                  Width           =   1185
               End
               Begin VB.TextBox txtContato 
                  Height          =   315
                  Left            =   930
                  TabIndex        =   26
                  Tag             =   "0TContato"
                  Top             =   1650
                  Width           =   2925
               End
               Begin OcxControl.Combo UseCmbAtividade 
                  Height          =   525
                  Left            =   930
                  TabIndex        =   31
                  Top             =   1830
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   926
                  Caption         =   ""
               End
               Begin VB.TextBox txtFABInscricao 
                  Height          =   315
                  Left            =   4140
                  MaxLength       =   15
                  TabIndex        =   15
                  Tag             =   "0TInscricao"
                  ToolTipText     =   "Inscrição Estadual"
                  Top             =   90
                  Width           =   1695
               End
               Begin VB.TextBox TxtNomeCidade1 
                  Height          =   315
                  Left            =   4680
                  MaxLength       =   30
                  TabIndex        =   19
                  Tag             =   "0FABcidade"
                  Text            =   " "
                  ToolTipText     =   "Cidade"
                  Top             =   870
                  Width           =   2775
               End
               Begin VB.TextBox txtFABbairro 
                  Height          =   315
                  Left            =   930
                  MaxLength       =   20
                  TabIndex        =   17
                  Tag             =   "0Tbairro"
                  Text            =   " "
                  ToolTipText     =   "Bairro"
                  Top             =   870
                  Width           =   2145
               End
               Begin VB.TextBox txtFABendereco 
                  Height          =   315
                  Left            =   930
                  MaxLength       =   50
                  TabIndex        =   16
                  Tag             =   "0TEndereco"
                  ToolTipText     =   "Endereço"
                  Top             =   480
                  Width           =   6495
               End
               Begin VB.TextBox TxtCidade1 
                  Height          =   315
                  Left            =   3780
                  MaxLength       =   7
                  TabIndex        =   18
                  Tag             =   "0NCidade"
                  Top             =   870
                  Width           =   885
               End
               Begin VB.ListBox Lista1 
                  Height          =   1230
                  ItemData        =   "frmFornecedores.frx":0ECF
                  Left            =   4680
                  List            =   "frmFornecedores.frx":0ED6
                  TabIndex        =   28
                  Top             =   870
                  Visible         =   0   'False
                  Width           =   2775
               End
               Begin VB.CheckBox chkDist 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Distribuidor"
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   6060
                  TabIndex        =   8
                  Tag             =   "0NDisFab"
                  Top             =   150
                  Width           =   2205
               End
               Begin VB.TextBox txtEmail1 
                  Height          =   315
                  Left            =   4530
                  TabIndex        =   29
                  Tag             =   "0TEmail"
                  Top             =   1650
                  Width           =   2925
               End
               Begin MSMask.MaskEdBox mskFABcgc 
                  Height          =   315
                  Left            =   930
                  TabIndex        =   14
                  Tag             =   "0Tcgc"
                  ToolTipText     =   "CGC"
                  Top             =   90
                  Width           =   1635
                  _ExtentX        =   2884
                  _ExtentY        =   556
                  _Version        =   393216
                  MaxLength       =   18
                  Mask            =   "99.999.999/9999-99"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskFABtelefone 
                  Height          =   315
                  Left            =   2370
                  TabIndex        =   23
                  Tag             =   "0TTelefone"
                  ToolTipText     =   "Telefone"
                  Top             =   1260
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  _Version        =   393216
                  AllowPrompt     =   -1  'True
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskFABcep 
                  Height          =   315
                  Left            =   930
                  TabIndex        =   21
                  Tag             =   "0Tcep"
                  ToolTipText     =   "CEP"
                  Top             =   1260
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   556
                  _Version        =   393216
                  MaxLength       =   9
                  Mask            =   "99999-999"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox txtFax 
                  Height          =   315
                  Left            =   4140
                  TabIndex        =   24
                  Tag             =   "0TFax"
                  ToolTipText     =   "Telefone"
                  Top             =   1260
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   556
                  _Version        =   393216
                  AllowPrompt     =   -1  'True
                  PromptChar      =   "_"
               End
               Begin VB.Label Label11 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Atividade"
                  ForeColor       =   &H8000000D&
                  Height          =   225
                  Left            =   210
                  TabIndex        =   32
                  Top             =   2130
                  Width           =   705
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Fone"
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Index           =   14
                  Left            =   1950
                  TabIndex        =   22
                  Top             =   1320
                  Width           =   360
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "CGC"
                  ForeColor       =   &H8000000D&
                  Height          =   165
                  Index           =   13
                  Left            =   510
                  TabIndex        =   13
                  Top             =   180
                  Width           =   330
               End
               Begin VB.Label lblFornecedores 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Inscrição Estadual"
                  ForeColor       =   &H8000000D&
                  Height          =   375
                  Index           =   12
                  Left            =   2730
                  TabIndex        =   12
                  Top             =   180
                  Width           =   1305
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "CEP"
                  ForeColor       =   &H8000000D&
                  Height          =   165
                  Index           =   11
                  Left            =   510
                  TabIndex        =   20
                  Top             =   1320
                  Width           =   315
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Cidade"
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Index           =   10
                  Left            =   3240
                  TabIndex        =   11
                  Top             =   960
                  Width           =   495
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Bairro"
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Index           =   5
                  Left            =   450
                  TabIndex        =   10
                  Top             =   930
                  Width           =   405
               End
               Begin VB.Label lblFornecedores 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Endereço"
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Index           =   1
                  Left            =   180
                  TabIndex        =   9
                  Top             =   540
                  Width           =   690
               End
               Begin VB.Label lblFax 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Fax"
                  ForeColor       =   &H8000000D&
                  Height          =   255
                  Left            =   3810
                  TabIndex        =   25
                  Top             =   1350
                  Width           =   315
               End
               Begin VB.Label lblEmail 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "E-mail"
                  ForeColor       =   &H8000000D&
                  Height          =   285
                  Left            =   4020
                  TabIndex        =   30
                  Top             =   1710
                  Width           =   735
               End
               Begin VB.Label lblContato 
                  BackColor       =   &H00EEEEEE&
                  Caption         =   "Contato"
                  ForeColor       =   &H8000000D&
                  Height          =   225
                  Left            =   300
                  TabIndex        =   27
                  Top             =   1710
                  Width           =   705
               End
            End
            Begin VB.CommandButton cmdExcluir1 
               BackColor       =   &H00FCCE9C&
               Caption         =   "&Excluir"
               Enabled         =   0   'False
               Height          =   375
               Left            =   5805
               Style           =   1  'Graphical
               TabIndex        =   35
               ToolTipText     =   "Salvar"
               Top             =   3240
               Width           =   1050
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00EEEEEE&
               Caption         =   "Código"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   450
               TabIndex        =   3
               Top             =   450
               Width           =   495
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00EEEEEE&
               Caption         =   "Descrição"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   2430
               TabIndex        =   5
               Top             =   450
               Width           =   720
            End
         End
         Begin VB.ComboBox cmbCampos1 
            Height          =   315
            ItemData        =   "frmFornecedores.frx":0EE2
            Left            =   90
            List            =   "frmFornecedores.frx":0EF2
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   630
            Width           =   3285
         End
         Begin VB.TextBox txtConteudo1 
            Height          =   315
            Left            =   3420
            TabIndex        =   40
            Top             =   630
            Width           =   8295
         End
         Begin MSFlexGridLib.MSFlexGrid Grid1 
            Height          =   5985
            Left            =   60
            TabIndex        =   41
            ToolTipText     =   "Pressione DELETE para excluir!"
            Top             =   1020
            Width           =   11685
            _ExtentX        =   20611
            _ExtentY        =   10557
            _Version        =   393216
            Rows            =   1
            Cols            =   5
            FormatString    =   $"frmFornecedores.frx":0F1C
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Campos"
            Height          =   195
            Left            =   -74880
            TabIndex        =   37
            Top             =   420
            Width           =   570
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Que possuam"
            Height          =   195
            Left            =   -71580
            TabIndex        =   38
            Top             =   420
            Width           =   975
         End
      End
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FCCE9C&
      BackStyle       =   1  'Opaque
      Height          =   420
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1620
   End
End
Attribute VB_Name = "frmFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Inclui As Boolean, Trava As Long, LinhaVelha As Long, Pode As Boolean
Dim Inclui1 As Boolean, Trava1 As Long, LinhaVelha1 As Long, Pode1 As Boolean
Public Distribuidor1 As Byte
Public TelaPedidoCotacao As Boolean

Private Sub CarregaQuery()
      Dim tbl As rdoResultset
      Dim sql As String

On Error GoTo Erro

20       sql = ""
30       sql = sql & " SELECT ATIVIDADE, NOME "
40       sql = sql & " FROM ATIVIDADE "
50       sql = sql & " ORDER BY NOME "
60       UseCmbAtividade.sql = sql
70       UseCmbAtividade.AtualizarDados
         
80       sql = ""
90       sql = sql & " SELECT ATIVIDADE, NOME "
100      sql = sql & " FROM ATIVIDADE "
110      sql = sql & " ORDER BY NOME "
120      UseCmbAtividade1.sql = sql
130      UseCmbAtividade1.AtualizarDados

140      sql = ""
150      sql = sql & " SELECT FORNECEDOR, RTRIM(LTRIM(CGC)) + ' - ' + DESCRICAO "
160      sql = sql & " FROM FORNECEDORES"
170      sql = sql & " WHERE ISNULL(DESATIVADO, 0) = 0"
180      sql = sql & " ORDER BY DESCRICAO "
190      useCmbFornecedor.sql = sql
200      useCmbFornecedor.AtualizarDados
         
210      sql = ""
220      sql = sql & " SELECT ATIVIDADE, NOME "
230      sql = sql & " FROM ATIVIDADE_ISS "
240      sql = sql & " ORDER BY NOME "
250      UseCmbAtividadeISS.sql = sql
260      UseCmbAtividadeISS.AtualizarDados

Exit Sub
Erro:
    TratarErro "frmFornecedores", "CarregaQuery", Err.Number, Err.Description, Erl

End Sub

Private Sub chkProdutos_Click()
          
On Error GoTo Erro

10        If chkProdutos.value = 1 Then
20            fraProdutos.Visible = True
              If Val(grdProdutos.Tag) = 0 Then
30              MontarGridProdutos
              End If
40        Else
50            fraProdutos.Visible = False
60        End If

Exit Sub
Erro:
    TratarErro "frmFornecedores", "chkProdutos_Click", Err.Number, Err.Description, Erl
          
End Sub

Private Sub chkRelacaoFornecedor_Click()
         
On Error GoTo Erro

20       If chkRelacaoFornecedor.value = 0 Then
30          FraRelacaoForn.Visible = False
40       Else
50          FraRelacaoForn.Visible = True
60       End If

Exit Sub
Erro:
    TratarErro "frmFornecedores", "chkRelacaoFornecedor_Click", Err.Number, Err.Description, Erl
        
End Sub

'FARMACIA (TABELA FABRICANTES)
'OUTROS   (TABELA FORNECEDORES)

Private Sub cmbCampos_Click()
On Error GoTo Erro
      Dim rs As New ADODB.Recordset
      
20       If cmbCampos.ListIndex = 5 Then
30       rs.Open "select * from fornecedores", SV
40       Screen.MousePointer = 11
50       Grid.Visible = False
60       Grid.Rows = 1
70       Do Until rs.EOF
80              Grid.AddItem "" & vbTab & rs!FORNECEDOR & vbTab & rs!Descricao & vbTab & rs!NOMEFANTASIA
90              rs.MoveNext
100      Loop
110      Grid.Visible = True
120      rs.Close
130      Screen.MousePointer = 0
140      End If


Exit Sub
Erro:
    TratarErro "frmFornecedores", "cmbCampos_Click", Err.Number, Err.Description, Erl

End Sub

'===============================================================================
'Segundo cadastro (Tabela Fornecedores)
'===============================================================================

Private Sub cmbCampos1_Click()
On Error GoTo Erro
      Dim rs As New ADODB.Recordset
20       If cmbCampos1.ListIndex = 3 Then
30          rs.Open "select * from fabricantes", SV
40          Grid1.Visible = False
50          Screen.MousePointer = 11
60          Grid1.Rows = 1
70          Do Until rs.EOF
80             Grid1.AddItem "" & vbTab & rs!Fabricante & vbTab & rs!Descricao & vbTab & rs!CGC & vbTab & rs!Inscricao
90             rs.MoveNext
100         Loop
110         Grid1.Visible = True
120         rs.Close
130         LinhaVelha1 = 0
140         Grid1.SetFocus
150         Screen.MousePointer = 0
160      End If

Exit Sub
Erro:
    TratarErro "frmFornecedores", "cmbCampos1_Click", Err.Number, Err.Description, Erl
         
End Sub

Private Sub cmdCancelar_Click()
On Error GoTo Erro
20       Inclui = True
30       Trava = 0
40       LimpaCamposPublico Me
50       chkRelacaoFornecedor.value = 0
60       chkFarmacia.value = 0: chkAlmoxarifado.value = 0: ChkSND.value = 0
70       chkTemporario.value = 0: chkTemporario.Enabled = False
80       txtDescricao.Enabled = False 'Descrição ou o próximo campo
90       FmeCampos.Enabled = False
100      txtCodigo.Enabled = True
110      cmdExcluir.Enabled = False
120      txtCodigo.SetFocus
130      UseCmbAtividade1.DataStorage = 0
140      txtCGCCPF.text = ""
150      txtBanco.text = ""
160      txtAgencia.text = ""
170      TxtContaCorrente.text = ""
180      txtISS.text = ""
190      UseCmbAtividadeISS.DataStorage = 0
200       optPessoa(0).value = 0
210       optPessoa(1).value = 1
220       txtFundacao.text = "__/__/____"
230       txtPrazoEntrega.text = ""
240       txtValorMinimo.text = ""
250       txtCondPgto.text = ""
260       txtRespNegoc.text = ""
270       InicializaGridProdutos
280       InicializaGridRelacao

Exit Sub
Erro:
    TratarErro "frmFornecedores", "cmdCancelar_Click", Err.Number, Err.Description, Erl

End Sub

Private Sub CmdCancelar1_Click()
On Error GoTo Erro
10          Inclui1 = True
20          Trava1 = 0
30          LimpaCamposPublico Me
40          TxtDescricao1.Enabled = False  'Descrição ou o próximo campo
50          FmeCampos1.Enabled = False
60          TxtCodigo1.Enabled = True
70          cmdExcluir1.Enabled = False
80          TxtCodigo1.SetFocus
90          UseCmbAtividade.DataStorage = 0

Exit Sub
Erro:
    TratarErro "frmFornecedores", "CmdCancelar1_Click", Err.Number, Err.Description, Erl

End Sub

Private Sub CmdEtiqueta_Click()
         Dim tbl As rdoResultset
         Dim sql As String
On Error GoTo Erro
10       sql = "SELECT   A.DESCRICAO,ISNULL(A.ENDERECO,'') AS ENDERECO," & Chr(13)
20       sql = sql & "   ISNULL(A.CEP,'') AS CEP,'' AS COMPLEMENTO," & Chr(13)
30       sql = sql & "  ISNULL(B.DESCRICAO,'') AS CIDADE,ISNULL(B.UF,0) AS UF, A.NOMEFANTASIA " & Chr(13)
40       sql = sql & "   FROM FORNECEDORES A LEFT JOIN CIDADES B ON A.CIDADE=B.CIDADE" & Chr(13)
50       If cmbCampos.text = "Código" Then
60          sql = sql & "  WHERE A.FORNECEDOR=" & txtConteudo.text
70       ElseIf cmbCampos.text = "Descrição" Then
80          sql = sql & "  WHERE A.DESCRICAO LIKE '" & txtConteudo.text & "%'"
90       ElseIf cmbCampos.text = "Atividade" Then
100         sql = sql & "  LEFT JOIN ATIVIDADE C ON A.ATIVIDADE=C.ATIVIDADE "
110         sql = sql & " WHERE C.NOME LIKE '" & txtConteudo & "%'"
120      ElseIf cmbCampos.text = "Atividade" Then
130         sql = sql & " WHERE A.NOMEFANTASIA LIKE '" & txtConteudo & "%'"
140      End If
150      Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
160      Etiqueta tbl
170      tbl.Close

Exit Sub
Erro:
    TratarErro "frmFornecedores", "CmdEtiqueta_Click", Err.Number, Err.Description, Erl
End Sub

Private Sub cmdExcluir_Click()
         
         Dim tbl As rdoResultset
         
On Error GoTo Erro
20       sql = ""
30       sql = sql & " SELECT CREDOR FROM FILANCAMENTO"
40       sql = sql & " WHERE TERCEIROS = 0 "
50       sql = sql & " AND   CREDOR = " & Val(txtCodigo.text)
         
60       Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
         
70       If tbl.EOF = False Then
80          MsgBox "Este fornecedor não pode ser excluído, pois há lançamentos de notas no financeiro!", vbInformation
90          Exit Sub
100      End If
         
110      If MsgBox("Deseja realmente excluir este regsitro?", vbQuestion + vbYesNo, Caption) = vbYes Then
120         SV.Execute "DELETE FROM fornecedores WHERE fornecedor = " & txtCodigo
130         cmdCancelar_Click
140      End If

Exit Sub
Erro:
    TratarErro "frmFornecedores", "cmdExcluir_Click", Err.Number, Err.Description, Erl

End Sub

Private Sub CmdExcluir1_Click()
         
On Error GoTo Erro
20       If MsgBox("Deseja realmente excluir este regsitro?", vbQuestion + vbYesNo, Caption) = vbYes Then
30            SV.Execute "DELETE FROM fabricantes WHERE fabricante = " & TxtCodigo1.text
40            CmdCancelar1_Click
50       End If

Exit Sub
Erro:
    TratarErro "frmFornecedores", "CmdExcluir1_Click", Err.Number, Err.Description, Erl

End Sub

Private Sub cmdIncluiRelacaoForn_Click()
        
         Dim tbl As rdoResultset
         
On Error GoTo Erro
20       If Val(txtCodigo.text) = 0 Then
30          MsgBox "É necessário informar o fornecedor!", vbInformation
40          Exit Sub
50       End If
         
60       If useCmbFornecedor.DataStorage = 0 Then
70          MsgBox "É necessário informar o fornecedor da associação!", vbInformation
80          useCmbFornecedor.SetFocus
90          Exit Sub
100      End If
            
110      If useCmbFornecedor.DataStorage = txtCodigo.text Then
120         MsgBox "O fornecedor da associação deve ser diferente do fornecedor!", vbInformation
130         useCmbFornecedor.SetFocus
140         Exit Sub
150      End If
            
160      sql = ""
170      sql = sql & " SELECT FORNECEDOR " & Chr(13)
180      sql = sql & " FROM FORNECEDOR_RELACAO " & Chr(13)
190      sql = sql & " WHERE FORNECEDOR = " & txtCodigo.text & Chr(13)
200      sql = sql & " AND   FORNECEDOR_RELACAO = " & useCmbFornecedor.DataStorage & Chr(13)
210      Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
220      If tbl.EOF = False Then
230         MsgBox "Este fornecedor já está associado!", vbInformation
240         useCmbFornecedor.SetFocus
250         Exit Sub
260      End If
270      tbl.Close
         
280      sql = ""
290      sql = sql & " INSERT INTO FORNECEDOR_RELACAO (" & Chr(13)
300      sql = sql & "        FORNECEDOR, FORNECEDOR_RELACAO, ATUALIZACAO )" & Chr(13)
310      sql = sql & " VALUES (" & Chr(13)
320      sql = sql & txtCodigo.text & "," & Chr(13)
330      sql = sql & useCmbFornecedor.DataStorage & "," & Chr(13)
340      sql = sql & "'" & CorrenteTimeStamp & "')" & Chr(13)
350      Banco.Execute sql
         
360      MontarGridRelacao
         
370      useCmbFornecedor.DataStorage = 0
380      useCmbFornecedor.SetFocus

Exit Sub
Erro:
    TratarErro "frmFornecedores", "cmdIncluiRelacaoForn_Click", Err.Number, Err.Description, Erl

End Sub

Private Sub cmdSair_Click()
On Error GoTo Erro
10        Unload Me

Exit Sub
Erro:
    TratarErro "frmFornecedores", "cmdSair_Click", Err.Number, Err.Description, Erl
End Sub

Private Sub cmdSair1_Click()
On Error GoTo Erro
10       Unload Me

Exit Sub
Erro:
    TratarErro "frmFornecedores", "cmdSair1_Click", Err.Number, Err.Description, Erl
End Sub

Private Sub cmdSalvar_Click()

      Dim sql As String, Afetado As Integer
      Dim tbl As rdoResultset
      Dim tblaux As rdoResultset
      Dim i As Integer

         'If Not PreencheuTodos(Me) Then 'Testa se todos os campos obrigatórios foram preenchidos
         '   MsgBox "Preencha os campos em vermelho!", vbInformation, Caption
         '   Exit Sub
         'End If
         
10    On Error GoTo Erro
20       If Not IsNumeric(txtCodigo.text) Then
30          MsgBox "O Codigo deve ser informado!", vbInformation
40          txtCodigo.SetFocus
50          Exit Sub
60       End If
         
70       If Trim(txtDescricao.text) = "" Then
80          MsgBox "A Descrição deve ser informada!", vbInformation
90          txtDescricao.SetFocus
100         Exit Sub
110      End If
         
120      If Inclui And Not TelaPedidoCotacao And Trim(TxtFORContaContabil.text) = "" Then
130         MsgBox "A Conta deve ser informada!", vbInformation
140         TxtFORContaContabil.SetFocus
150         Exit Sub
160      End If
         
170      If Not Corretos(Me) Then 'Testa se os campos preenchidos estão com os dados corretos
180         MsgBox "Corrija os campos em vermelho!", vbInformation, Caption
190         Exit Sub
200      End If
         
210      If Layout <> 20 And Layout <> 21 And txtCGCCPF.Vazio Then
220         If Layout = 1 Or Layout >= 10 Then
230            MsgBox "CGC/CPF não informado", vbInformation
240            If txtCGCCPF.Enabled Then txtCGCCPF.SetFocus
250            Exit Sub
260         Else
270            If MsgBox("CPF/CNPJ não informado, deseja continuar!", vbInformation + vbYesNo) = vbNo Then
280               If txtCGCCPF.Enabled Then txtCGCCPF.SetFocus
290               Exit Sub
300            End If
310         End If
            
320      End If
         
330      If Layout <> 11 And Trim(txtCGCCPF.text) <> "" Then
340         sql = ""
350         sql = sql & " SELECT CGC, DESCRICAO FROM FORNECEDORES WHERE (CGC = '" & txtCGCCPF.text & "' OR CGC='" & Replace(Replace(Replace(txtCGCCPF.text, ".", ""), "-", ""), "/", "") & "')"
360         sql = sql & " AND FORNECEDOR<>" & txtCodigo.text
370         Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
380         If tbl.EOF = False Then
390            MsgBox "Este CNPJ já consta cadastrado para o fornecedor " & tbl!Descricao & "!", vbInformation
400            If txtCGCCPF.Enabled Then txtCGCCPF.SetFocus
410            Exit Sub
420         End If
430      End If

440      If Layout = 1 Then
450         If Val(TxtFORContaContabil.text) = 0 Then
460            MsgBox "A Conta deve ser informada!", vbInformation
470            TxtFORContaContabil.SetFocus
480            Exit Sub
490         End If
500      Else
510         TxtFORContaContabil.text = Trim(TxtFORContaContabil.text)
520      End If
         
530      If Inclui Then 'Se a operação for inclusão
         
540         If (Layout >= 9 Or Layout = 7) And Val(TxtFORContaContabil.text) <> 0 Then
            
550            sql = ""
560            sql = sql & " SELECT ISNULL(DESCRICAO,'') AS DESCRICAO FROM CONT_PLANOCONTA" & Chr(13)
570            sql = sql & " WHERE  CONTA = '" & Replace(TxtFORContaContabil.text, "'", "") & "'"
580            Set tblaux = Banco.OpenResultset(sql, rdOpenStatic)
590            If tblaux.EOF Then
600               sql = ""
610               sql = sql & " INSERT INTO CONT_PLANOCONTA (" & Chr(13)
620               sql = sql & "        CONTA, TIPOCONTA,DESCRICAO, ATUALIZACAO ) " & Chr(13)
630               sql = sql & " VALUES ( " & Chr(13)
640               sql = sql & "'" & Replace(TxtFORContaContabil.text, "'", "") & "'," & Chr(13)
650               sql = sql & "2," & Chr(13) ' PASSIVO
660               sql = sql & "'" & Replace(txtDescricao.text, "'", "") & "'," & Chr(13)
670               sql = sql & "'" & CorrenteTimeStamp & "')" & Chr(13)
680               Banco.Execute sql
690            Else
700               If MsgBox("Deseja associar o Forncedor a Conta: '" & tblaux("DESCRICAO") & "'?", vbQuestion + vbYesNo) = vbYes Then
710                  sql = ""
720                  sql = sql & " UPDATE CONT_PLANOCONTA SET DESCRICAO = '" & Replace(txtDescricao.text, "'", "") & "'" & Chr(13)
730                  sql = sql & " WHERE  CONTA = '" & Replace(TxtFORContaContabil.text, "'", "") & "'"
740                  Banco.Execute sql
750               End If
760            End If
770            tblaux.Close
               
780            sql = ""
790            sql = sql & " SELECT CONTA FROM CONT_FORNECEDORCONTA" & Chr(13)
800            sql = sql & " WHERE  CONTA = '" & Trim(Replace(TxtFORContaContabil.text, "'", "")) & "'" & Chr(13)
810            sql = sql & " AND    FORNECEDOR = " & txtCodigo.text
820            Set tblaux = Banco.OpenResultset(sql, rdOpenStatic)
830            If tblaux.EOF Then
840               sql = ""
850               sql = sql & " INSERT INTO CONT_FORNECEDORCONTA (" & Chr(13)
860               sql = sql & "        CONTA, FORNECEDOR, HISTORICO,TERCEIRO, " & Chr(13)
870               sql = sql & "        TIPOCONTA, PERCENTUAL, ATUALIZACAO ) " & Chr(13)
880               sql = sql & " VALUES ( " & Chr(13)
890               sql = sql & "'" & TxtFORContaContabil.text & "'," & Chr(13)
900               sql = sql & txtCodigo.text & "," & Chr(13)
910               sql = sql & IIf(Layout = 9, 2380, IIf(Layout = 11, 94, IIf(Layout = 10, 2, 1))) & "," & Chr(13) 'fixo Historico
920               sql = sql & 0 & "," & Chr(13)
930               sql = sql & 1 & "," & Chr(13)
940               sql = sql & 100 & "," & Chr(13)
950               sql = sql & "'" & CorrenteTimeStamp & "')" & Chr(13)
960               Banco.Execute sql
970            End If
980            tblaux.Close
               
990         End If
            
1000        If Not validarContaContabil Then
1010           If MsgBox("Conta Contábil informada não consta no plano de contas." & Chr(13) & Chr(13) & _
                      "Continuar mesmo assim?", vbQuestion + vbYesNo) = vbNo Then
1020              TxtFORContaContabil.SetFocus
1030              Exit Sub
1040           End If
1050        End If
            
1060        sql = " INSERT INTO fornecedores (fornecedor,descricao,endereco,bairro,cidade,cep, " & Chr(13)
1070        sql = sql & " inscricao,cgc,telefone,contacontabil,FAX,Email, Observacao, ATIVIDADE, ATIVIDADE_ISS, DESATIVADO, " & Chr(13)
1080        sql = sql & " FARMACIA, ALMOXARIFADO,SND, NOMEFANTASIA, BANCO, AGENCIA, CONTABANCO, "
1090        sql = sql & " TIPOPESSOA, DATAFUNDACAO, PRAZOENTREGA, VALORMINIMOPEDIDO, CONDICAOPAGAMENTO, RESPNEGOCIACAO, "
1100        sql = sql & " TEMPORARIO, CodigoTotvs ) "
1110        sql = sql & " VALUES ("
1120        sql = sql & TestaDados(Me.txtCodigo) & ","
1130        sql = sql & "'" & Replace(txtDescricao.text, "'", """") & "',"
1140        sql = sql & TestaDados(Me.txtFORendereco) & ","
1150        sql = sql & TestaDados(Me.txtFORbairro) & ","
1160        sql = sql & TestaDados(Me.txtCidade) & ","
1170        sql = sql & TestaDados(Me.mskFORcep) & ","
1180        sql = sql & TestaDados(Me.txtFORInscricao) & ","
1190        sql = sql & IIf(Trim(txtCGCCPF.text) = "", "''", "'" & txtCGCCPF.text & "'") & ","
1200        sql = sql & "'" & SoNumeros(mskFORtelefone.text) & "',"
1210        sql = sql & TestaDados(Me.TxtFORContaContabil) & ","
1220        sql = sql & "'" & acertaTelefone(mskFax.text) & "',"
1230        sql = sql & TestaDados(txtEmail) & ","
1240        sql = sql & TestaDados(txtObservacao) & ","
1250        sql = sql & UseCmbAtividade1.DataStorage & ","
1260        sql = sql & UseCmbAtividadeISS.DataStorage & ","
1270        sql = sql & chkDesativado.value & ","
1280        sql = sql & chkFarmacia.value & ","
1290        sql = sql & chkAlmoxarifado.value & ","
1300        sql = sql & ChkSND.value & ","
1310        sql = sql & "'" & IIf(Trim(txtNomeFantasia.text) = "", "", Trim(Replace(txtNomeFantasia.text, "'", """"))) & "',"
1320        sql = sql & "'" & IIf(Trim(txtBanco.text) = "", 0, Trim(txtBanco.text)) & "',"
1330        sql = sql & "'" & IIf(Trim(txtAgencia.text) = "", 0, Trim(txtAgencia.text)) & "',"
1340        sql = sql & "'" & IIf(Trim(TxtContaCorrente.text) = "", 0, Trim(TxtContaCorrente.text)) & "',"
            
1350        sql = sql & IIf(Trim(optPessoa(0).value) = 1, 1, 2) & ","
1360        sql = sql & IIf(Not IsDate(txtFundacao.text), "NULL", Format(Trim(txtFundacao.text), "'YYYY/MM/DD'")) & ","
1370        sql = sql & IIf(Trim(txtPrazoEntrega.text) = "", 0, Trim(txtPrazoEntrega.text)) & ","
1380        sql = sql & sqlFormataDecimal(IIf(Trim(txtValorMinimo.text) = "", 0, Trim(txtValorMinimo.text))) & ","
1390        sql = sql & "'" & IIf(Trim(txtCondPgto.text) = "", 0, Trim(txtCondPgto.text)) & "',"
1400        sql = sql & "'" & IIf(Trim(txtRespNegoc.text) = "", "", Trim(txtRespNegoc.text)) & "',"
            sql = sql & chkTemporario.value & ","
1410        sql = sql & txtCodigoTotvs.text & ")"
1420     Else 'Senão é Alteração
            
1430        If Not validarContaContabil Then
1440           If MsgBox("Conta Contábil informada não consta no plano de contas." & Chr(13) & Chr(13) & _
                      "Continuar mesmo assim?", vbQuestion + vbYesNo) = vbNo Then
1450              TxtFORContaContabil.SetFocus
1460              Exit Sub
1470           End If
1480        End If
            
1490        sql = "  UPDATE fornecedores SET " & Chr(13)
1500        If TelaPedidoCotacao = True And IsNumeric(TxtFORContaContabil.text) And Layout = 10 And chkTemporario.value = 0 Then sql = sql & " FORNECEDOR = " & TxtFORContaContabil.text & "," & Chr(13)
1510        sql = sql & "  descricao='" & Replace(txtDescricao.text, "'", """") & "'," & Chr(13)
1520        sql = sql & "  endereco=" & TestaDados(Me.txtFORendereco) & "," & Chr(13)
1530        sql = sql & "  bairro=" & TestaDados(Me.txtFORbairro) & "," & Chr(13)
1540        sql = sql & "  cidade=" & TestaDados(Me.txtCidade) & "," & Chr(13)
1550        sql = sql & "  cep=" & TestaDados(Me.mskFORcep) & "," & Chr(13)
1560        sql = sql & "  inscricao=" & TestaDados(Me.txtFORInscricao) & "," & Chr(13)
1570        sql = sql & "  cgc=" & IIf(Trim(txtCGCCPF.text) = "", "''", "'" & txtCGCCPF.text & "'") & "," & Chr(13)
1580        sql = sql & "  telefone='" & SoNumeros(mskFORtelefone.text) & "'," & Chr(13)
1590        sql = sql & "  contacontabil=" & TestaDados(Me.TxtFORContaContabil) & "," & Chr(13)
1600        sql = sql & "  FAX = '" & acertaTelefone(mskFax.text) & "'," & Chr(13)
1610        sql = sql & "  EMAIL = " & TestaDados(txtEmail) & "," & Chr(13)
1620        sql = sql & "  OBSERVACAO = " & TestaDados(txtObservacao) & "," & Chr(13)
1630        sql = sql & "  ATIVIDADE = " & UseCmbAtividade1.DataStorage & "," & Chr(13)
1640        sql = sql & "  ATIVIDADE_ISS = " & UseCmbAtividadeISS.DataStorage & "," & Chr(13)
1650        sql = sql & "  DESATIVADO = " & chkDesativado.value & "," & Chr(13)
1660        sql = sql & "  FARMACIA = " & chkFarmacia.value & "," & Chr(13)
1670        sql = sql & "  ALMOXARIFADO = " & chkAlmoxarifado.value & "," & Chr(13)
1680        sql = sql & "  SND = " & ChkSND.value & "," & Chr(13)
1690        sql = sql & "  NOMEFANTASIA = '" & IIf(Trim(txtNomeFantasia.text) = "", "", Trim(Replace(txtNomeFantasia.text, "'", """"))) & "'," & Chr(13)
1700        sql = sql & "  BANCO = '" & IIf(Trim(txtBanco.text) = "", 0, Trim(txtBanco.text)) & "'," & Chr(13)
1710        sql = sql & "  AGENCIA = '" & IIf(Trim(txtAgencia.text) = "", 0, Trim(txtAgencia.text)) & "'," & Chr(13)
1720        sql = sql & "  CONTABANCO = '" & IIf(Trim(TxtContaCorrente.text) = "", 0, Trim(TxtContaCorrente.text)) & "'," & Chr(13)
            
1730        sql = sql & "  TIPOPESSOA = " & IIf(optPessoa(0).value = True, 1, 2) & "," & Chr(13)
1740        sql = sql & "  DATAFUNDACAO = " & IIf(Not IsDate(txtFundacao.text), "NULL", Format(Trim(txtFundacao.text), "'YYYY/MM/DD'")) & "," & Chr(13)
1750        sql = sql & "  PRAZOENTREGA = " & IIf(Trim(txtPrazoEntrega.text) = "", 0, Trim(txtPrazoEntrega.text)) & "," & Chr(13)
1760        sql = sql & "  VALORMINIMOPEDIDO = " & sqlFormataDecimal(IIf(Trim(txtValorMinimo.text) = "", 0, Trim(txtValorMinimo.text))) & "," & Chr(13)
1770        sql = sql & "  CONDICAOPAGAMENTO = '" & IIf(Trim(txtCondPgto.text) = "", 0, Trim(txtCondPgto.text)) & "'," & Chr(13)
1780        sql = sql & "  RESPNEGOCIACAO = '" & IIf(Trim(txtRespNegoc.text) = "", "", Trim(txtRespNegoc.text)) & "'," & Chr(13)
            
1790        sql = sql & "  TEMPORARIO = " & chkTemporario.value & "," & Chr(13)
1800        sql = sql & "  PERC_ISS = " & IIf(Not IsNumeric(txtISS.text), 0, Replace(txtISS.text, ",", ".")) & "," & Chr(13)
            sql = sql & "  CodigoTotvs = " & txtCodigoTotvs.text
1810        sql = sql & "  WHERE convert(int,Trava) = " & Trava & "" & Chr(13)
1820        sql = sql & "  AND fornecedor = " & txtCodigo.text & Chr(13)
1830     End If
1840     SV.Execute sql, Afetado 'Executa a SQL de Inclusão ou Alteração
         
         'Relaciona o Fornecedor ao tipo de produto
1850     With grdProdutos
              
1860        sql = ""
1870        sql = sql & "DELETE FROM FORNPRODTIPO " & vbCrLf
1880        sql = sql & "WHERE FORNECEDOR = " & Val(txtCodigo.text)
1890        Banco.Execute sql
              
1900        For i = 0 To (.Rows - 1)

1910            If .TextMatrix(i, 2) = "X" Then
1920                      sql = ""
1930                      sql = "INSERT INTO FORNPRODTIPO(FORNECEDOR,PRODUTOTIPO) " & vbCrLf
1940                      sql = sql & " VALUES (" & Val(txtCodigo.text) & "," & .TextMatrix(i, 0) & ") " & vbCrLf
1950                      Banco.Execute sql
1960            End If
1970        Next i
1980        .Tag = 0
1990     End With
         
2000     If Not Inclui Then 'Se a operação for inclusão
2010        If TelaPedidoCotacao = True And IsNumeric(TxtFORContaContabil.text) And Layout = 10 And chkTemporario.value = 0 Then
2020           sql = "UPDATE TMPCOTACAO1 SET FORNECEDOR=" & TxtFORContaContabil.text
2030           sql = sql & " WHERE FORNECEDOR=" & txtCodigo.text & Chr(13)
2040           Banco.Execute sql
2050        End If
2060     End If
2070     If Afetado = 0 Then 'Testa se a Execução da SQL deu certo
2080        MsgBox "Ocorreu um erro interno na tentativa da gravação! Tente novamente!", vbInformation, Caption
2090     End If
2100     LimpaCamposPublico Me
2110     chkFarmacia.value = 0: chkAlmoxarifado.value = 0: ChkSND.value = 0
2120     TelaPedidoCotacao = False
         
SaiSub:
2130   cmdCancelar_Click
       
2140     If TelaPedidoCotacao Then Unload Me

2150  Exit Sub
Erro:
2160      TratarErro "frmFornecedores", "cmdSalvar_Click", Err.Number, Err.Description, Erl
End Sub

Private Sub CmdSalvar1_Click()
       Dim sql As String, Afetado As Integer
       
         'If Not PreencheuTodos(Me) Then 'Testa se todos os campos obrigatórios foram preenchidos
         '      MsgBox "Preencha os campos em vermelho!", vbInformation, Caption
         '      Exit Sub
         'End If
         
On Error GoTo Erro
20       If Not IsNumeric(TxtCodigo1.text) Then
30          MsgBox "O Codigo deve ser informado!", vbInformation
40          txtCodigo.SetFocus
50          Exit Sub
60       End If
         
70       If Trim(TxtDescricao1.text) = "" Then
80          MsgBox "A Descrição deve ser informada!", vbInformation
90          txtDescricao.SetFocus
100         Exit Sub
110      End If
         
         'If Not Trim(TxtFORContaContabil.Text) = "" Then
         '   MsgBox "A Conta deve ser informada!", vbInformation
         '   TxtFORContaContabil.SetFocus
         '   Exit Sub
         'End If
         
120      If Not Corretos(Me) Then 'Testa se os campos preenchidos estão com os dados corretos
130            MsgBox "Corrija os campos em vermelho!", vbInformation, Caption
140            Exit Sub
150      End If
160      If Inclui1 Then 'Se a operação for inclusão
170      sql = "INSERT INTO fabricantes (fabricante,descricao,endereco,bairro,cidade,cep,inscricao,cgc,telefone,DisFab,fax,email,contato,ATIVIDADE, DESATIVADO)" & _
               "VALUES (" & TestaDados(Me.TxtCodigo1) & "," & TestaDados(Me.TxtDescricao1) & "," & _
               TestaDados(Me.txtFABendereco) & "," & TestaDados(Me.txtFABbairro) & "," & _
               TestaDados(Me.TxtCidade1) & "," & TestaDados(Me.mskFABcep) & "," & TestaDados(Me.txtFABInscricao) & "," & _
               TestaDados(Me.mskFABcgc) & ",'" & SoNumeros(mskFABtelefone.text) & "'," & chkDist.value & ",'" & acertaTelefone(txtFax.text) & "'," & TestaDados(txtEmail1) & "," & TestaDados(txtContato) & "," & _
               UseCmbAtividade.DataStorage & "," & ChkDesativado1.value & " ) "
180    Else 'Senão é Alteração
190      sql = "UPDATE fabricantes SET descricao=" & TestaDados(Me.TxtDescricao1) & ",endereco=" & _
               TestaDados(Me.txtFABendereco) & ",bairro=" & TestaDados(Me.txtFABbairro) & ",cidade=" & _
               TestaDados(Me.TxtCidade1) & ",cep=" & TestaDados(Me.mskFABcep) & ",inscricao=" & _
               TestaDados(Me.txtFABInscricao) & ",cgc=" & TestaDados(Me.mskFABcgc) & ",telefone='" & _
               SoNumeros(mskFABtelefone.text) & "',fax = '" & acertaTelefone(txtFax.text) & "',email = " & TestaDados(txtEmail1) & ",contato = " & TestaDados(txtContato) & _
               ",ATIVIDADE = " & UseCmbAtividade.DataStorage & ",DESATIVADO = " & ChkDesativado1.value & _
               ",DisFab = " & chkDist.value & " WHERE convert(int,Trava) = " & Trava1 & " AND fabricante = " & TxtCodigo1.text
200    End If
210    SV.Execute sql, Afetado 'Executa a SQL de Inclusão ou Alteração
220    If Afetado = 0 Then 'Testa se a Execução da SQL deu certo
230      MsgBox "Ocorreu um erro interno na tentativa da gravação! Tente novamente!", vbInformation, Caption
240    End If
250    LimpaCamposPublico Me
SaiSub:
260    CmdCancelar1_Click

Exit Sub
Erro:
    TratarErro "frmFornecedores", "CmdSalvar1_Click", Err.Number, Err.Description, Erl
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Erro
10     If KeyAscii = 13 Then 'Força o evento [TAB], com a tecla [ENTER]
20       SendKeys "{TAB}"
30       KeyAscii = 0
40     End If
       
50     fraProdutos.Visible = False

Exit Sub
Erro:
    TratarErro "frmFornecedores", "Form_KeyPress", Err.Number, Err.Description, Erl
       
End Sub

Private Sub Form_Load()
On Error GoTo Erro
20        UseCmbAtividade.CorFundo = &HEEEEEE
30        UseCmbAtividade1.CorFundo = &HEEEEEE
40        Me.Top = 0
50        Me.Left = 0
60        TxtDescricao1.Enabled = False 'Descricao ou próximo campo
70        FmeCampos1.Enabled = False
80        LinhaVelha1 = 0 'Guarda Linha Atual da Grid1
90        Inclui1 = True
100       Pode1 = True
110       Grid1.ColAlignment(0) = 4
120       SSTab1.TabEnabled(0) = False
130       txtDescricao.Enabled = False 'Descricao ou próximo campo
140       FmeCampos.Enabled = False
150       LinhaVelha = 0 'Guarda Linha Atual da Grid
160       Inclui = True
170       Pode = True
180       Grid.ColAlignment(0) = 4
190       CarregaQuery
200       PosicionaForm Me, 1

          'Se for o layout Dracena:
          'Haverá a possibilidade do usuário atribuir ao Fornecedor da base do Save
          'o código do fornecedor usado na base da Totvs, facilitando então o processo
          'de exportação
          If Layout = 53 Then
            lblCodigoTotvs.Visible = True
            txtCodigoTotvs.Visible = True
          End If
          
210       InicializaGridRelacao

220       If TelaPedidoCotacao = True Then
230         txtDescricao.TabIndex = 0
240         FraTipoFornecedor.Enabled = False
250         chkAlmoxarifado.value = 1: ChkSND.value = 1: chkFarmacia.value = 1
260         txtCGCCPF.TabIndex = 1
270         cmdSalvar.TabIndex = 2
280         chkTemporario.value = 1
290         chkTemporario.Visible = False: chkDesativado.Visible = False
300         chkTemporario.Enabled = False
310         cmdCancelar.Enabled = False
320         cmdExcluir.Enabled = False
330         stbPadrao.TabEnabled(1) = False
            
340         TxtCodigo_KeyPress 13
            'txtDescricao.SetFocus
350       End If

Exit Sub
Erro:
    TratarErro "frmFornecedores", "Form_Load", Err.Number, Err.Description, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)
10    On Error GoTo Erro
         Dim Indice As Long
         Dim Aberto As Boolean
20       For Indice = 0 To Forms.Count - 1
30          If UCase(Forms(Indice).Name) = "FRMPEDIDOSCOTACAO" Then
40             Aberto = True
50             Exit For
60          End If
70       Next
80       If TelaPedidoCotacao = True And Aberto Then
90          FrmPedidosCotacao.PedidoCarregaListaFornecedores
100         FrmPedidosCotacao.lstFornecedor.SetFocus
110      End If

120   Exit Sub
Erro:
130      TratarErro
End Sub

Private Sub grdProdutos_Click()

10            With grdProdutos
20                If .TextMatrix(.RowSel, .ColSel) = "" Then
30                    .TextMatrix(.RowSel, .ColSel) = "X"
40                Else
50                    .TextMatrix(.RowSel, .ColSel) = ""
60                End If
70            End With
              
End Sub

Private Sub grdRelacao_AlterarDados()
10       MsgBox "Não é possível realizar a alteração, exclua e lance novamente", vbInformation
End Sub

Private Sub grdRelacao_ExcluirDados()
10       On Error GoTo Erro
         
20       If Val(grdRelacao.TextMatrix(grdRelacao.RowSel, 0)) = 0 Then Exit Sub
         
30       If MsgBox("Confirma a exclusão da relação para esse fornecedor?", vbInformation + vbYesNo) = vbNo Then
40          Exit Sub
50       End If
         
60       sql = ""
70       sql = sql & " DELETE FROM FORNECEDOR_RELACAO " & Chr(13)
80       sql = sql & " WHERE FORNECEDOR_RELACAO = " & grdRelacao.TextMatrix(grdRelacao.RowSel, 0) & Chr(13)
90       sql = sql & " AND   FORNECEDOR = " & txtCodigo.text & Chr(13)
100      Banco.Execute sql
         
110      MontarGridRelacao
         
120      Exit Sub
Erro:
130      TratarErro
End Sub

Private Sub GRID_Click()
10    On Error GoTo Erro
      Dim i As Long, Atulinha As Long, Col As Long
20    Pode = False
       
30        If Grid.MouseRow = 0 Then
40             Grid.Sort = 1
50        ElseIf Grid.Rows > 1 Then
60          If Grid.Row <> LinhaVelha Then
70             Atulinha = Grid.Row
80             Col = Grid.Col
90             If LinhaVelha <> 0 Then
100               Grid.Row = LinhaVelha
110               Grid.TextMatrix(Grid.Row, 0) = "" 'tira o * da linha anteiror
120               For i = 1 To Grid.Cols - 1 'Volta a cor normal da linha
130                  Grid.Col = i
140                  Grid.CellForeColor = &H80000008
150                  Grid.CellBackColor = &H80000005
160               Next
170            End If
180            Grid.Row = Atulinha
190            Grid.TextMatrix(Grid.Row, 0) = "*" 'Coloca o * na Linha Atual
200            Grid.Col = 0
210            Grid.CellFontBold = True
220            LinhaVelha = Atulinha
230            For i = 1 To Grid.Cols - 1 'Seleciona a Linha Atual
240              Grid.Col = i
250              Grid.CellForeColor = &H80000005
260              Grid.CellBackColor = &H8000000D
270            Next
280            Grid.Col = Col
290         End If
300       End If
310       Pode = True

320   Exit Sub
Erro:
330      TratarErro
          
End Sub

Private Sub GRID_DblClick()
10    On Error GoTo Erro
20       If Grid.Rows > 1 Then
30         stbPadrao.Tab = 0
40         txtCodigo = Grid.TextMatrix(Grid.Row, 1)
50         TxtCodigo_KeyPress 13
60         txtCodigo = Grid.TextMatrix(Grid.Row, 1)
70       End If

80    Exit Sub
Erro:
90       TratarErro

End Sub

Private Sub Grid_GotFocus()
10          Me.KeyPreview = False
End Sub

Private Sub Grid_KeyUp(KeyCode As Integer, Shift As Integer)
10       On Error GoTo Erro
            
         Dim tbl As rdoResultset
         
20       sql = ""
30       sql = sql & " SELECT CREDOR FROM FILANCAMENTO"
40       sql = sql & " WHERE TERCEIROS = 0 "
50       sql = sql & " AND   CREDOR = " & Val(txtCodigo.text)
         
60       Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
         
70       If tbl.EOF = False Then
80          MsgBox "Este fornecedor não pode ser excluído, pois há lançamentos de notas no financeiro!", vbInformation
90          Exit Sub
100      End If
         
110      If KeyCode = vbKeyDelete Then
120        If Grid.Rows > 1 Then
130          If MsgBox("Confirma a Exclusão do registro selecionado?", vbQuestion + vbYesNo, Caption) = vbYes Then
140            Banco.Execute "DELETE FROM fornecedores WHERE fornecedor = " & Grid.TextMatrix(Grid.Row, 1)
150            If Grid.Rows > 2 Then
160              Grid.RemoveItem Grid.Row
170            Else
180              Grid.Rows = 1
190            End If
200          End If
210        End If
220      End If

230      Exit Sub
Erro:
240      TratarErro
End Sub

Private Sub Grid_LostFocus()
10        Me.KeyPreview = True
End Sub

Private Sub Grid_RowColChange()
10       If Pode Then GRID_Click
End Sub

Private Sub GRID1_Click()
10    On Error GoTo Erro
      Dim i As Long, Atulinha As Long, Col As Long
       
20     Pode1 = False
30     If Grid1.MouseRow = 0 Then
40       Grid1.Sort = 1
50     ElseIf Grid1.Rows > 1 Then
60       If Grid1.Row <> LinhaVelha1 Then
70         Atulinha = Grid1.Row
80         Col = Grid1.Col
90         If LinhaVelha1 <> 0 Then
100          Grid1.Row = LinhaVelha1
110          Grid1.TextMatrix(Grid1.Row, 0) = "" 'tira o * da linha anteiror
120          For i = 1 To Grid1.Cols - 1 'Volta a cor normal da linha
130            Grid1.Col = i
140            Grid1.CellForeColor = &H80000008
150            Grid1.CellBackColor = &H80000005
160          Next
170        End If
180        Grid1.Row = Atulinha
190        Grid1.TextMatrix(Grid1.Row, 0) = "*" 'Coloca o * na Linha Atual
200        Grid1.Col = 0
210        Grid1.CellFontBold = True
220        LinhaVelha1 = Atulinha
230        For i = 1 To Grid1.Cols - 1 'Seleciona a Linha Atual
240          Grid1.Col = i
250          Grid1.CellForeColor = &H80000005
260          Grid1.CellBackColor = &H8000000D
270        Next
280        Grid1.Col = Col
290      End If
300    End If
310    Pode1 = True

320   Exit Sub
Erro:
330      TratarErro
End Sub

Private Sub GRID1_DblClick()

10       If Grid1.Rows > 1 Then
20          stbPadrao1.Tab = 0
30          TxtCodigo1 = Grid1.TextMatrix(Grid1.Row, 1)
40          TxtCodigo1_KeyPress 13
50          TxtCodigo1 = Grid1.TextMatrix(Grid1.Row, 1)
60       End If

End Sub

Private Sub Grid1_GotFocus()
10     Me.KeyPreview = False
End Sub

Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
10     If KeyCode = vbKeyDelete Then
20       If Grid1.Rows > 1 Then
30         If MsgBox("Confirma a Exclusão do registro selecionado?", vbQuestion + vbYesNo, Caption) = vbYes Then
40           SV.Execute "DELETE FROM (Tabela) WHERE Campo = " & Grid1.TextMatrix(Grid1.Row, 1)
50           If Grid1.Rows > 2 Then
60             Grid1.RemoveItem Grid1.Row
70           Else
80             Grid1.Rows = 1
90           End If
100        End If
110      End If
120    End If
End Sub

Private Sub Grid1_LostFocus()
10     Me.KeyPreview = True
End Sub

Private Sub Grid1_RowColChange()
10     If Pode1 Then GRID1_Click
End Sub

Private Sub InicializaGridProdutos()
10       On Error GoTo Erro
         
20       With grdProdutos
30          .GrdZeraColecaoTitulo
40          .GrdRowsTitulo = 1
50          .GrdAddTitulo "Código", 600, flexAlignLeftCenter, flexAlignLeftCenter
60          .GrdAddTitulo "Tipo de Produto", 3000, flexAlignLeftCenter, flexAlignLeftCenter
70          .GrdAddTitulo "Associado", 850, flexAlignCenterCenter, flexAlignCenterCenter
80          .GrdInicializa
90          .Col = 0
            
100          Set .CellPicture = Nothing
110      End With
         
120      Exit Sub
Erro:
130      TratarErro Me.Name, "InicializaGridProdutos", Err.Number, Err.Description, Erl
End Sub

Private Sub InicializaGridRelacao()
10       On Error GoTo Erro
         
20       With grdRelacao
30          .GrdZeraColecaoTitulo
40          .GrdRowsTitulo = 1
50          .GrdAddTitulo "Fornecedor", 0, flexAlignLeftCenter, flexAlignLeftCenter
60          .GrdAddTitulo "Descrição", 4000, flexAlignLeftCenter, flexAlignLeftCenter
70          .GrdAddTitulo "CNPJ ", 2000, flexAlignLeftCenter, flexAlignLeftCenter

80          .GrdInicializa
90          .Col = 0
            
100          Set .CellPicture = Nothing
110      End With
         
120      Exit Sub
Erro:
130      TratarErro Me.Name, "Execucutado", Err.Number, Err.Description, Erl
End Sub

Private Sub Lista_DblClick()

10       If Lista.ListIndex <> -1 Then
20          Me.Controls(Vcod).text = Lista.ItemData(Lista.ListIndex)
30          Me.Controls(Vtex).text = Lista.List(Lista.ListIndex)
40          Lista.Visible = False
50          Me.Controls(Vtex).SetFocus
60       End If

End Sub

Private Sub Lista_GotFocus()
10     Me.KeyPreview = False
End Sub

Private Sub lista_KeyPress(KeyAscii As Integer)
10     If KeyAscii = 27 Then
20       Lista.Visible = False
30       Me.Controls(Vtex).SetFocus
40     ElseIf KeyAscii = 13 Then
50       Lista_DblClick
60     End If
End Sub

Private Sub Lista_LostFocus()
10     Me.KeyPreview = True
20     Lista.Visible = False
End Sub

Private Sub Lista1_DblClick()
10    On Error GoTo Erro
       
20        If Lista1.ListIndex <> -1 Then
30             Me.Controls(Vcod).text = Lista1.ItemData(Lista1.ListIndex)
40             Me.Controls(Vtex).text = Lista1.List(Lista1.ListIndex)
50             Lista1.Visible = False
60             Me.Controls(Vtex).SetFocus
70        End If
          
80    Exit Sub
Erro:
90       TratarErro
End Sub

Private Sub Lista1_GotFocus()
10     Me.KeyPreview = False
End Sub

Private Sub lista1_KeyPress(KeyAscii As Integer)
10     If KeyAscii = 27 Then
20       Lista1.Visible = False
30       Me.Controls(Vtex).SetFocus
40     ElseIf KeyAscii = 13 Then
50       Lista1_DblClick
60     End If
End Sub

Private Sub Lista1_LostFocus()
10     Me.KeyPreview = True
20     Lista1.Visible = False
End Sub

Private Sub MontarGridProdutos()
10    On Error GoTo Erro
         
20       SetAmpulheta
         
         Dim sql As String
         Dim tbl As rdoResultset
         Dim linha As Integer
         
30       InicializaGridProdutos
         
40        sql = ""
50        sql = sql & "SELECT A.CODIGO, " & vbCrLf
60        sql = sql & "       A.TIPO, " & vbCrLf
70        sql = sql & "       ( CASE " & vbCrLf
80        sql = sql & "           WHEN ISNULL(B.FORNECEDOR, 0) = 0 THEN 0 " & vbCrLf
90        sql = sql & "           ELSE 1 " & vbCrLf
100       sql = sql & "         END ) AS EXISTE " & vbCrLf
110       sql = sql & "FROM   PRODUTOTIPO A " & vbCrLf
120       sql = sql & "       LEFT JOIN FORNPRODTIPO B " & vbCrLf
130       sql = sql & "              ON A.CODIGO = B.PRODUTOTIPO " & vbCrLf
140       sql = sql & "                 AND B.FORNECEDOR = " & Val(txtCodigo.text) & vbCrLf
150       sql = sql & "WHERE  1 = 1 " & vbCrLf
160       sql = sql & "       AND ISNULL(A.DESATIVADO, 0) = 0 " & vbCrLf
170       sql = sql & "ORDER  BY A.TIPO "

180      Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
         
190      grdProdutos.GrdLimpar grdProdutos.GrdRowsTitulo
         
200      If Not tbl.EOF Then
            
210         tbl.MoveLast
220         tbl.MoveFirst
            
230         With grdProdutos
                   
240            If tbl.RowCount > .Rows - 1 Then
250              .Rows = tbl.RowCount + 1
260              .GrdTarjar 1
270            End If
               
280            linha = 1
            
290            Do Until tbl.EOF

300               .RowData(linha) = tbl!Codigo
                  
310               .TextMatrix(linha, 0) = tbl!Codigo
320               .TextMatrix(linha, 1) = tbl!Tipo
330               .TextMatrix(linha, 2) = IIf(tbl!Existe = 1, "X", "")
                  
340               tbl.MoveNext
350               linha = linha + 1
               
360            Loop
               
370            .Row = 1
380            .Col = 0
               .Tag = 1
390         End With
400      Else
410         MsgBox "Não há tipos de produto cadastrados."
420      End If
         
430      tbl.Close
440   reSetAmpulheta
450   Exit Sub

Erro:
460       reSetAmpulheta
470      TratarErro Me.Name, "MontarGridProdutos", Err.Number, Err.Description, Erl
End Sub

   
   
   
   
Private Sub MontarGridRelacao()
10       On Error GoTo Erro
         
20       SetAmpulheta
         
         Dim sql As String
         Dim tbl As rdoResultset
         Dim linha As Integer

30       InicializaGridRelacao
             
40       sql = ""
50       sql = sql & " SELECT A.FORNECEDOR_RELACAO, B.DESCRICAO, B.CGC " & Chr(13)
60       sql = sql & " FROM FORNECEDOR_RELACAO A INNER JOIN FORNECEDORES B ON A.FORNECEDOR_RELACAO = B.FORNECEDOR " & Chr(13)
70       sql = sql & " WHERE A.FORNECEDOR = " & txtCodigo.text & Chr(13)
80       sql = sql & " ORDER BY DESCRICAO" & Chr(13)
90       Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
         
100      grdRelacao.GrdLimpar grdRelacao.GrdRowsTitulo
         
110      If Not tbl.EOF Then
            
120         tbl.MoveLast
130         tbl.MoveFirst
            
140         With grdRelacao
                   
150            If tbl.RowCount > .Rows - 1 Then
160              .Rows = tbl.RowCount + 1
170              .GrdTarjar 1
180            End If
               
190            linha = 1
            
200            Do Until tbl.EOF

210               .RowData(linha) = tbl!FORNECEDOR_RELACAO
                  
220               .TextMatrix(linha, 0) = tbl!FORNECEDOR_RELACAO
230               .TextMatrix(linha, 1) = tbl!Descricao
240               .TextMatrix(linha, 2) = IIf(IsNull(tbl!CGC), "", tbl("CGC"))
                  
250               tbl.MoveNext
260               linha = linha + 1
               
270            Loop
               
280            .Row = 1
290            .Col = 0
300         End With
310      End If
         
320      tbl.Close
330      reSetAmpulheta
         
340      Exit Sub
Erro:
350      TratarErro Me.Name, "Execucutado", Err.Number, Err.Description, Erl
End Sub

Private Sub mskFax_Validate(Cancel As Boolean)
10       validateTelefone mskFax, Cancel
End Sub

Private Sub mskFORcep_GotFocus()
10     MarcaTexto mskFORcep
20     Me.KeyPreview = False
End Sub

Private Sub mskFORcep_KeyPress(KeyAscii As Integer)
10       On Error GoTo Erro
         
         Dim rs As New Recordset, sql As String
         
20       If KeyAscii = 13 Then
30          If mskFORcep.ClipText <> "" Then
40             sql = "SELECT * FROM Cidades WHERE CEP = '" & mskFORcep.text & "'"
50             rs.Open sql, SV
60             If Not rs.EOF Then
70                txtCidade = rs!Cidade
80                TxtNomeCidade = rs!Descricao
90             End If
100            rs.Close
110        End If
120        SendKeys "{TAB}"
130      End If

140      Exit Sub
Erro:
150      TratarErro
End Sub

Private Sub mskFORcep_LostFocus()
10     Me.KeyPreview = True
End Sub

'remove caracteres de uma string
Public Function Retira(Alvo As String, OQue As String, como As Integer) As String
10    On Error GoTo Erro
  Dim X As String, K As String, i As Integer, _
      j As Integer, p As Integer                       'dimensiona
20     If como = UM_A_UM Then                               'se um a um
30       X$ = ""                                            'vamos concatenar em x
40       For i = 1 To Len(Alvo$)                            'cada caracter que
50         K$ = Mid$(Alvo$, i, 1)                           'não estiver
60         If InStr(OQue$, K$) = 0 Then X$ = X$ + K$        'contido na string a regirar
70       Next
80     Else                                                 'se não for um a um
90       X$ = Alvo$                                         'vamos tirar
100      p = InStr(X$, OQue$)                               'toda a string
110      If p > 0 Then                                      'de uma só vez
120        X$ = Left$(X$, p - 1) + Mid$(X$, p + Len(OQue$)) 'da string alvo
130      End If
140    End If
150    Retira$ = X$                                         'retorna nova string
160   Exit Function
Erro:
170      TratarErro "ModuloGeral", "Retira", Err.Number, Err.Description, Erl
End Function

Private Sub mskFORtelefone_Validate(Cancel As Boolean)
   validateTelefone mskFORtelefone, Cancel
End Sub

Private Sub optPessoa_Click(Index As Integer)
10    On Error GoTo Erro

20        If Index = 0 Then
30            lblFornecedores(8).Caption = "CPF"
40            txtCGCCPF.text = saveFormataCPF(txtCGCCPF.text)
50        Else
60            lblFornecedores(8).Caption = "CNPJ"
70            txtCGCCPF.text = saveFormataCNPJ(txtCGCCPF.text)
80        End If
90        txtCGCCPF.SelStart = txtCGCCPF.MaxLength
100   Exit Sub
Erro:
110       TratarErro "frmFornecedores", "optPessoa_Click", Err.Number, Err.Description, Erl
          
End Sub

Private Sub stbPadrao_Click(PreviousTab As Integer)
10       On Error GoTo Erro
         
         'Ao selecionar esta guia (Inclusão/Alteração) - Focaliza o primeiro campo!
20       If stbPadrao.Tab = 1 Then
30          Grid.Rows = 1
40          cmbCampos.SetFocus
50       ElseIf stbPadrao.Tab = 0 And Inclui Then
60          If txtCodigo.Enabled Then txtCodigo.SetFocus
70       End If
         
80       Exit Sub
Erro:
90       TratarErro
End Sub

Private Sub stbPadrao_GotFocus()
10    On Error GoTo Erro
20     If stbPadrao.Tab = 1 Then
30       Grid.Rows = 1
40       cmbCampos.SetFocus
50     ElseIf stbPadrao.Tab = 0 And Inclui Then
60          If txtCodigo.Enabled Then
                txtCodigo.SetFocus
            Else
                txtDescricao.SetFocus
            End If
70     End If

80    Exit Sub
Erro:
90       TratarErro
End Sub

Private Sub stbPadrao1_Click(PreviousTab As Integer)
10     On Error Resume Next
        'Ao selecionar esta guia (Inclusão/Alteração) - Focaliza o primeiro campo!
20     If stbPadrao1.Tab = 1 Then
30       Grid1.Rows = 1
40       cmbCampos1.SetFocus
50     ElseIf stbPadrao1.Tab = 0 And Inclui1 Then
60       TxtCodigo1.SetFocus
70     End If
End Sub

Private Sub stbPadrao1_GotFocus()
10     If stbPadrao1.Tab = 1 Then
20       Grid1.Rows = 1
30       cmbCampos1.SetFocus
40     ElseIf stbPadrao1.Tab = 0 And Inclui1 Then
50       TxtCodigo1.SetFocus
60     End If
End Sub

Private Sub txtCGCCPF_GotFocus()
10        If optPessoa(0).value = True Then
20            txtCGCCPF.text = saveFormataCPF(txtCGCCPF.text)
30        Else
40            txtCGCCPF.text = saveFormataCNPJ(txtCGCCPF.text)
50        End If
          
60        txtCGCCPF.SelStart = txtCGCCPF.MaxLength
          
End Sub

Private Sub txtCGCCPF_KeyUp(KeyCode As Integer, Shift As Integer)
10        If optPessoa(0).value = True Then
20            txtCGCCPF.text = saveFormataCPF(txtCGCCPF.text)
30        Else
40            txtCGCCPF.text = saveFormataCNPJ(txtCGCCPF.text)
50        End If
60        txtCGCCPF.SelStart = txtCGCCPF.MaxLength
End Sub

Private Sub txtCGCCPF_LostFocus()
10    On Error GoTo Erro
Dim Tamanho As Byte
Dim Copia As String

20    If Not txtCGCCPF.Vazio And 1 = 2 Then
30       Tamanho = Len(txtCGCCPF.text)
40       If Tamanho = 14 And IsNumeric(txtCGCCPF.text) Then
50          Copia = txtCGCCPF.text
60          txtCGCCPF.text = Mid(Copia, 1, 2) & "." _
                        & Mid(Copia, 3, 3) & "." _
                        & Mid(Copia, 6, 3) & "/" _
                        & Mid(Copia, 9, 4) & "-" _
                        & Mid(Copia, 13, 2)
70          If VCGC(txtCGCCPF.text) <> -1 Then
80             MsgBox "CNPJ inválido!", vbInformation
90             txtCGCCPF.SetFocus
100            Exit Sub
110         End If
120      ElseIf Tamanho = 18 And IsNumeric(Mid(txtCGCCPF.text, 1, 2)) _
                             And IsNumeric(Mid(txtCGCCPF.text, 4, 3)) _
                             And IsNumeric(Mid(txtCGCCPF.text, 8, 3)) _
                             And IsNumeric(Mid(txtCGCCPF.text, 12, 4)) _
                             And IsNumeric(Mid(txtCGCCPF.text, 17, 2)) Then
130         If VCGC(txtCGCCPF.text) <> -1 Then
140            MsgBox "CNPJ inválido", vbInformation
               'txtCGCCPF.SetFocus
150            Exit Sub
160         End If
170      ElseIf Tamanho = 11 Then
180         Copia = txtCGCCPF.text
190         txtCGCCPF.text = Mid(Copia, 1, 3) & "." _
                        & Mid(Copia, 4, 3) & "." _
                        & Mid(Copia, 7, 3) & "-" _
                        & Mid(Copia, 10, 2)
200         If VDV2(txtCGCCPF.text) <> -1 Then
210            MsgBox "CPF inválido!", vbInformation
220            txtCGCCPF.SetFocus
230            Exit Sub
240         End If
250      ElseIf Tamanho = 14 And Not IsNumeric(txtCGCCPF.text) Then
260         If VDV2(txtCGCCPF.text) <> -1 Then
270            MsgBox "CPF inválido!", vbInformation
               'txtCGCCPF.SetFocus
280            Exit Sub
290         End If
300      Else
310         MsgBox "CNPJ inválido!", vbInformation
320         txtCGCCPF.SetFocus
330      End If
340   End If

350   Exit Sub
Erro:
360      TratarErro "frmEmpresas", "txtCGC_LostFocus", Err.Number, Err.Description, Erl
End Sub

Private Sub txtCidade_GotFocus()
10    On Error GoTo Erro
20     MarcaTexto txtCidade
30     Me.KeyPreview = False
40    Exit Sub
Erro:
50        TratarErro "frmFornecedores", "txtCidade_GotFocus", Err.Number, Err.Description, Erl
End Sub

Private Sub TxtCidade_KeyPress(KeyAscii As Integer)
       Dim rs As New Recordset, sql As String
10    On Error GoTo Erro

20     If KeyAscii = 13 Then
30       KeyAscii = 0
40       If Trim(txtCidade) <> "" Then
50         sql = "SELECT * FROM CIDADES WHERE Cidade = " & txtCidade
60         rs.Open sql, SV
70         If rs.EOF Then
80           MsgBox "Cidade não Cadastrada!", vbInformation, Caption
90           rs.Close
100          Exit Sub
110        End If
120        txtCidade = rs!Cidade
130        TxtNomeCidade = rs!Descricao
140        mskFORcep.text = IIf(rs!CEP <> "", rs!CEP, "_____-___")
150        rs.Close
160      End If
170      SendKeys "{TAB}"
180    End If

190   Exit Sub
Erro:
200       TratarErro "frmFornecedores", "TxtCidade_KeyPress", Err.Number, Err.Description, Erl
End Sub

Private Sub TxtCidade1_Change()
10    If TxtCidade1.text <> "" Then
20       ProcuraValor "select * from cidades where cidade=" & TxtCidade1.text, "Descricao", _
         Me.TxtCidade1, Me.TxtNomeCidade1, "Cidade não cadastrada!"
30    End If
End Sub

Private Sub TxtCidade1_KeyPress(KeyAscii As Integer)
10       KeyAscii = VerificaCampo(Me.TxtCidade1, KeyAscii, tipInteiro, 1, 999999)
End Sub

Private Sub TxtCodigo_GotFocus()
10     Me.KeyPreview = False
20     MarcaTexto Me.txtCodigo
End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)
10    On Error GoTo Erro
      Dim rs As New Recordset, sql As String
       
20      KeyAscii = VerificaCampo(Me.txtCodigo, KeyAscii, tipInteiro, 0, 999999999)
30      If KeyAscii = 13 And Trim(txtCodigo) <> "" And IsNumeric(txtCodigo.text) Then
40          sql = "SELECT *,convert(int,Trava) as Travar FROM fornecedores WHERE fornecedor= " & txtCodigo.text
50          rs.Open sql, SV
60          txtCodigo.Enabled = False
70          If Not rs.EOF Then
80              Inclui = False
90              Trava = rs!Travar
100             cmdExcluir.Enabled = True
110             MostraCamposPublico Me, rs
120             UseCmbAtividade1.DataStorage = rs!Atividade
130             UseCmbAtividadeISS.DataStorage = IIf(IsNull(rs!ATIVIDADE_ISS), 0, rs!ATIVIDADE_ISS)
                
140             chkAlmoxarifado.value = IIf(IsNull(rs!ALMOXARIFADO), 0, rs!ALMOXARIFADO)
150             ChkSND.value = IIf(IsNull(rs!SND), 0, rs!SND)
160             chkFarmacia.value = IIf(IsNull(rs!Farmacia), 0, rs!Farmacia)
170             txtCGCCPF.text = IIf(IsNull(rs!CGC), "", rs!CGC)
180             txtObservacao.text = IIf(IsNull(rs!Observacao), "", rs!Observacao)
190             txtBanco.text = IIf(IsNull(rs!Banco), "", rs!Banco)
200             txtAgencia.text = IIf(IsNull(rs!Agencia), "", rs!Agencia)
210             TxtContaCorrente.text = IIf(IsNull(rs!CONTABANCO), "", rs!CONTABANCO)
220             chkTemporario.value = IIf(IsNull(rs!TEMPORARIO), 0, rs!TEMPORARIO)
230             TelaPedidoCotacao = chkTemporario.value
240             chkTemporario.Enabled = chkTemporario.value
250             txtISS.text = IIf(IsNull(rs!PERC_ISS), 0, rs!PERC_ISS)
260             mskFORtelefone.text = acertaTelefone(rs!Telefone & "")
270             mskFax.text = acertaTelefone(rs!FAX & "")
                txtCodigoTotvs.text = IIf(IsNull(rs!CodigoTotvs), "", rs!CodigoTotvs)
280             mskFORtelefone_Validate False
                mskFax_Validate False
                
                'limpa campos da tab farmacia
290             TxtCodigo1.text = ""
300             TxtDescricao1.text = ""
310             txtFABendereco.text = ""
320             txtFABbairro.text = ""
330             txtFABInscricao.text = ""
340             TxtCidade1.text = ""
350             TxtNomeCidade1.text = ""
360             txtContato.text = ""
370             mskFABtelefone.text = ""
380             txtFax.text = ""
390             LimpaCampoMascara Me.mskFABcep
400             LimpaCampoMascara Me.mskFABcgc
410             txtEmail1.text = ""
                                    
420             chkDist.value = 0
430             ProcuraValor "select * from cidades where cidade=" & txtCidade.text, "Descricao", _
                Me.txtCidade, Me.TxtNomeCidade, ""
                
440             optPessoa(0).value = IIf(Not IsNull(rs!TipoPessoa) And rs!TipoPessoa = 1, 1, 0)
450             optPessoa(1).value = IIf(Not IsNull(rs!TipoPessoa) And rs!TipoPessoa = 2, 1, 0)
460             txtFundacao.text = IIf(Not IsNull(rs!DATAFUNDACAO), Format(rs!DATAFUNDACAO, "DD/MM/YYYY"), "__/__/____")
470             txtPrazoEntrega.text = IIf(Not IsNull(rs!PRAZOENTREGA), rs!PRAZOENTREGA, 0)
480             txtValorMinimo.text = saveFormataDecimal(IIf(Not IsNull(rs!VALORMINIMOPEDIDO), rs!VALORMINIMOPEDIDO, 0))
490             txtCondPgto.text = IIf(Not IsNull(rs!CondicaoPagamento), rs!CondicaoPagamento, 0)
500             txtRespNegoc.text = IIf(Not IsNull(rs!Respnegociacao), rs!Respnegociacao, "")
510             MontarGridRelacao
520             MontarGridProdutos
530            ElseIf Layout <> 11 Then
540               Call MsgBox("Fornecedor não localizado.", vbExclamation, App.Title)
550               cmdCancelar_Click
560               Exit Sub
570            End If
580         rs.Close
590         FmeCampos.Enabled = True
600         txtDescricao.Enabled = True 'Descrição ou o Próximo Campo
610         If FmeCampos.Visible = True Then txtDescricao.SetFocus 'Descrição ou o Próximo Campo
620       ElseIf KeyAscii = 13 And Trim(txtCodigo) = "" Then
630         sql = "SELECT MAX(fornecedor) AS Novo FROM fornecedores"
640         rs.Open sql, SV
650         If IsNull(rs!Novo) Then
660           txtCodigo = 1
670         Else
680           txtCodigo = rs!Novo + 1
690           If Layout = 9 Then TxtFORContaContabil.text = txtCodigo.text
700         End If
710         rs.Close
720         txtCodigo.Enabled = False
730         Inclui = True
740         FmeCampos.Enabled = True
750       If Not TelaPedidoCotacao And txtDescricao.Enabled Then txtDescricao.SetFocus   'Descrição ou o Próximo Campo
760       txtDescricao.Enabled = True 'Descrição ou o Próximo Campo
770       If fmeDados.Visible = True Then txtDescricao.SetFocus
780      End If
        
790     reSetAmpulheta
        
800   Exit Sub
Erro:
810      reSetAmpulheta
820      TratarErro
End Sub

Private Sub txtCodigo_LostFocus()
10        Me.KeyPreview = True
End Sub

Private Sub TxtCodigo1_GotFocus()
10       Me.KeyPreview = False
20       MarcaTexto Me.TxtCodigo1
End Sub

Private Sub TxtCodigo1_KeyPress(KeyAscii As Integer)
10    On Error GoTo Erro
      Dim rs As New Recordset, sql As String
       
20     KeyAscii = VerificaCampo(Me.TxtCodigo1, KeyAscii, tipInteiro, 0, 999999999)
30     If KeyAscii = 13 And Trim(TxtCodigo1) <> "" And IsNumeric(TxtCodigo1.text) Then
40       sql = "SELECT *,convert(int,Trava) as Travar FROM fabricantes WHERE fabricante = " & TxtCodigo1
50       rs.Open sql, SV
60       TxtCodigo1.Enabled = False
70       If rs.EOF Then
80         Inclui1 = True
90       Else
100        Inclui1 = False
110        Trava1 = rs!Travar
120        cmdExcluir1.Enabled = True
130        MostraCamposPublico Me, rs
140        UseCmbAtividade.DataStorage = rs!Atividade
           
150        txtFax.text = acertaTelefone(rs!FAX & "")
160        txtFax_Validate False
           
           'limpa campos da tab outros
170        txtCodigo.text = ""
180        txtDescricao.text = ""
190        txtFORendereco.text = ""
200        txtFORbairro.text = ""
210        txtFORInscricao.text = ""
220        txtCidade.text = ""
230        TxtNomeCidade.text = ""
240        TxtFORContaContabil.text = ""
250        mskFORtelefone.text = ""
260        LimpaCampoMascara Me.mskFax
270        LimpaCampoMascara Me.mskFORcep
           'LimpaCampoMascara Me.mskFORcgc
280        txtCGCCPF.text = ""
290        txtEmail.text = ""
           txtCodigoTotvs.text = ""
           
300        ProcuraValor "select * from cidades where cidade=" & TxtCidade1.text, "Descricao", _
           Me.TxtCidade1, Me.TxtNomeCidade1, ""
310      End If
320      rs.Close
330      FmeCampos1.Enabled = True
340      TxtDescricao1.Enabled = True 'Descrição ou o Próximo Campo
350      TxtDescricao1.SetFocus 'Descrição ou o Próximo Campo
360    ElseIf KeyAscii = 13 And Trim(TxtCodigo1) = "" Then
370      sql = "SELECT MAX(Fabricante) AS Novo FROM Fabricantes"
380      rs.Open sql, SV
390      If IsNull(rs!Novo) Then
400        TxtCodigo1 = 1
410      Else
420        TxtCodigo1 = rs!Novo + 1
430      End If
440      rs.Close
450      TxtCodigo1.Enabled = False
460      Inclui1 = True
470      FmeCampos1.Enabled = True
480      TxtDescricao1.Enabled = True 'Descrição ou o Próximo Campo
490      TxtDescricao1.SetFocus 'Descrição ou o Próximo Campo
500    End If


510   Exit Sub
Erro:
520      TratarErro
End Sub

Private Sub TxtCodigo1_LostFocus()
10        Me.KeyPreview = True
End Sub

Private Sub txtConteudo_GotFocus()
10        Me.KeyPreview = False
End Sub

Private Sub txtConteudo_KeyPress(KeyAscii As Integer)
10    On Error GoTo Erro
       Dim rs As New Recordset, sql As String
       
20     If cmbCampos.ListIndex <> -1 And Trim(txtConteudo) > "" And KeyAscii = 13 Then
30         Screen.MousePointer = 11
40         If cmbCampos.text = "Código" Then
50           If Not IsNumeric(txtConteudo.text) Then
60             MsgBox "O campo a ser pesquisado é numérico, deve haver algum caracter diferente na pesquisa!", vbInformation, Caption
70             Screen.MousePointer = 0
80             txtConteudo.SetFocus
90             Exit Sub
100          End If
110          sql = "SELECT * FROM fornecedores Where fornecedor = " & txtConteudo.text
120        ElseIf cmbCampos.text = "Descrição" Then
130          sql = "SELECT * FROM fornecedores Where Descricao LIKE '" & txtConteudo.text & "%'"
140        ElseIf cmbCampos.text = "Atividade" Then
150          sql = ""
160          sql = sql & " SELECT * FROM FORNECEDORES A "
170          sql = sql & "     LEFT JOIN ATIVIDADE B "
180          sql = sql & "        ON A.ATIVIDADE=B.ATIVIDADE "
190          sql = sql & " WHERE B.NOME LIKE '" & txtConteudo & "%'"
200        ElseIf cmbCampos.text = "Nome Fantasia" Then
210          sql = "SELECT * FROM fornecedores Where NomeFantasia LIKE '" & txtConteudo.text & "%'"
           '--------------------------------------------------------
           'Solicitacao de Araras para permitir consulta por CNPJ
           'Protocolo: 11283
           '--------------------------------------------------------
220        ElseIf cmbCampos.text = "CPF/CNPJ" Then
              
230           sql = ""
240           sql = " SELECT * FROM fornecedores"
250           sql = sql & " Where REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(CGC,':',''),',',''),'-',''),'/',''),'.','') "
260           sql = sql & " LIKE '" & RTrim$(Retira$(txtConteudo.text, "./-,:", UM_A_UM)) & "%'"
              
           '--------------------------------------------------------
270        End If
           
280        rs.Open sql, SV
290        Grid.Visible = False
300        Grid.Rows = 1
310        Do Until rs.EOF
320          Grid.AddItem "" & vbTab & rs!FORNECEDOR & vbTab & rs!Descricao & vbTab & rs!NOMEFANTASIA
330          rs.MoveNext
340        Loop
350        Grid.Visible = True
360        rs.Close
370        Screen.MousePointer = 0
380        LinhaVelha = 0
390    End If
       
400   Exit Sub
Erro:
410      TratarErro
End Sub

Private Sub TxtConteudo_LostFocus()
10       Me.KeyPreview = True
End Sub

Private Sub txtConteudo1_GotFocus()
10        Me.KeyPreview = False
End Sub

Private Sub TxtConteudo1_KeyPress(KeyAscii As Integer)
10    On Error GoTo Erro
      Dim rs As New Recordset, sql As String
20      If cmbCampos1.ListIndex <> -1 And Trim(txtConteudo1) > "" And KeyAscii = 13 Then
30         Screen.MousePointer = 11
40         If cmbCampos1.text = "Código" Then
50           If Not IsNumeric(txtConteudo1.text) Then
60             MsgBox "O campo a ser pesquisado é numérico, deve haver algum caracter diferente na pesquisa!", vbInformation, Caption
70             Screen.MousePointer = 0
80             txtConteudo1.SetFocus
90             Exit Sub
100          End If
110          sql = "SELECT * FROM fabricantes Where fabricante = " & txtConteudo1.text
120        ElseIf cmbCampos1.text = "Descrição" Then
130          sql = "SELECT * FROM fabricantes Where Descricao LIKE '" & txtConteudo1 & "%'"
140        ElseIf cmbCampos1.text = "Atividade" Then
150          sql = ""
160          sql = sql & " SELECT * FROM FABRICANTES A "
170          sql = sql & "     LEFT JOIN ATIVIDADE B "
180          sql = sql & "        ON A.ATIVIDADE=B.ATIVIDADE "
190          sql = sql & " WHERE B.NOME LIKE '" & txtConteudo1 & "%'"
200        End If
210        rs.Open sql, SV
220        Grid1.Visible = False
230        Grid1.Rows = 1
240        Do Until rs.EOF
250          Grid1.AddItem "" & vbTab & rs!Fabricante & vbTab & rs!Descricao & vbTab & rs!CGC & vbTab & rs!Inscricao
260          rs.MoveNext
270        Loop
280        Grid1.Visible = True
290        rs.Close
300        Screen.MousePointer = 0
310        LinhaVelha = 0
320    End If
       
330   Exit Sub
Erro:
340      TratarErro
End Sub

Private Sub TxtConteudo1_LostFocus()
10        Me.KeyPreview = True
End Sub

Private Sub TxtDescricao_GotFocus()
On Error GoTo Erro

10       MarcaTexto Me.txtDescricao

Exit Sub
Erro:
    TratarErro "frmFornecedores", "TxtDescricao_GotFocus", Err.Number, Err.Description, Erl
End Sub

Private Sub txtFax_Validate(Cancel As Boolean)
On Error GoTo Erro

10       validateTelefone txtFax, Cancel

Exit Sub
Erro:
    TratarErro "frmFornecedores", "txtFax_Validate", Err.Number, Err.Description, Erl
End Sub

Private Sub TxtNomeCidade_GotFocus()
On Error GoTo Erro

10       Me.KeyPreview = False

Exit Sub
Erro:
    TratarErro "frmFornecedores", "TxtNomeCidade_GotFocus", Err.Number, Err.Description, Erl
End Sub

Private Sub txtNomeCidade_KeyPress(KeyAscii As Integer)
On Error GoTo Erro
      Dim sql As String, rs As New ADODB.Recordset

20     If KeyAscii = 13 And Trim(TxtNomeCidade) <> "" Then
30       sql = "SELECT * FROM cidades WHERE descricao = '" & Trim(TxtNomeCidade) & "'"
40       rs.Open sql, SV
50       If rs.EOF Then
60         sql = "SELECT * FROM cidades WHERE descricao Like '" & Trim(TxtNomeCidade.text) & "%' ORDER BY descricao"
70         PreencheCombo SV, Lista, sql, "descricao", "cidade"
80         If Lista.ListCount = 0 Then
90           MsgBox "Não existem Cidades com estas iniciais!", vbInformation, Caption
100          rs.Close
110          Exit Sub
120        ElseIf Lista.ListCount = 1 Then
130          sql = "SELECT * FROM cidades WHERE Cidade = '" & Lista.ItemData(0)
140          rs.Open sql, SV
150          txtCidade.text = rs!Cidade
160          TxtNomeCidade.text = rs!Descricao
170          'mskFORcep.text = IIf(RS!CEP <> "", RS!CEP, "_____-___")
180          SendKeys "{TAB}"
190        Else
200          Lista.Left = 4680
210          Lista.Top = 600
220          Vtex = "txtnomecidade"
230          Vcod = "txtcidade"
240          Lista.Visible = True
250          Lista.SetFocus
260        End If
270      Else
280        txtCidade.text = rs!Cidade
290        TxtNomeCidade.text = rs!Descricao
300        'mskFORcep.text = IIf(RS!CEP <> "", RS!CEP, "_____-___")
310        SendKeys "{TAB}"
320      End If
330      rs.Close
340    End If

Exit Sub
Erro:
    TratarErro "frmFornecedores", "txtNomeCidade_KeyPress", Err.Number, Err.Description, Erl
       
End Sub

Private Sub TxtNomeCidade_LostFocus()
On Error GoTo Erro

10    Me.KeyPreview = True

Exit Sub
Erro:
    TratarErro "frmFornecedores", "TxtNomeCidade_LostFocus", Err.Number, Err.Description, Erl
End Sub

Private Sub TxtNomeCidade1_GotFocus()
On Error GoTo Erro

10        Me.KeyPreview = False

Exit Sub
Erro:
    TratarErro "frmFornecedores", "TxtNomeCidade1_GotFocus", Err.Number, Err.Description, Erl
End Sub

Private Sub TxtNomeCidade1_KeyPress(KeyAscii As Integer)
On Error GoTo Erro
      Dim sql As String, rs As New ADODB.Recordset
      
20     If KeyAscii = 13 And Trim(TxtNomeCidade1) <> "" Then
30       sql = "SELECT * FROM cidades WHERE descricao = '" & Trim(TxtNomeCidade1) & "'"
40       rs.Open sql, SV
50       If rs.EOF Then
60         sql = "SELECT * FROM cidades WHERE descricao Like '" & Trim(TxtNomeCidade1.text) & "%' ORDER BY descricao"
70         PreencheCombo SV, Lista1, sql, "descricao", "cidade"
80         If Lista1.ListCount = 0 Then
90           MsgBox "Não existem Cidades com estas iniciais!", vbInformation, Caption
100          rs.Close
110          Exit Sub
120        ElseIf Lista1.ListCount = 1 Then
130          TxtCidade1.text = Lista1.ItemData(0)
140          TxtNomeCidade1.text = Lista1.List(0)
150          SendKeys "{TAB}"
160        Else
170          Lista1.Left = 4680
180          Lista1.Top = 510
190          Vtex = "TxtNomeCidade1"
200          Vcod = "TxtCidade1"
210          Lista1.Visible = True
220          Lista1.SetFocus
230        End If
240      Else
250        TxtCidade1.text = rs!Cidade
260        TxtNomeCidade1.text = rs!Descricao
270        SendKeys "{TAB}"
280      End If
290      rs.Close
300    End If

Exit Sub
Erro:
    TratarErro "frmFornecedores", "TxtNomeCidade1_KeyPress", Err.Number, Err.Description, Erl
      
End Sub

Private Sub TxtNomeCidade1_LostFocus()
10    On Error GoTo Erro

20    Me.KeyPreview = True

30    Exit Sub
Erro:
40        TratarErro "frmFornecedores", "TxtNomeCidade1_LostFocus", Err.Number, Err.Description, Erl
End Sub

Private Sub txtPrazoEntrega_Validate(Cancel As Boolean)
On Error GoTo Erro

10        txtPrazoEntrega.text = Val(txtPrazoEntrega.text)

Exit Sub
Erro:
    TratarErro "frmFornecedores", "txtPrazoEntrega_Validate", Err.Number, Err.Description, Erl
End Sub

Private Sub txtValorMinimo_Click()
On Error GoTo Erro

10            txtValorMinimo.text = saveFormataDecimal(txtValorMinimo.text, 2)
20            txtValorMinimo.SelStart = Len(txtValorMinimo)

Exit Sub
Erro:
    TratarErro "frmFornecedores", "txtValorMinimo_Click", Err.Number, Err.Description, Erl
End Sub

Private Sub txtValorMinimo_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Erro

10            txtValorMinimo.text = saveFormataDecimal(txtValorMinimo.text, 2)
20            txtValorMinimo.SelStart = Len(txtValorMinimo)

Exit Sub
Erro:
    TratarErro "frmFornecedores", "txtValorMinimo_KeyUp", Err.Number, Err.Description, Erl
End Sub


Public Function VCGC(ST As String) As Integer
On Error GoTo Erro
  Dim retVal As Integer, X As String, Posi As Integer, _
        dv1c As Integer, dv2c As Integer, dv1f As Integer, dv2f As Integer, _
        Num As String, Mu As String, Resto As Integer, DV As String, QtdeC As Integer

20     X$ = RTrim$(Retira$(ST$, "./-,:", UM_A_UM))
30     If Len(X$) = 0 Then                             'se vazio...
40       retVal = True                                 'preparamos retorno true
50     Else                                            'senão,
60       retVal = False                                'preparamos false
70     End If
80     If Len(X$) > 10 Then                            'se tem 14 caracteres = ok
90       QtdeC = Len(X$)
100      dv1f = Val(Mid$(X$, QtdeC - 1, 1))            'salva os dígitos
110      dv2f = Val(Right$(X$, 1))                     'fornecidos
120      Num$ = Left$(X$, QtdeC - 2)                   'separa o número
130      dv1c = 0                                      'inicializa dv1 a calcular
140      If Len(X$) = 14 Then
150        Mu$ = "543298765432"                        'constante multiplicadora
160      Else
170        Mu$ = "298765432"                           'constante multiplicadora
180      End If
190      Posi = QtdeC - 2                              'inicializa posição
200      While Posi > 0
210        dv1c = dv1c + Val(Mid$(Num$, Posi, 1)) * Val(Mid$(Mu$, Posi, 1))
220        Posi = Posi - 1                             'acumulando cada dígido X o seu multiplicador
230      Wend                                          'decrementa contador de posição
240      Resto = dv1c Mod 11                           'caula o resto (módulo 11)
250      If Resto < 2 Then                             'se menor do que 2
260        dv1c = 0                                    'o dv é o resto
270      Else                                          'senão,
280        dv1c = 11 - Resto                           'este dv é a diferença 11 - resto
290      End If
300      DV$ = Right$(STR$(dv1c), 1)                   'salva o dv calculado como string
310      Num$ = Num$ + DV$                             'incorpora dv1
320      dv2c = 0                                      'inicializa dv2
330      If Len(X$) = 14 Then
340        Mu$ = "6" + Mu$                             'poe mais um dígito nos multiplicadores
350      Else
360        Mu$ = "3" + Mu$                             'poe mais um dígito nos multiplicadores
370      End If
380      Posi = QtdeC - 1                              'posição agora inicia em 13
390      While Posi > 0                                'vamos fazer a mesma coisa,
400        dv2c = dv2c + Val(Mid$(Num$, Posi, 1)) * Val(Mid$(Mu$, Posi, 1))
410        Posi = Posi - 1                             'que fizemos acima
420      Wend
430      Resto = dv2c Mod 11                           'pega o resto da divisão por 11
440      If Resto < 2 Then                             'se menor do que 2
450        dv2c = 0                                    'o dv é 0
460      Else                                          'senão,
470        dv2c = 11 - Resto                           'o dv é a diferença
480      End If
490      retVal = (dv1c = dv1f And dv2c = dv2f)        'prepara retorno
500    End If
510    VCGC = retVal                                   'true se DVs fornecidos iguais aos calculados

Exit Function
Erro:
    TratarErro "frmFornecedores", "VCGC", Err.Number, Err.Description, Erl

End Function

'Private Sub cmdImportar_Click()
'Dim xl As New Excel.Application
'   Dim xlw As Excel.Workbook
'   Dim i As Integer
'   Dim tbl As rdoResultset
'
'   Set xlw = xl.Workbooks.Open("C:\Users\User\Documents\Clientes\UnimedBP\Tabelas - SAVE ERP\Tabelas - SAVE ERP\tabela-11-fornecedor.xls")
'
'   xlw.Sheets("fornecedor").Select
'
'   With xlw.Application
'      For i = 2 To 2685
'         CmdCancelar_Click
'
'         TxtCodigo.Text = .Range("A" & i).Text
'         TxtCodigo_KeyPress 13
'
'         txtDescricao.Text = Replace(.Range("B" & i).Text, "'", "")
'         txtFORendereco.Text = Replace(.Range("J" & i).Text & " " & .Range("K" & i).Text, "'", "")
'         txtFORbairro.Text = Replace(.Range("P" & i).Text, "'", "")
'         TxtNomeCidade.Text = Replace(.Range("N" & i).Text, "'", "")
'         txtNomeCidade_KeyPress 13
'         If Trim(.Range("M" & i).Text) <> "" Then
'            mskFORcep.Text = .Range("M" & i).Text
'         End If
'         txtNomeFantasia.Text = Replace(.Range("C" & i).Text, "'", "")
'         TxtFORContaContabil.Text = "123"
'         txtFORInscricao.Text = .Range("F" & i).Text
'         txtCGCCPF.Text = .Range("E" & i).Text
'         txtCGCCPF_LostFocus
'
'         CmdSalvar_Click
'      Next
'   End With
'End Sub


Private Function validarContaContabil() As Boolean
On Error GoTo Erro
10       validarContaContabil = True
20       If Trim(TxtFORContaContabil.text) = "" Then
30          Exit Function
40       End If
         
50       If RecuperaCampo(Trim(TxtFORContaContabil.text), "conta", "conta", "cont_planoconta") = "" Then
60          validarContaContabil = False
70       End If

Exit Function
Erro:
    TratarErro "frmFornecedores", "validarContaContabil", Err.Number, Err.Description, Erl
End Function



Public Function VDV2(ST As String) As Integer
On Error GoTo Erro
  Dim X As String, i As Integer, Num As String, dvf As String, _
      dvc As String, retVal As Integer             'dimensiona

20     X$ = RTrim$(Retira$(ST$, "./-,:", UM_A_UM))      'tira separadores
30     If Len(X$) = 0 Then                              'se nada veio,
40       retVal = True                                  'retorna true
50     Else                                             'senão,
60       Num$ = Left$(X$, Len(X$) - 2)                  'salva o número
70       dvf$ = Right$(X$, 2)                           'e o dv fornecido
80       dvc$ = GDV1$(Num$)                             'primeiro digito
90       Num$ = Num$ + dvc$                             'incorpora 1o digito
100      dvc$ = dvc$ + GDV1$(Num$)                      'calcula o 2o. dv
110      retVal = (dvc$ = dvf$)                         'se igual - true, senão falso
120    End If
130    VDV2 = retVal                                    'retorna

Exit Function
Erro:
    TratarErro "frmFornecedores", "VDV2", Err.Number, Err.Description, Erl
End Function

