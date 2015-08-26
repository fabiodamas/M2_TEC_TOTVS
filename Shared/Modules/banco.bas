Attribute VB_Name = "AtualizaBanco"
Public Function AtualizaMes(MesAux As Long)
   On Error GoTo Erro
   
   Dim mes As Long
   Dim ano As Long
   Dim MesStr As String
   Dim DataAux As String
   
   MesStr = MesAux
   
   If Len(MesStr) = 5 Then
      mes = Mid(MesStr, 1, 1)
      ano = Mid(MesStr, 2, 4)
   Else
      mes = Mid(MesStr, 1, 2)
      ano = Mid(MesStr, 3, 4)
   End If
   
   DataAux = "01/" & Format(mes, "00") & "/" & ano
   
   '08/2010
   If Format(DataAux, "yyyy/mm/dd") <= "2010/08/31" Then
      AtualizaMes82010
   End If
   
   '09/2010
   If Format(DataAux, "yyyy/mm/dd") <= "2010/09/30" Then
      AtualizaMes92010
   End If
   
   '10/2010
   If Format(DataAux, "yyyy/mm/dd") < "2010/10/31" Then
      AtualizaMes102010
   End If
      
   '11/2010
   If Format(DataAux, "yyyy/mm/dd") < "2010/11/30" Then
      AtualizaMes112010
   End If
   
   '12/2010
   If Format(DataAux, "yyyy/mm/dd") < "2010/12/31" Then
      AtualizaMes122010
   End If
   
   '01/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/01/31" Then
      AtualizaMes012011
   End If
   
   '02/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/02/28" Then
      AtualizaMes022011
   End If
   
   '03/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/03/31" Then
      AtualizaMes032011
   End If
   
   '04/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/04/30" Then
      AtualizaMes042011
   End If
   
   '05/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/05/31" Then
      AtualizaMes052011
   End If
   
   '06/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/06/30" Then
      AtualizaMes062011
   End If
   
   '07/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/07/31" Then
      AtualizaMes072011
   End If
   
   '08/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/08/31" Then
      AtualizaMes082011
   End If
   
   '09/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/09/30" Then
      AtualizaMes092011
   End If
   
   '10/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/10/31" Then
      AtualizaMes102011
   End If
   
   '11/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/11/30" Then
      AtualizaMes112011
   End If
   
   '12/2011
   If Format(DataAux, "yyyy/mm/dd") < "2011/12/31" Then
      AtualizaMes122011
   End If
   
   '01/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/01/31" Then
      AtualizaMes012012
   End If
   
   '02/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/02/29" Then
      AtualizaMes022012
   End If
   
   '03/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/03/31" Then
      AtualizaMes032012
   End If
   
   '04/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/04/30" Then
      AtualizaMes042012
   End If
   
   '05/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/05/31" Then
      AtualizaMes052012
   End If
   
   '06/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/06/30" Then
      AtualizaMes062012
   End If
   
   '07/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/07/31" Then
      AtualizaMes072012
   End If
   
   '08/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/08/31" Then
      AtualizaMes082012
   End If
   
   '09/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/09/30" Then
      AtualizaMes092012
   End If
   
   '10/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/10/31" Then
      AtualizaMes102012
   End If
   
   '11/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/11/30" Then
      AtualizaMes112012
   End If
   
   '12/2012
   If Format(DataAux, "yyyy/mm/dd") < "2012/12/31" Then
      AtualizaMes122012
   End If
   
   '01/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/01/31" Then
      AtualizaMes012013
   End If
   
   '02/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/02/31" Then
      AtualizaMes022013
   End If
   
   '03/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/03/31" Then
      AtualizaMes032013
   End If
   
   '04/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/04/30" Then
      AtualizaMes042013
   End If
   
   '05/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/05/31" Then
      AtualizaMes052013
   End If
   
   '06/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/06/30" Then
      AtualizaMes062013
   End If
   
   '07/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/07/31" Then
      AtualizaMes072013
   End If
   
   '08/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/08/31" Then
      AtualizaMes082013
   End If
   
   '09/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/09/30" Then
      AtualizaMes092013
   End If
   
   '10/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/10/31" Then
      AtualizaMes102013
   End If
   
   '11/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/11/30" Then
      AtualizaMes112013
   End If
      
   '12/2013
   If Format(DataAux, "yyyy/mm/dd") < "2013/12/31" Then
      AtualizaMes122013
   End If
   
   '01/2014
   If Format(DataAux, "yyyy/mm/dd") < "2014/01/31" Then
      AtualizaMes012014
   End If
   
   '02/2014
   If Format(DataAux, "yyyy/mm/dd") < "2014/02/28" Then
      AtualizaMes022014
   End If
   
   '03/2014
   If Format(DataAux, "yyyy/mm/dd") < "2014/03/31" Then
      AtualizaMes032014
   End If
   
   '04/2014
   If Format(DataAux, "yyyy/mm/dd") < "2014/04/30" Then
      AtualizaMes042014
   End If
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes82010()
   On Error GoTo Erro
   
   
   sql = " ALTER TABLE CONVENIOS ADD TISS_INTERNACAO_EXPORTAPROCEDIMENTOOBSERVACAO INT"
   Banco.Execute sql
   
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes92010()
   On Error GoTo Erro
   
   
   sql = " ALTER TABLE PARAMETRO ADD ULTIMAATUALIZACAO INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE FICHAS ADD CARTAOCIDADAO INT "
   Banco.Execute sql
      
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[COR]("
   sql = sql & "    [COR] [int] NOT NULL,"
   sql = sql & "    [DESCRICAO] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [CODIGOCOR] [varchar](30) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_COR] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [Cor] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = " INSERT INTO COR (COR, DESCRICAO) VALUES (1, 'Vermelho')"
   Banco.Execute sql
   
   sql = " INSERT INTO COR (COR, DESCRICAO) VALUES (2, 'Verde')"
   Banco.Execute sql
   
   sql = " INSERT INTO COR (COR, DESCRICAO) VALUES (3, 'Amarelo')"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD CORATENDIMENTO INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE MEDICOHORARIO ADD TIPOAGENDAMENTO INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE AGENDAMENTOCONSULTA ADD TIPOAGENDAMENTO INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE USUARIO ADD NAOPERMITEALTERARAGENDAMENTO INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE AGENDAMENTOCONSULTA ADD FICHA_AGENDA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD PRES_AGRUPAPERIODO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE LAU_MOVIM_AMB ADD PROCEDIMENTOTUSS INT"
   Banco.Execute sql

   sql = " ALTER TABLE LAU_MOVIM_AMB ADD PROCEDIMENTONOMETUSS VARCHAR(250)"
   Banco.Execute sql
   
   sql = " ALTER TABLE LAU_MOVIM_INT ADD PROCEDIMENTOTUSS INT"
   Banco.Execute sql

   sql = " ALTER TABLE LAU_MOVIM_INT ADD PROCEDIMENTONOMETUSS VARCHAR(250)"
   Banco.Execute sql
   
   sql = " ALTER TABLE LAU_MOVIM_EXT ADD PROCEDIMENTOTUSS INT"
   Banco.Execute sql

   sql = " ALTER TABLE LAU_MOVIM_EXT ADD PROCEDIMENTONOMETUSS VARCHAR(250)"
   Banco.Execute sql
      
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[AVISO_PRESCRICAO]("
   sql = sql & "    [SEQUENCIA]     [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & "    [PRESCRICAO]    [int] NULL,"
   sql = sql & "    [REGISTRO]      [int] NULL,"
   sql = sql & "    [PRODUTO]       [int] NULL,"
   sql = sql & "    [DESCRICAO]     [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [PACIENTE]      [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [AVISORECEBIDO] [int] NULL,"
   sql = sql & "    [TEXTO]         [varchar](250) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [ATUALIZACAO]   [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_AVISO_PRESCRICAO] PRIMARY KEY CLUSTERED"
   sql = sql & "  ("
   sql = sql & "      [Sequencia] Asc"
   sql = sql & "  ) WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
      
   sql = " ALTER TABLE PARAMETRO ADD PRES_CONFERENCIA_HORA  INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT ADD HORACONFERIDA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_AMB ADD HORACONFERIDA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOTIVOCOBRANCA ADD DESATIVADO INT "
   Banco.Execute sql

   sql = " ALTER TABLE ALTAGERAL ADD DESATIVADO INT"
   Banco.Execute sql

   'ALTERAÇÃO PORTARIA
   sql = " UPDATE MOTIVOCOBRANCA SET DESATIVADO = 1 WHERE MOTIVOCOBRANCA IN (13, 17, 29)"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESATIVADO = 1 WHERE ALTAGERAL IN (13, 17, 29)"
   Banco.Execute sql

  '===============
   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'ALTA DE PACIENTE AGUDO EM PSIQUIATRIA' WHERE MOTIVOCOBRANCA = 19"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESCRICAO = 'ALTA DE PACIENTE AGUDO EM PSIQUIATRIA' WHERE ALTAGERAL = 19"
   Banco.Execute sql
  '===============

   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'TRANSFERENCIA PARA INTERNACAO DOMICILIAR' WHERE MOTIVOCOBRANCA = 32"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESCRICAO = 'TRANSFERENCIA PARA INTERNACAO DOMICILIAR' WHERE ALTAGERAL = 32"
   Banco.Execute sql
   '===============

   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'ALTA DA MAE/PUERPERA E DO RECEM-NASCIDO' WHERE MOTIVOCOBRANCA = 61"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESCRICAO = 'ALTA DA MAE/PUERPERA E DO RECEM-NASCIDO' WHERE ALTAGERAL = 61"
   Banco.Execute sql
   '===============

   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'ALTA DA MAE/PUERPERA E PERMANENCIA DO RECEM-NASCIDO' WHERE MOTIVOCOBRANCA = 62"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESCRICAO = 'ALTA DA MAE/PUERPERA E PERMANENCIA DO RECEM-NASCIDO' WHERE ALTAGERAL = 62"
   Banco.Execute sql
   '===============
   
   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'ALTA DA MAE/PUERPERA E OBITO DO RECEM-NASCIDO' WHERE MOTIVOCOBRANCA = 63"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESCRICAO = 'ALTA DA MAE/PUERPERA E OBITO DO RECEM-NASCIDO' WHERE ALTAGERAL = 63"
   Banco.Execute sql
   '===============
   
   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'ALTA DA MAE/PUERPERA COM OBITO FETAL' WHERE MOTIVOCOBRANCA = 64"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESCRICAO = 'ALTA DA MAE/PUERPERA COM OBITO FETAL' WHERE ALTAGERAL = 64"
   Banco.Execute sql
   '===============
   
   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'OBITO DA GESTANTE E DO CONCEPTO' WHERE MOTIVOCOBRANCA = 65"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESCRICAO = 'OBITO DA GESTANTE E DO CONCEPTO' WHERE ALTAGERAL = 65"
   Banco.Execute sql
   '===============
   
   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'OBITO DA MAE/PUERPERA E ALTA DO RECEM-NASCIDO' WHERE MOTIVOCOBRANCA = 66"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESCRICAO = 'OBITO DA MAE/PUERPERA E ALTA DO RECEM-NASCIDO' WHERE ALTAGERAL = 66"
   Banco.Execute sql
   '===============
   
   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'OBITO DA MAE/PUERPERA E PERMANENCIA DO RECEM-NASCIDO' WHERE MOTIVOCOBRANCA = 67"
   Banco.Execute sql

   sql = " UPDATE ALTAGERAL SET DESCRICAO = 'OBITO DA MAE/PUERPERA E PERMANENCIA DO RECEM-NASCIDO' WHERE ALTAGERAL = 67"
   Banco.Execute sql
   '===============
   
   sql = " UPDATE MOTIVOCOBRANCA SET DESCRICAO = 'OBITO DA MAE/PUERPERA E PERMANENCIA DO RECEM-NASCIDO' WHERE MOTIVOCOBRANCA = 67"
   Banco.Execute sql
   
   
   Exit Function
Erro:
   Resume Next
End Function


Public Function AtualizaMes102010()
   On Error GoTo Erro
   
   
   sql = " ALTER TABLE CONVENIOS ADD TISS_INTERNACAO_EXPORTAPROCEDIMENTOOBSERVACAO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD TUSS_PESQUISA_CODIGOTUSS INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE USUARIO ADD NAOPERMITEALTERARAGENDAMENTO INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[LAU_ASSOCIACAOLAUDO]("
   sql = sql & "    [MEDICO] [int] NOT NULL,"
   sql = sql & "    [CENTROCUSTO] [int] NOT NULL,"
   sql = sql & "    [LAUDO] [int] NOT NULL,"
   sql = sql & "    [LAUDOASSOCIADO] [int] NOT NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_LAU_ASSOCIACAOLAUDO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [MEDICO] ASC,"
   sql = sql & "    [CENTROCUSTO] ASC,"
   sql = sql & "    [LAUDO] ASC,"
   sql = sql & "    [LAUDOASSOCIADO] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = " ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT ADD HORACONFERIDA_AUX INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT ADD IDIMPRESSAO VARCHAR(255)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " ALTER TABLE PRES_TMP_MEDICACAO ADD"
   sql = sql & " HORAQTDE1 INT,"
   sql = sql & " HORAQTDE2 INT,"
   sql = sql & " HORAQTDE3 INT,"
   sql = sql & " HORAQTDE4 INT,"
   sql = sql & " HORAQTDE5 INT,"
   sql = sql & " HORAQTDE6 INT,"
   sql = sql & " HORAQTDE7 INT,"
   sql = sql & " HORAQTDE8 INT,"
   sql = sql & " HORAQTDE9 INT,"
   sql = sql & " HORAQTDE10 INT,"
   sql = sql & " HORAQTDE11 INT,"
   sql = sql & " HORAQTDE12 INT,"
   sql = sql & " HORAQTDE13 INT,"
   sql = sql & " HORAQTDE14 INT,"
   sql = sql & " HORAQTDE15 INT,"
   sql = sql & " HORAQTDE16 INT,"
   sql = sql & " HORAQTDE17 INT,"
   sql = sql & " HORAQTDE18 INT,"
   sql = sql & " HORAQTDE19 INT,"
   sql = sql & " HORAQTDE20 INT,"
   sql = sql & " HORAQTDE21 INT,"
   sql = sql & " HORAQTDE22 INT,"
   sql = sql & " HORAQTDE23 INT,"
   sql = sql & " HORAQTDE24 INT,"
   sql = sql & " HORAQTDE7_2 INT,"
   sql = sql & " HORAQTDE8_2 INT"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MOTIVOCOBRANCA (MOTIVOCOBRANCA, DESCRICAO, CONVENIO)"
   sql = sql & " SELECT TOP 1 61, 'ALTA DA MAE/PUERPERA E DO RECEM-NASCIDO', 0"
   sql = sql & " From MOTIVOCOBRANCA"
   sql = sql & " WHERE 61 NOT IN (SELECT MOTIVOCOBRANCA"
   sql = sql & "             FROM MOTIVOCOBRANCA)"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MOTIVOCOBRANCA (MOTIVOCOBRANCA, DESCRICAO, CONVENIO)"
   sql = sql & " SELECT TOP 1 62, 'ALTA DA MAE/PUERPERA E PERMANENCIA DO RECEM-NASCIDO', 0"
   sql = sql & " From MOTIVOCOBRANCA"
   sql = sql & " WHERE 62 NOT IN (SELECT MOTIVOCOBRANCA"
   sql = sql & "             FROM MOTIVOCOBRANCA)"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MOTIVOCOBRANCA (MOTIVOCOBRANCA, DESCRICAO, CONVENIO)"
   sql = sql & " SELECT TOP 1 63, 'ALTA DA MAE/PUERPERA E OBITO DO RECEM-NASCIDO', 0"
   sql = sql & " From MOTIVOCOBRANCA"
   sql = sql & " WHERE 63 NOT IN (SELECT MOTIVOCOBRANCA"
   sql = sql & "             FROM MOTIVOCOBRANCA)"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MOTIVOCOBRANCA (MOTIVOCOBRANCA, DESCRICAO, CONVENIO)"
   sql = sql & " SELECT TOP 1 64, 'ALTA DA MAE/PUERPERA COM OBITO FETAL', 0"
   sql = sql & " From MOTIVOCOBRANCA"
   sql = sql & " WHERE 64 NOT IN (SELECT MOTIVOCOBRANCA"
   sql = sql & "             FROM MOTIVOCOBRANCA)"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MOTIVOCOBRANCA (MOTIVOCOBRANCA, DESCRICAO, CONVENIO)"
   sql = sql & " SELECT TOP 1 65, 'OBITO DA GESTANTE E DO CONCEPTO', 0"
   sql = sql & " From MOTIVOCOBRANCA"
   sql = sql & " WHERE 65 NOT IN (SELECT MOTIVOCOBRANCA"
   sql = sql & "             FROM MOTIVOCOBRANCA)"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MOTIVOCOBRANCA (MOTIVOCOBRANCA, DESCRICAO, CONVENIO)"
   sql = sql & " SELECT TOP 1 66, 'OBITO DA MAE/PUERPERA E ALTA DO RECEM-NASCIDO', 0"
   sql = sql & " From MOTIVOCOBRANCA"
   sql = sql & " WHERE 66 NOT IN (SELECT MOTIVOCOBRANCA"
   sql = sql & "             FROM MOTIVOCOBRANCA)"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MOTIVOCOBRANCA (MOTIVOCOBRANCA, DESCRICAO, CONVENIO)"
   sql = sql & " SELECT TOP 1 67, 'OBITO DA MAE/PUERPERA E PERMANENCIA DO RECEM-NASCIDO', 0"
   sql = sql & " From MOTIVOCOBRANCA"
   sql = sql & " WHERE 67 NOT IN (SELECT MOTIVOCOBRANCA"
   sql = sql & "             FROM MOTIVOCOBRANCA)"
   Banco.Execute sql

   sql = " ALTER TABLE PRES_ENFERMAGEM_MOVIM_INT ADD DIAGNOSTICO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PRES_ENFERMAGEM_MOVIM_INT ADD DEFINICAO VARCHAR(255)"
   Banco.Execute sql
   
   sql = " ALTER TABLE OUTROS_TIPO ADD TIPOASSOCIADO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD PRONTOATENDIMENTO_CC INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE PARAMETRO ADD REPETE_PRESCRICAO_MEDICA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD USUARIOREPETIRPRESCRICAO VARCHAR(50) "
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD SENHAREPETIRPRESCRICAO VARCHAR(50) "
   Banco.Execute sql
   
   sql = " ALTER TABLE USUARIO ADD NAOPERMITEALTERACAOFINANCEIRO INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE PARAMETRO ADD DATAREPETICAO DATETIME"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[FORNECEDOR_RELACAO]("
   sql = sql & "    [FORNECEDOR] [int] NOT NULL,"
   sql = sql & "    [FORNECEDOR_RELACAO] [int] NOT NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_FORNECEDOR_RELACAO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [FORNECEDOR] ASC,"
   sql = sql & "    [FORNECEDOR_RELACAO] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = " ALTER TABLE PRODUTOTIPO ADD ITEMSOUTILIZADOPRESCRICAO INT"
   Banco.Execute sql

   sql = " ALTER TABLE MOVIM_PRES_INT ADD ITEMASSOCIADO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOVIM_PRES_AMB ADD ITEMASSOCIADO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE MEDICOHORARIO ADD SOREMARCACAO INT"
   Banco.Execute sql
   
   sql = "  ALTER TABLE MEDICOHORARIO ADD OBSERVACAO VARCHAR(255)"
   Banco.Execute sql

   sql = " ALTER TABLE MOVIM_PRES_EXT ADD ITEMASSOCIADO INT"
   Banco.Execute sql
   
   sql = " "
   sql = sql & " CREATE TABLE FERIADOS"
   sql = sql & " ("
   sql = sql & "     FERIADO INT,"
   sql = sql & "     DESCRICAO VARCHAR(80),"
   sql = sql & "     DIA INT,"
   sql = sql & "     MES INT,"
   sql = sql & "     ANO INT,"
   sql = sql & "     Atualizacao VarChar(155)"
   sql = sql & " )"
   Banco.Execute sql

   sql = " ALTER TABLE PRES_CAD_INTERVALO ADD NAOEXIBEHORARIO INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Tempo de Espera de Atendimento', 'MnuRec_Rel_TEs', ' ', 'MnuRec_Rel_TEs',"
   sql = sql & "        1, 1, '0213120000', 'MnuRec_Rel_TEs', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes112010()
   On Error GoTo Erro

   '=================================
   'MES DE NOVEMBRO DE 2010
   '=================================
   
   sql = ""
   sql = sql & " ALTER TABLE CONVENIOS ADD PROCEDIMENTO_REGISTROCONSULTA INT"
   Banco.Execute sql
   
   
   sql = ""
   sql = sql & " ALTER TABLE PARAMETRO ADD CENTROCUSTOFARMACIAPOPULAR INT"
   Banco.Execute sql
    
   sql = ""
   sql = sql & " ALTER TABLE PARAMETRO ADD CAMINHOFOTO VARCHAR(155)"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "    MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "    NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Farmácia Popular', 'mnuPrd_Fpo', ' ', 'mnuPrd_Fpo',"
   sql = sql & "    1, 2, '0113000000', 'mnuPrd_Fpo', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[CBHPMPORTE5]("
   sql = sql & " [TIPO] [int] NOT NULL,"
   sql = sql & " [CODIGO] [char](5) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & " [VALOR] [money] NOT NULL,"
   sql = sql & " [VALOR_LAUDO] [money] NULL,"
   sql = sql & " CONSTRAINT [PK_CBHPMPORTE5] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "     [TIPO] ASC,"
   sql = sql & "     [Codigo] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
      
   'INSERE OS PORTES NOVOS DA CBHPM5
   sql = ""
   sql = sql & " INSERT INTO CBHPMPORTE5 (TIPO, CODIGO, VALOR, VALOR_LAUDO ) " & Chr(13)
   sql = sql & " SELECT A.TIPO, A.CODIGO, A.VALOR, A.VALOR_LAUDO" & Chr(13)
   sql = sql & " FROM CBHPMPORTE A" & Chr(13)
   sql = sql & " WHERE A.CODIGO NOT IN (SELECT B.CODIGO" & Chr(13)
   sql = sql & "                          FROM CBHPMPORTE5 B" & Chr(13)
   sql = sql & "                          WHERE A.TIPO = B.TIPO " & Chr(13)
   sql = sql & "                          AND   A.CODIGO = B.CODIGO)" & Chr(13)
   Banco.Execute sql
         
   '=====================================================
   'SCRIPT JAGUARIUNA
   '=====================================================
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[INTERNO_AVALIACAO]("
   sql = sql & " [AVALIACAO] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & " [REGISTRO] [int] NOT NULL,"
   sql = sql & " [SERVICO] [int] NOT NULL,"
   sql = sql & " [NOTA] [int] NULL,"
   sql = sql & " [ATUALIZACAO] [varchar](80) NULL,"
   sql = sql & " CONSTRAINT [PK_INTERNO_AVALIACAO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & " [AVALIACAO] Asc"
   sql = sql & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[ESCALAPLANTAO]("
   sql = sql & " [ESCALAPLANTAO] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & " [MEDICO] [int] NOT NULL,"
   sql = sql & " [DATAINICIO] [datetime] NOT NULL,"
   sql = sql & " [DATAFINAL] [datetime] NOT NULL,"
   sql = sql & " [HORAINICIO] [varchar](5) NOT NULL,"
   sql = sql & " [HORAFINAL] [varchar](5) NOT NULL,"
   sql = sql & " [ATUALIZACAO] [varchar](80) NULL,"
   sql = sql & " CONSTRAINT [PK_ESCALAPLANTAO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & " [ESCALAPLANTAO] Asc"
   sql = sql & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = " ALTER TABLE FICHAS ADD LOCALARMAZENAMENTO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE FICHAS ADD USUARIOARMAZENAMENTO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE FICHAS ADD DATAULTIMATRANSFERENCIAPRONTUARIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE FICHAS ADD HOTAULTIMATRANSFERENCIAPRONTUARIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO ADD OBS_AVALIACAO VARCHAR(155)"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD PRONTUARIOUNICO INT"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & " MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Avaliação de Pacientes internados', 'mnuRec_Ain', ' ', 'mnuRec_Ain',"
   sql = sql & " 1, 1, '0217000000', 'mnuRec_Ain', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql

   sql = " ALTER TABLE INTERNO ADD OBSERVACAOTRANSFERENCIAPRONTUARIO VARCHAR(250)"
   Banco.Execute sql
      
   sql = " ALTER TABLE INTERNO ADD OBS_AVALIACAO VARCHAR(155)"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD PRONTUARIOUNICO INT"
   Banco.Execute sql

   sql = " alter table TMPSAME_MOVIMENTACAO_PRONTUARIO add OBSERVACAO VARCHAR(255)"
   Banco.Execute sql

   sql = " alter table SAME_MOVIMENTACAO_PRONTUARIO_INT add OBSERVACAO VARCHAR(255)"
   Banco.Execute sql

   sql = " alter table SAME_MOVIMENTACAO_PRONTUARIO_EXT add OBSERVACAO VARCHAR(255)"
   Banco.Execute sql

   sql = " alter table SAME_MOVIMENTACAO_PRONTUARIO_AMB add OBSERVACAO VARCHAR(255)"
   Banco.Execute sql

   sql = " ALTER TABLE TMPRELATORIOUSUARIOTRANSFERENCIA ADD OBSERVACAO VARCHAR(250) "
   Banco.Execute sql
   
   sql = " ALTER TABLE SAME_MOVIMENTACAO_PRONTUARIO_FICHA ADD OBSERVACAO VARCHAR(155)"
   Banco.Execute sql

   sql = " ALTER TABLE TMPSAME_MOVIMENTACAO_PRONTUARIO ADD OBSERVACAO VARCHAR(155)"
   Banco.Execute sql

   sql = " alter table INTERNO ADD OBSERVACAOTRANSFERENCIAPRONTUARIO VARCHAR(255)"
   Banco.Execute sql
   
   sql = " alter table EXTERNO ADD OBSERVACAOTRANSFERENCIAPRONTUARIO VARCHAR(255)"
   Banco.Execute sql

   sql = " alter table FICHAS ADD OBSERVACAOTRANSFERENCIAPRONTUARIO VARCHAR(255)"
   Banco.Execute sql

   sql = " alter table AMBULATORIAL ADD OBSERVACAOTRANSFERENCIAPRONTUARIO VARCHAR(255)"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOVIM_PRES_INT ADD SEQUENCIAITEMASSOCIADO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOVIM_PRES_AMB ADD SEQUENCIAITEMASSOCIADO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOVIM_PRES_EXT ADD SEQUENCIAITEMASSOCIADO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PRES_CAD_INTERVALO ADD HORAINICIO char(5)"
   Banco.Execute sql

   sql = " ALTER TABLE AGENDAMENTOCIRURGICO ADD FICHA_CIRURGIA INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE DADOSAIH ALTER COLUMN NumeroantAIH VARCHAR(13)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MOTIVOCOBRANCA (MOTIVOCOBRANCA, DESCRICAO, CONVENIO)"
   sql = sql & " SELECT TOP 1 41, 'OBITO COM DECLARACAO DE OBITO FORNECIDA PELO MEDICO ASSISTENTE', 0"
   sql = sql & " From MOTIVOCOBRANCA"
   sql = sql & " WHERE 41 NOT IN (SELECT MOTIVOCOBRANCA"
   sql = sql & "             FROM MOTIVOCOBRANCA)"
   Banco.Execute sql
   
   sql = " ALTER TABLE SUS_PROCEDIMENTO_CBO ADD VALORCBO MONEY"
   Banco.Execute sql
      
   sql = " ALTER TABLE PARAMETRO ADD SUS_AMB_VALORCBO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO ADD GESTANTEALTORISCO INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE USUARIO ADD REFERENCIA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE tmpINTERNACAO ALTER COLUMN PROCEDIMENTONOME VARCHAR(255)"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD BANCOLEITE_CC INT"
   Banco.Execute sql
   
   If Layout = 10 Then sql = " UPDATE PARAMETRO SET BANCOLEITE_CC = 24"
   Banco.Execute sql
      
   sql = "  ALTER TABLE SUS_PROCEDIMENTO ADD GUIAAUTORIZACAO INT"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes122010()
   On Error GoTo Erro
   
   sql = "ALTER TABLE SAME_MOVIMENTACAO_PRONTUARIO_FICHA ADD REGISTROANTIGO INT   "
   Banco.Execute sql
   
   sql = " ALTER TABLE CUSTO_PARAMETRO ADD ALMOXARIFADO_CC_AUX INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CUSTO_PARAMETRO ADD FARMACIA_CC_AUX INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CUSTO_PARAMETRO ADD SND_CC_AUX INT"
   Banco.Execute sql
   
   'CAMPINAS
   If Layout = 10 Then
      sql = " UPDATE CUSTO_PARAMETRO SET ALMOXARIFADO_CC_AUX = 74 "
      Banco.Execute sql
      
      sql = " UPDATE CUSTO_PARAMETRO SET FARMACIA_CC_AUX = 77 "
      Banco.Execute sql
   
      sql = " UPDATE CUSTO_PARAMETRO SET SND_CC_AUX = 66 "
      Banco.Execute sql
   End If
   
   
   'MARILIA
   If Layout = 21 Then
      sql = " UPDATE CUSTO_PARAMETRO SET ALMOXARIFADO_CC_AUX = 8305 "
      Banco.Execute sql
      
      sql = " UPDATE CUSTO_PARAMETRO SET FARMACIA_CC_AUX = 77 "
      'Banco.Execute Sql
   
      sql = " UPDATE CUSTO_PARAMETRO SET SND_CC_AUX = 66 "
      'Banco.Execute Sql
   End If
      
   sql = " ALTER TABLE CENTROCUSTO ADD COMPUTADORES INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE USUARIO ADD NAOPERMITEALTERACAOFINANCEIRO INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD TISS_GUIA_PRESTADOR_AMB INT"
   Banco.Execute sql

   sql = " ALTER TABLE CONVENIOS ADD TISS_GUIA_OPERADORA_AMB INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE FORNECEDORES ADD PERC_ISS MONEY"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE TMP_ALTERACAO_SALDO ("
   sql = sql & " PRODUTO               INT,"
   sql = sql & " CENTROCUSTO           INT,"
   sql = sql & " SALDOATUAL            MONEY,"
   sql = sql & " SALDOUNITARIOATUAL    MONEY,"
   sql = sql & " LOTE                  CHAR(20),"
   sql = sql & " VALIDADE              DATETIME,"
   sql = sql & " IP                    VARCHAR(100))"
   Banco.Execute sql

   sql = ""
   sql = sql & " CREATE TABLE savelog..ALTERACAO_SALDO_LOG ("
   sql = sql & " PRODUTO               INT,"
   sql = sql & " CENTROCUSTO           INT,"
   sql = sql & " SALDOATUAL            MONEY,"
   sql = sql & " SALDOUNITARIOATUAL    MONEY,"
   sql = sql & " SALDONOVO             MONEY,"
   sql = sql & " LOTE                  CHAR(20),"
   sql = sql & " VALIDADE              DATETIME,"
   sql = sql & " DATAOPERACAO          DATETIME,"
   sql = sql & " USUARIO               VARCHAR(100),"
   sql = sql & " IP                    VARCHAR(100))"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "    MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "    NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Alteração Saldo em Lote','mnuPrd_PCC_ASL','IMPORTADO 091210',"
   sql = sql & "    'mnuPrd_PCC_ASL',"
   sql = sql & "    1,2,'0102180000','mnuPrd_PCC_ASL',1"
   sql = sql & " From Menu"
   Banco.Execute sql

   sql = " ALTER TABLE TMPRELATORIOUSUARIOTRANSFERENCIA ADD ATUALIZACAO VARCHAR(255)"
   Banco.Execute sql

   sql = ""
   sql = sql & " CREATE TABLE [dbo].[CBHPMPORTE4]("
   sql = sql & " [TIPO] [int] NOT NULL,"
   sql = sql & " [CODIGO] [char](5) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & " [VALOR] [money] NOT NULL,"
   sql = sql & " [VALOR_LAUDO] [money] NULL,"
   sql = sql & " CONSTRAINT [PK_CBHPMPORTE4] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "     [TIPO] ASC,"
   sql = sql & "     [Codigo] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
      
   'INSERE OS PORTES NOVOS DA CBHPM4
   sql = ""
   sql = sql & " INSERT INTO CBHPMPORTE4 (TIPO, CODIGO, VALOR, VALOR_LAUDO ) " & Chr(13)
   sql = sql & " SELECT A.TIPO, A.CODIGO, A.VALOR, A.VALOR_LAUDO" & Chr(13)
   sql = sql & " FROM CBHPMPORTE A" & Chr(13)
   sql = sql & " WHERE A.CODIGO NOT IN (SELECT B.CODIGO" & Chr(13)
   sql = sql & "                          FROM CBHPMPORTE4 B" & Chr(13)
   sql = sql & "                          WHERE A.TIPO = B.TIPO " & Chr(13)
   sql = sql & "                          AND   A.CODIGO = B.CODIGO)" & Chr(13)
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MENU ( " & Chr(13)
   sql = sql & " MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA," & Chr(13)
   sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & " SELECT MAX(MENU)+1,'Consulta de Log de Sistema', 'mnuPar_Utl_LOG', ' ', 'mnuPar_Utl_LOG'," & Chr(13)
   sql = sql & " 1, 1, '0104090000', 'mnuPar_Utl_LOG', 1" & Chr(13)
   sql = sql & " FROM MENU "
   Banco.Execute sql

   sql = " ALTER TABLE INTERNO_DADOS_OBSTETRICO ADD DADOS_OB_SALA_AMAMENTOUPRIMHORA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO_DADOS_OBSTETRICO ADD DADOS_OB_POS_AMAMENTOUPRIMHORA INT"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes012011()
   On Error GoTo Erro

   sql = " ALTER TABLE TMP_EXPORTACAO_DMED ALTER COLUMN CPF BIGINT"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMP_EXPORTACAO_DMED ADD TIPOPACIENTE CHAR(10)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU ( " & Chr(13)
   sql = sql & " MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA," & Chr(13)
   sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & " SELECT MAX(MENU)+1,'DMED - Declaração de Serviços Médicos e Saúde',"
   sql = sql & "       'mnuPar_Utl_Exp_DME', ' ', 'mnuPar_Utl_Exp_DME',"
   sql = sql & "       1, 1, '0104050600', 'mnuPar_Utl_Exp_DME', 1"
   sql = sql & " FROM MENU"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[TMP_EXPORTACAO_DMED]("
   sql = sql & "    [SEQUENCIA] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & "    [CPF] [bigint] NULL,"
   sql = sql & "    [NOME] [varchar](200) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [VALOR] [money] NULL,"
   sql = sql & "    [IP] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [TIPOPACIENTE] [char](10) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_TMP_EXPORTACAO_DMED] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [Sequencia] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = " DROP TABLE [dbo].[TMPPRESTACAOMORADORES]"
   Banco.Execute sql

   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[TMPPRESTACAOMORADORES]("
   sql = sql & "    [FICHA] [int] NULL,"
   sql = sql & "    [NOME] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [DATA] [datetime] NULL,"
   sql = sql & "    [DESCRICAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [TIPO] [varchar](1) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [VALOR] [money] NULL,"
   sql = sql & "    [SALDO] [money] NULL,"
   sql = sql & "    [IP] [varchar](80) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [NF] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [MES] [varchar](7) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [QTDE] [int] NULL,"
   sql = sql & "    [VALORUNIT] [money] NULL,"
   sql = sql & "    [SALDOANTERIOR] [money] NULL,"
   sql = sql & "    [GRUPO] [int] NULL,"
   sql = sql & "    [GRUPODESCRICAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [ESTABELECIMENTO] [int] NULL,"
   sql = sql & "    [ESTABELECIMENTODESCRICAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [LEITOUNIDADE] [int] NULL,"
   sql = sql & "    [LEITODESCRICAO] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [SALDOGRUPO] [money] NULL,"
   sql = sql & "    [SALDOANTERIORGRUPO] [money] NULL"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMP_EXPORTACAO_DMED ADD LANCAMENTO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMP_EXPORTACAO_DMED ADD REGISTRO INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE PARAMETRO ADD REPETE_PRESCRICAO_ENFERMAGEM INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " ALTER TABLE PARAMETRO ADD ITEM_ENFERMAGEM_GENERICO INT"
   Banco.Execute sql

   sql = " ALTER TABLE PARAMETRO ADD DIAGNOSTICO_ENFERMAGEM_GENERICO INT"
   Banco.Execute sql

   sql = " ALTER TABLE FILANCAMENTORECEBIMENTO ADD CPF_DMED BIGINT "
   Banco.Execute sql

   sql = " ALTER TABLE FILANCAMENTORECEBIMENTO ADD NOME_DMED VARCHAR(100) "
   Banco.Execute sql

   sql = " alter table PRES_CAD_TUTOR ADD CHECAR_ENFERMAGEM INT "
   Banco.Execute sql
   
   sql = " alter table PRES_CAD_TUTOR ADD PRESCREVER_ENFERMAGEM INT "
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[QUIMIOTERAPIA]("
   sql = sql & "    [QUIMIOTERAPIA] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & "    [FICHA] [int] NOT NULL,"
   sql = sql & "    [DATA] [datetime] NULL,"
   sql = sql & "    [OBSERVACAO] [varchar](200) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_QUIMIOTERAPIA] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [Quimioterapia] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = ""
   sql = sql & " ALTER TABLE EXTERNO ADD QUIMIOTERAPIA INT"
   Banco.Execute sql
   
   sql = " "
   sql = sql & " ALTER TABLE EXTERNO ADD CONSTRAINT FK_EXTERNO_QUIMIOTERAPIA FOREIGN KEY (QUIMIOTERAPIA)"
   sql = sql & "  References Quimioterapia(Quimioterapia)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " ALTER TABLE fichas ADD PASTA VARCHAR(10)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE INDEX IX_Fichas_pasta ON fichas(pasta)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE VIEW V_RECUPERA_ULTIMAPASSAGE " & Chr(13)
   sql = sql & " AS " & Chr(13)
   sql = sql & " SELECT 'INTERNO' AS TIPO, DATAINTERNACAO, FICHA " & Chr(13)
   sql = sql & " FROM INTERNO WITH (NOLOCK) " & Chr(13)
   sql = sql & " Union All " & Chr(13)
   sql = sql & " SELECT 'EXTERNO' AS TIPO, DATAINTERNACAO, FICHA " & Chr(13)
   sql = sql & " FROM EXTERNO WITH (NOLOCK) " & Chr(13)
   sql = sql & " Union All " & Chr(13)
   sql = sql & " SELECT 'AMBULATORIAL' AS TIPO, DATAINTERNACAO, FICHA " & Chr(13)
   sql = sql & " FROM AMBULATORIAL WITH (NOLOCK) " & Chr(13)
   
   Banco.Execute sql
   
   sql = ""
   sql = sql & " ALTER TABLE INTERNO ADD PRAZOPERMANENCIA DATETIME"
   Banco.Execute sql
      
   sql = " ALTER TABLE TMPLUCROPERDA ADD CODIGOCENTROCUSTO INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU ( " & Chr(13)
   sql = sql & " MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA," & Chr(13)
   sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & " SELECT MAX(MENU)+1, 'Pesquisa Orientação Pré-Natal',"
   sql = sql & "        'mnuRec_POP', ' ', 'mnuRec_POP',"
   sql = sql & "        1, 1, '0218000000', 'mnuRec_POP', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU ( " & Chr(13)
   sql = sql & " MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA," & Chr(13)
   sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & " SELECT MAX(MENU)+1, 'Quimioterapia',"
   sql = sql & "        'mnuFat_SUS_Amb_Qui', ' ', 'mnuFat_SUS_Amb_Qui',"
   sql = sql & "        1, 1, '0723000000', 'mnuFat_SUS_Amb_Qui', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[PERGUNTAS]( " & Chr(13)
   sql = sql & "   [PERGUNTA] [int] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "   [DESCRICAO] [varchar](500) COLLATE Latin1_General_CI_AS NULL, " & Chr(13)
   sql = sql & "   [ATIVA] [bit] NULL CONSTRAINT [DF__PERGUNTAS__ATIVA__174ABBBB]  DEFAULT ((1)), " & Chr(13)
   sql = sql & "   [NUMPERGUNTA] [int] NULL, " & Chr(13)
   sql = sql & "   [ATUALIZACAO] [varchar](40) COLLATE Latin1_General_CI_AS NULL, " & Chr(13)
   sql = sql & "   [TIPORESPOSTA] [varchar](20) COLLATE Latin1_General_CI_AS NULL, " & Chr(13)
   sql = sql & " CONSTRAINT [PK_PERGUNTAS] PRIMARY KEY NONCLUSTERED " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "   [Pergunta] Asc " & Chr(13)
   sql = sql & ")WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY] " & Chr(13)
   sql = sql & ") ON [PRIMARY] " & Chr(13)
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[RESPOSTAS]( " & Chr(13)
   sql = sql & "   [PERGUNTA] [int] NOT NULL, " & Chr(13)
   sql = sql & "   [FICHA] [int] NOT NULL, " & Chr(13)
   sql = sql & "   [RESPOSTA] [varchar](500) COLLATE Latin1_General_CI_AS NULL, " & Chr(13)
   sql = sql & "   [NOTA] [tinyint] NULL, " & Chr(13)
   sql = sql & "   [DATA] [datetime] NULL, " & Chr(13)
   sql = sql & " CONSTRAINT [PK_RESPOSTAS] PRIMARY KEY NONCLUSTERED " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "   [PERGUNTA] ASC, " & Chr(13)
   sql = sql & "   [Ficha] Asc " & Chr(13)
   sql = sql & ")WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY] " & Chr(13)
   sql = sql & ") ON [PRIMARY] " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[RESPOSTAS]  WITH CHECK ADD  CONSTRAINT [FK_RESPOSTAS_FICHA] FOREIGN KEY([FICHA]) " & Chr(13)
   sql = sql & "References [dbo].[Fichas]([Ficha]) " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[RESPOSTAS]  WITH CHECK ADD  CONSTRAINT [FK_RESPOSTAS_PERGUNTA] FOREIGN KEY([PERGUNTA]) " & Chr(13)
   sql = sql & "References [dbo].[PERGUNTAS]([Pergunta]) " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE interno ADD PSI_ORDEMJUDICIAL BIT"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes022011()
   On Error GoTo Erro

   sql = "ALTER TABLE interno ADD PSI_ORDEMJUDICIAL BIT "
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD QUANTIDADE_CARACTER_CARTEIRINHA TINYINT "
   Banco.Execute sql
      
   If Layout = 35 Then   'guariba
      sql = "UPDATE CONVENIOS SET " & Chr(13)
      sql = sql & "QUANTIDADE_CARACTER_CARTEIRINHA = 7 " & Chr(13)
      sql = sql & "WHERE CONVENIO = 1 "
      Banco.Execute sql
      
      sql = "UPDATE CONVENIOS SET " & Chr(13)
      sql = sql & "QUANTIDADE_CARACTER_CARTEIRINHA = 8 " & Chr(13)
      sql = sql & "WHERE CONVENIO = 3 "
      Banco.Execute sql
   End If
   
   sql = ""
   sql = sql & "CREATE TABLE [dbo].[DMED]("
   sql = sql & "   [DMED] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & "   [NOME] [varchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,"
   sql = sql & "   [CPF] [varchar](18) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,"
   sql = sql & "   [DATA] [datetime] NOT NULL,"
   sql = sql & "   [VALOR] [money] NOT NULL,"
   sql = sql & "   [ATUALIZACAO] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,"
   sql = sql & " CONSTRAINT [PK_DMED] PRIMARY KEY CLUSTERED"
   sql = sql & "("
   sql = sql & "   [DMED] Asc"
   sql = sql & ")WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & ") ON [PRIMARY]"
   Banco.Execute sql

   sql = ""
   sql = sql & " CREATE TABLE [dbo].[TmpTer_Rendimento]("
   sql = sql & "    [MES] [int] NULL,"
   sql = sql & "    [DESCRICAOMES] [varchar](30) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [VALORBRUTO] [money] NULL,"
   sql = sql & "    [VALORDESCONTO] [money] NULL,"
   sql = sql & "    [VALORINSS] [money] NULL,"
   sql = sql & "    [VALORIRRF] [money] NULL,"
   sql = sql & "    [VALORPIS] [money] NULL,"
   sql = sql & "    [VALORCOFINS] [money] NULL,"
   sql = sql & "    [VALORCSSL] [money] NULL,"
   sql = sql & "    [TERCEIRO] [int] NULL,"
   sql = sql & "    [NOME] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [EMPRESA] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [CPF] [varchar](30) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [CGC] [varchar](30) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [ANO] [int] NULL,"
   sql = sql & "    [TIPOPESSOA] [int] NULL,"
   sql = sql & "    [CIDADE] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [IP] [varchar](255) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [CODIGORETENCAO] [int] NULL,"
   sql = sql & "    [GRUPO] [int] NULL,"
   sql = sql & "    [VALOR] [money] NULL,"
   sql = sql & "    [SEQUENCIA] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & " PRIMARY KEY CLUSTERED"
   sql = sql & "  ("
   sql = sql & "     [Sequencia] Asc"
   sql = sql & "  )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & "  ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "ALTER TABLE PMOR_PATRIMONIORATEADO "
   sql = sql & "ADD QUANTIDADE DECIMAL(18,3) "
   Banco.Execute sql
   
   sql = "ALTER TABLE PMOR_MOVIMENTO "
   sql = sql & "ADD NUMERONOTA VARCHAR(60)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PMOR_MOVIMENTO "
   sql = sql & "ADD PERCARREDONDADO DECIMAL(18,2)"
   Banco.Execute sql

   sql = "ALTER TABLE PMOR_MOVIMENTO "
   sql = sql & "ADD PERC DECIMAL(31,15)"
   Banco.Execute sql

   sql = "ALTER TABLE PMOR_MOVIMENTO "
   sql = sql & "ADD FORNECEDOR INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PMOR_PATRIMONIORATEADO "
   sql = sql & "ALTER COLUMN [PERCENTUAL] [money] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PMOR_PATRIMONIORATEADO "
   sql = sql & "ALTER COLUMN [NUMERONOTA] [varchar](60) COLLATE Latin1_General_CI_AS NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PMOR_PATRIMONIORATEADO "
   sql = sql & "ALTER COLUMN [TOTALMORADOR] [money] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PMOR_PATRIMONIORATEADO "
   sql = sql & "ALTER COLUMN [UNIDADE] [int] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PMOR_PATRIMONIORATEADO "
   sql = sql & "ALTER COLUMN [ESTABELECIMENTO] [int] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PMOR_PATRIMONIORATEADO "
   sql = sql & "ALTER COLUMN [QUANTIDADE] [decimal](18, 3) NULL"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU ( " & Chr(13)
   sql = sql & " MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA," & Chr(13)
   sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & " SELECT MAX(MENU)+1, 'Produtos',"
   sql = sql & "        'mnuFin_PCM_Rel_Pro', ' ', 'mnuFin_PCM_Rel_Pro',"
   sql = sql & "        1, 1, '0918050300', 'mnuFin_PCM_Rel_Pro', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOPROCEDIMENTO "
   sql = sql & "ADD UCO_PROCEDIMENTO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOPROCEDIMENTOHONORARIO "
   sql = sql & "ADD UCO_PROCEDIMENTO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PLANOCONVENIO ADD CONTAPLANO CHAR(15) "
   Banco.Execute sql
      
   sql = " ALTER TABLE SUS_PROCEDIMENTO ADD PERMITELANCARSEMFATURAR_INT INT   "
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPFICHAFISIO ADD NUMERO VARCHAR(5) "
   Banco.Execute sql
   
   sql = "ALTER TABLE ALTAS ADD CIDOBITO CHAR(4) "
   Banco.Execute sql
   
   sql = "ALTER TABLE FIADIANTAMENTO "
   sql = sql & "ADD CENTROCUSTO SMALLINT "
   sql = sql & "Constraint FK_FIADIANTAMENTO_CENTROCUSTO "
   sql = sql & "Foreign Key (CentroCusto) "
   sql = sql & "REFERENCES CENTROCUSTO(CENTROCUSTO) "
   Banco.Execute sql
   
   If Layout = 38 Then
      sql = ""
      sql = sql & "UPDATE PARAMETRO SET "
      sql = sql & "PRESCRICAO_PERIODO_INICIAL_MANHA = '07:59',"
      sql = sql & "PRESCRICAO_PERIODO_FINAL_MANHA = '12:59',"
      sql = sql & "PRESCRICAO_PERIODO_INICIAL_TARDE = '13:00',"
      sql = sql & "PRESCRICAO_PERIODO_FINAL_TARDE = '18:59',"
      sql = sql & "PRESCRICAO_PERIODO_INICIAL_NOITE = '19:00'"
      Banco.Execute sql
   End If
   
   sql = ""
   sql = sql & "CREATE TABLE [dbo].[LOCALTRANSFERENCIAAMBULATORIAL]("
   sql = sql & "   [LOCALTRANSFERENCIA] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & "   [NOME] [varchar](155) COLLATE Latin1_General_CI_AS NOT NULL"
   sql = sql & ") ON [PRIMARY]"
   Banco.Execute sql

   sql = " ALTER TABLE TMP_EXPORTACAO_DMED ADD BENEFICIARIO VARCHAR(60)"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMP_EXPORTACAO_DMED ADD CPFBENEFICIARIO BIGINT "
   Banco.Execute sql
   
   sql = " ALTER TABLE TMP_EXPORTACAO_DMED ADD NASCIMENTOBENEFICIARIO DATETIME "
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU ( " & Chr(13)
   sql = sql & " MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA," & Chr(13)
   sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & " SELECT MAX(MENU)+1, 'Atendimentos por Classificação de Risco',"
   sql = sql & "        'mnuRec_Rel_Acr', ' ', 'mnuRec_Rel_Acr',"
   sql = sql & "        1, 1, '0213130000', 'mnuRec_Rel_Acr', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = " ALTER TABLE REQUISICAOPRODUTO ADD IMPRIMIU INT "
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU ( " & Chr(13)
   sql = sql & " MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA," & Chr(13)
   sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & " SELECT MAX(MENU)+1, 'Lista de Espera',"
   sql = sql & "        'mnuRec_Les', ' ', 'mnuRec_Les',"
   sql = sql & "        1, 1, '0219000000', 'mnuRec_Les', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = "  CREATE TABLE [dbo].[LISTAESPERA]("
   sql = sql & "  [LISTAESPERA] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & "  [CENTROCUSTO] [smallint] NOT NULL,"
   sql = sql & "  [DATASOLICITACAO] [datetime] NOT NULL,"
   sql = sql & "  [PACIENTE] [varchar](155) NOT NULL,"
   sql = sql & "  [TELEFONE] [varchar](50) NULL,"
   sql = sql & "  [UNIDADEENCAMINHAMENTO] [int] NOT NULL,"
   sql = sql & "  [CID] [char](10) NULL,"
   sql = sql & "  [OBSERVACAO] [varchar](155) NULL,"
   sql = sql & "  [INTERNACAO] [int] NULL,"
   sql = sql & "  [INTERNACAO_ATUALIZACAO] [varchar](155) NULL,"
   sql = sql & "  [ATUALIZACAO] [varchar](155) NOT NULL,"
   sql = sql & "  CONSTRAINT [PK_LISTAESPERA] PRIMARY KEY CLUSTERED"
   sql = sql & "  ("
   sql = sql & "     [ListaEspera] Asc"
   sql = sql & "  )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
   sql = sql & "  ) ON [PRIMARY]"
   Banco.Execute sql

   sql = "  ALTER TABLE [dbo].[LISTAESPERA]  WITH NOCHECK ADD  CONSTRAINT [FK_LISTAESPERA_CentroCusto] FOREIGN KEY([CENTROCUSTO])"
   sql = sql & "  References [dbo].[CentroCusto]([CentroCusto]) NOT FOR REPLICATION"
   Banco.Execute sql
   
   sql = "alter table altas add CIDOBITO char(4)"
   Banco.Execute sql

   sql = "ALTER TABLE [dbo].[Altas]  WITH CHECK ADD  CONSTRAINT [FK_ALTAS_CID] FOREIGN KEY([CIDOBITO]) "
   sql = sql & " REFERENCES [dbo].[CID] ([CID])"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[CONT_CONVENIO_PLANO]( "
   sql = sql & "     [TABELA] [varchar](20) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "     [PROCEDIMENTO] [int] NOT NULL,"
   sql = sql & "     [TIPOPLANO] [int] NOT NULL,"
   sql = sql & "     [CONVENIO] [smallint] NOT NULL,"
   sql = sql & "     [TIPOLANCAMENTO] [smallint] NOT NULL,"
   sql = sql & "     [CONTA] [char](15) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "  CONSTRAINT [PK_CONT_CONVENIO_PLANO] PRIMARY KEY NONCLUSTERED"
   sql = sql & "     ("
   sql = sql & "     [TABELA] ASC,"
   sql = sql & "     [PROCEDIMENTO] ASC,"
   sql = sql & "     [TIPOPLANO] ASC,"
   sql = sql & "     [CONVENIO] ASC,"
   sql = sql & "     [TipoLancamento] Asc"
   sql = sql & "  )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & "  ) ON [PRIMARY]"
   
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes032011()
   On Error GoTo Erro
   
   sql = ""
   sql = sql & " INSERT INTO MENU ( " & Chr(13)
   sql = sql & " MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA," & Chr(13)
   sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & " SELECT MAX(MENU)+1, 'Atendimentos por Classificação de Risco',"
   sql = sql & "        'mnuRec_Rel_Acr', ' ', 'mnuRec_Rel_Acr',"
   sql = sql & "        1, 1, '0213130000', 'mnuRec_Rel_Acr', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
      
   sql = " ALTER TABLE PARAMETRO ADD VERIFICA_PROVISAO_FINANCEIRO INT "
   Banco.Execute sql
   
   If Layout = 40 Then
      sql = " UPDATE PARAMETRO SET VERIFICA_PROVISAO_FINANCEIRO = 1 "
      Banco.Execute sql
   End If
   
   sql = " ALTER TABLE FILANCAMENTO ADD PEDIDO_PROVISIONADO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_ENFERMAGEM_MOVIM_INT ADD CANCELADO BIT DEFAULT((0))"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD ESPECMEDICA INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " IF NOT EXISTS(SELECT 1 FROM SYSCOLUMNS WHERE ID = OBJECT_ID('CONT_CONVENIO_PLANO') AND NAME = 'TIPOREGISTRO')"
   sql = sql & " BEGIN"
   sql = sql & "   DROP TABLE CONT_CONVENIO_PLANO"
   sql = sql & "   CREATE TABLE [dbo].[CONT_CONVENIO_PLANO]("
   sql = sql & "      [TIPOREGISTRO] INT NOT NULL,"
   sql = sql & "      [TABELA] [varchar](20) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "      [PROCEDIMENTO] [int] NOT NULL,"
   sql = sql & "      [TIPOPLANO] [int] NOT NULL,"
   sql = sql & "      [CONVENIO] [smallint] NOT NULL,"
   sql = sql & "      [TIPOLANCAMENTO] [smallint] NOT NULL,"
   sql = sql & "      [CONTA] [char](15) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "      CONSTRAINT [PK_CONT_CONVENIO_PLANO] PRIMARY KEY NONCLUSTERED     ("
   sql = sql & "         [TIPOREGISTRO]ASC,"
   sql = sql & "         [TABELA] ASC,"
   sql = sql & "         [PROCEDIMENTO] ASC,"
   sql = sql & "         [TIPOPLANO] ASC,"
   sql = sql & "         [CONVENIO] ASC,"
   sql = sql & "         [TipoLancamento] Asc  )"
   sql = sql & "      WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]  ) ON [PRIMARY]"
   sql = sql & " End "
   Banco.Execute sql

   Banco.Execute " ALTER TABLE FILANCAMENTORECEBIMENTO ADD NAOEXPORTA_DMED INT " & Chr(13)

   sql = " ALTER TABLE tmpCOMPENSACAO ADD ENTRADA MONEY "
   Banco.Execute sql

   sql = " ALTER TABLE tmpCOMPENSACAO ADD SAIDA MONEY "
   Banco.Execute sql
   
   sql = " ALTER TABLE TmpTaxaOcupacao1 ADD QUANTIDADELEITOAGRUPADO INT"
   Banco.Execute sql

   sql = " ALTER TABLE USUARIO ADD UNIDADEENCAMINHAMENTO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE AGENDAMENTOCONSULTA ADD UNIDADEENCAMINHAMENTO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE USUARIO ADD UNIDADEENCAMINHAMENTO INT"
   Banco.Execute sql
   

   sql = " ALTER TABLE AGENDAMENTOCONSULTA ADD UNIDADEENCAMINHAMENTO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD ENDERECOPASSAGEM VARCHAR(50) " & Chr(13)
   sql = sql & "ALTER TABLE INTERNO ADD COMPLEMENTOPASSAGEM VARCHAR(20) " & Chr(13)
   sql = sql & "ALTER TABLE INTERNO ADD BAIRROPASSAGEM VARCHAR(155) " & Chr(13)
   sql = sql & "ALTER TABLE INTERNO ADD CIDADEPASSAGEM INT " & Chr(13)
   sql = sql & "ALTER TABLE INTERNO ADD UFPASSAGEM TINYINT " & Chr(13)
   sql = sql & "ALTER TABLE INTERNO ADD CEPPASSAGEM VARCHAR(9) " & Chr(13)
   sql = sql & "ALTER TABLE INTERNO ADD TELEFONEPASSAGEM VARCHAR(20) " & Chr(13)
   sql = sql & "ALTER TABLE INTERNO ADD TELEFONE2PASSAGEM VARCHAR(20) " & Chr(13)
   sql = sql & "ALTER TABLE INTERNO ADD NUMEROPASSAGEM VARCHAR(5) " & Chr(13)

   sql = sql & "ALTER TABLE INTERNO  WITH CHECK ADD  CONSTRAINT [FK_Interno_Cidades] FOREIGN KEY([CidadePASSAGEM]) " & Chr(13)
   sql = sql & "References [dbo].[Cidades]([Cidade]) " & Chr(13)

   sql = sql & "ALTER TABLE EXTERNO ADD ENDERECOPASSAGEM VARCHAR(50) " & Chr(13)
   sql = sql & "ALTER TABLE EXTERNO ADD COMPLEMENTOPASSAGEM VARCHAR(20) " & Chr(13)
   sql = sql & "ALTER TABLE EXTERNO ADD BAIRROPASSAGEM VARCHAR(155) " & Chr(13)
   sql = sql & "ALTER TABLE EXTERNO ADD CIDADEPASSAGEM INT " & Chr(13)
   sql = sql & "ALTER TABLE EXTERNO ADD UFPASSAGEM TINYINT " & Chr(13)
   sql = sql & "ALTER TABLE EXTERNO ADD CEPPASSAGEM VARCHAR(9) " & Chr(13)
   sql = sql & "ALTER TABLE EXTERNO ADD TELEFONEPASSAGEM VARCHAR(20) " & Chr(13)
   sql = sql & "ALTER TABLE EXTERNO ADD TELEFONE2PASSAGEM VARCHAR(20) " & Chr(13)
   sql = sql & "ALTER TABLE EXTERNO ADD NUMEROPASSAGEM VARCHAR(5) " & Chr(13)
   
   sql = sql & "ALTER TABLE EXTERNO  WITH CHECK ADD  CONSTRAINT [FK_Externo_Cidades] FOREIGN KEY([CidadePASSAGEM]) " & Chr(13)
   sql = sql & "References [dbo].[Cidades]([Cidade]) " & Chr(13)
   
   sql = sql & "ALTER TABLE AMBULATORIAL ADD ENDERECOPASSAGEM VARCHAR(50) " & Chr(13)
   sql = sql & "ALTER TABLE AMBULATORIAL ADD COMPLEMENTOPASSAGEM VARCHAR(20) " & Chr(13)
   sql = sql & "ALTER TABLE AMBULATORIAL ADD BAIRROPASSAGEM VARCHAR(155) " & Chr(13)
   sql = sql & "ALTER TABLE AMBULATORIAL ADD CIDADEPASSAGEM INT " & Chr(13)
   sql = sql & "ALTER TABLE AMBULATORIAL ADD UFPASSAGEM TINYINT " & Chr(13)
   sql = sql & "ALTER TABLE AMBULATORIAL ADD CEPPASSAGEM VARCHAR(9) " & Chr(13)
   sql = sql & "ALTER TABLE AMBULATORIAL ADD TELEFONEPASSAGEM VARCHAR(20) " & Chr(13)
   sql = sql & "ALTER TABLE AMBULATORIAL ADD TELEFONE2PASSAGEM VARCHAR(20) " & Chr(13)
   sql = sql & "ALTER TABLE AMBULATORIAL ADD NUMEROPASSAGEM VARCHAR(5) " & Chr(13)
   
   sql = sql & "ALTER TABLE AMBULATORIAL  WITH CHECK ADD  CONSTRAINT [FK_Ambulatorial_Cidades] FOREIGN KEY([CidadePASSAGEM]) " & Chr(13)
   sql = sql & "References [dbo].[Cidades]([Cidade]) " & Chr(13)
   
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD LEITOINTERNACAO VARCHAR(10) " & Chr(13)
   
   Banco.Execute sql

   sql = sql & "ALTER TABLE AMBULATORIAL ADD LEITOINTERNACAO VARCHAR(10) " & Chr(13)
   
   Banco.Execute sql

   sql = sql & "ALTER TABLE EXTERNO ADD LEITOINTERNACAO VARCHAR(10) " & Chr(13)
   
   Banco.Execute sql
   
   sql = "CREATE TABLE DRS_LEITO_ENVIO("
   sql = sql & "LEITO VARCHAR(10),"
   sql = sql & "DATAENVIO DATETIME,"
   sql = sql & "CONSTRAINT [FK_DRS_LEITO] FOREIGN KEY(LEITO)"
   sql = sql & "REFERENCES [dbo].[QUARTOS] ([LEITO]))"
   
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Relatório DRS', 'mnuSam_Est_RDR', ' ', 'mnuSam_Est_RDR',"
   sql = sql & "        1, 1, '0302160000', 'mnuSam_Est_RDR', 1"
   sql = sql & " FROM MENU "
   
   Banco.Execute sql

   sql = " alter table tmpPSICOLIVRO ADD LOTE VARCHAR(15)"
   Banco.Execute sql

   sql = " alter table tmpPSICOLIVRO ADD VALIDADELOTE DATETIME"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE SAVELOG..BACKUP_LOG("
   sql = sql & "    [SEQUENCIA] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & "    [DATAENTRADA] [datetime] NULL,"
   sql = sql & "    [DATASAIDA] [datetime] NULL,"
   sql = sql & "  CONSTRAINT [PK_BACKUP_LOG] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [Sequencia] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = " ALTER TABLE TmpTABELA_FECHAMENTOCONTABIL add HISTORICO VARCHAR(250)"
   Banco.Execute sql
   
   sql = "EXEC SP_RENAME 'DBO.CONVENIOS.QUANTIDADE_CARACTER_CARTEIRINHA', 'QUANTIDADE_CARACTER_GUIA', 'COLUMN'"
   
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes042011()
   On Error GoTo Erro

   sql = "ALTER TABLE ALTAS ADD CIDOBITO CHAR(4) "
   Banco.Execute sql
   
   sql = " ALTER TABLE TER_TERCEIRO_IMPOSTO ADD PERCENTUAL_ISS MONEY"
   Banco.Execute sql
   
   sql = "CREATE TABLE DRS_LEITO_ENVIO("
   sql = sql & "LEITO VARCHAR(10),"
   sql = sql & "DATAENVIO DATETIME,"
   sql = sql & "CONSTRAINT [FK_DRS_LEITO] FOREIGN KEY(LEITO)"
   sql = sql & "REFERENCES [dbo].[QUARTOS] ([LEITO]))"
   
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Relatório DRS', 'mnuSam_Est_RDR', ' ', 'mnuSam_Est_RDR',"
   sql = sql & "        1, 1, '0302160000', 'mnuSam_Est_RDR', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql

   sql = ""
   sql = sql & " CREATE TABLE [dbo].[FISALDOBANCODIARIO]("
   sql = sql & "    [BANCO] [int] NOT NULL,"
   sql = sql & "    [DATA] [datetime] NOT NULL,"
   sql = sql & "    [SALDO] [money] NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](255) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [IP] [varchar](255) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_FISALDOBANCODIARIO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [BANCO] ASC,"
   sql = sql & "    [DATA] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = "ALTER TABLE DRS_LEITO_ENVIO DROP CONSTRAINT FK_DRS_LEITO"
   Banco.Execute sql
      
   sql = " ALTER TABLE convenios ADD TISS_EXPORTA_CNPJ_FONTEPAGADORA INT"
   Banco.Execute sql
   
   If Layout = 17 Then
      sql = ""
      sql = sql & " UPDATE CONVENIOS SET TISS_EXPORTA_CNPJ_FONTEPAGADORA = 1 "
      Banco.Execute sql
   End If
   
   sql = " ALTER TABLE CENTROCUSTO ADD CONTROLAESTOQUEMINIMO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONTROLEVISITA ADD ENDERECO VARCHAR(255)"
   Banco.Execute sql

   sql = "ALTER TABLE CONTROLEVISITA ADD NUMERO VARCHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONTROLEVISITA ADD BAIRRO VARCHAR(155)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONTROLEVISITA ADD CIDADE INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONTROLEVISITA ADD TELEFONE VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONTROLEVISITA ADD CONSTRAINT [FK_CONTROLEVISITA_CIDADE] FOREIGN KEY(CIDADE) REFERENCES CIDADES (CIDADE)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCOMPENSACAO ALTER COLUMN DESCRICAO VARCHAR(255)"
   Banco.Execute sql
      
   sql = " ALTER TABLE INTERNO ADD PSI_ORDEMJUDICIAL INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPFLUXOCAIXA_TIPOBAIXA ADD TIPO VARCHAR(30)"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPFLUXOCAIXA_TIPOBAIXA ADD ORDEM INT "
   Banco.Execute sql
   
   
   sql = " ALTER TABLE USUARIO ADD PERMITE_APENAS_LANCAR_CONTASRECEBER INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Saldo Produto Mín e Máx', 'mnuPrd_Rel_PSa', ' ', 'mnuPrd_Rel_PSa',"
   sql = sql & "        1, 2, '0112100000', 'mnuPrd_Rel_PSa', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
      
   sql = " ALTER TABLE TER_EVENTO ADD NAOEXPORTADIRF INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & "CREATE TABLE savelog..CONVENIOPROCEDIMENTO_LOG ( " & Chr(13)
   sql = sql & " [CONVENIO] [smallint] NOT NULL, " & Chr(13)
   sql = sql & " [PROCEDIMENTO] [int] NOT NULL, " & Chr(13)
   sql = sql & " [UNIDADE] [int] NOT NULL, " & Chr(13)
   sql = sql & " [TIPOREGISTRO] [int] NOT NULL, " & Chr(13)
   sql = sql & " [VALOR] [money] NULL DEFAULT (0), " & Chr(13)
   sql = sql & " [VALORANTIGO] [money] NULL DEFAULT (0), " & Chr(13)
   sql = sql & " [VALORANESTESISTA] [money] NULL DEFAULT (0), " & Chr(13)
   sql = sql & " [VALORANESTESISTAANTIGO] [money] NULL DEFAULT (0), " & Chr(13)
   sql = sql & " [CH] [money] NULL DEFAULT (0), " & Chr(13)
   sql = sql & " [CHANTIGO] [money] NULL DEFAULT (0), " & Chr(13)
   sql = sql & " [CHANESTESISTA] [money] NULL DEFAULT (0), " & Chr(13)
   sql = sql & " [CHANESTESISTAANTIGO] [money] NULL DEFAULT (0), " & Chr(13)
   sql = sql & " [ATUALIZACAO] [varchar](60) COLLATE Latin1_General_CI_AS NULL, " & Chr(13)
   sql = sql & " [VALORANTERIORPARTICULAR] [money] NULL, " & Chr(13)
   sql = sql & " [valorantparticular] [money] NULL, " & Chr(13)
   sql = sql & " [UCO_PROCEDIMENTO] [money] NULL, " & Chr(13)
   sql = sql & " [UCO_PROCEDIMENTOANTIGO] [money] NULL, " & Chr(13)
   sql = sql & " DATAALTERACAO datetime, " & Chr(13)
   sql = sql & " TIPOALTERACAO int) " & Chr(13)
   
   Banco.Execute sql
   
   sql = ""
   sql = sql & "CREATE TRIGGER TR_LOG_CONVENIOPROCEDIMENTO " & Chr(13)
   sql = sql & " ON CONVENIOPROCEDIMENTO " & Chr(13)
   sql = sql & " With ENCRYPTION " & Chr(13)
   sql = sql & " FOR INSERT,UPDATE,DELETE " & Chr(13)
   sql = sql & " AS " & Chr(13)
   sql = sql & " DECLARE @TIPOOPERACAO AS INT " & Chr(13)
   sql = sql & " DECLARE @GERALOG AS INT " & Chr(13)

   sql = sql & " -- 0 - INCLUSÃO " & Chr(13)
   sql = sql & " -- 1 - ALTERAÇÃO " & Chr(13)
   sql = sql & " -- 2 - EXCLUSÃO " & Chr(13)

   sql = sql & " SET @GERALOG = 0 " & Chr(13)
   sql = sql & " SET @TIPOOPERACAO = 2 " & Chr(13)

   sql = sql & " IF EXISTS(SELECT CONVENIO FROM INSERTED) " & Chr(13)
   sql = sql & " BEGIN " & Chr(13)
   sql = sql & "    IF NOT EXISTS(SELECT CONVENIO FROM DELETED) " & Chr(13)
   sql = sql & "       SET @TIPOOPERACAO=0 " & Chr(13)
   sql = sql & "    Else " & Chr(13)
   sql = sql & "       SET @TIPOOPERACAO=1 " & Chr(13)
   sql = sql & " End " & Chr(13)
   sql = sql & "  " & Chr(13)
   sql = sql & " IF @TIPOOPERACAO = 1 " & Chr(13)
   sql = sql & "    IF (( " & Chr(13)
   sql = sql & "        (SELECT UCO_PROCEDIMENTO FROM INSERTED) <> (SELECT UCO_PROCEDIMENTO FROM DELETED) " & Chr(13)
   sql = sql & "       OR (SELECT VALOR FROM INSERTED) <> (SELECT VALOR FROM DELETED) " & Chr(13)
   sql = sql & "       OR (SELECT CH FROM INSERTED) <> (SELECT CH FROM DELETED) " & Chr(13)
   sql = sql & "       OR (SELECT VALORANESTESISTA FROM INSERTED) <> (SELECT VALORANESTESISTA FROM DELETED) " & Chr(13)
   sql = sql & "       OR (SELECT CHANESTESISTA FROM INSERTED) <> (SELECT CHANESTESISTA FROM DELETED) " & Chr(13)
   sql = sql & "       )) " & Chr(13)
   sql = sql & "    BEGIN " & Chr(13)
   sql = sql & "       SELECT 'DIFERENTE' " & Chr(13)
   sql = sql & "       SET @GERALOG = 1 " & Chr(13)
   sql = sql & "    End " & Chr(13)
   sql = sql & "    Else " & Chr(13)
   sql = sql & "       SELECT 'IGUAL' " & Chr(13)
   sql = sql & " ELSE IF @TIPOOPERACAO = 2 " & Chr(13)
   sql = sql & "    BEGIN " & Chr(13)
   sql = sql & "       SELECT 'DELETE' " & Chr(13)
   sql = sql & "       SET @GERALOG = 1 " & Chr(13)
   sql = sql & "    End " & Chr(13)
   sql = sql & " Else " & Chr(13)
   sql = sql & "    BEGIN " & Chr(13)
   sql = sql & "       SELECT 'INSERT' " & Chr(13)
   sql = sql & "    End " & Chr(13)

   sql = sql & " IF @GERALOG = 1 " & Chr(13)
   sql = sql & "    IF @TIPOOPERACAO = 2    --EXCLUSÃO " & Chr(13)
   sql = sql & "    BEGIN " & Chr(13)
   sql = sql & "       INSERT INTO Savelog..CONVENIOPROCEDIMENTO_LOG ( " & Chr(13)
   sql = sql & "          CONVENIO, " & Chr(13)
   sql = sql & "          PROCEDIMENTO, " & Chr(13)
   sql = sql & "          UNIDADE, " & Chr(13)
   sql = sql & "          TIPOREGISTRO, " & Chr(13)
   sql = sql & "          VALOR, " & Chr(13)
   sql = sql & "          VALORANTIGO, " & Chr(13)
   sql = sql & "          VALORANESTESISTA, " & Chr(13)
   sql = sql & "          VALORANESTESISTAANTIGO, " & Chr(13)
   sql = sql & "          CH, " & Chr(13)
   sql = sql & "          CHANTIGO, " & Chr(13)
   sql = sql & "          CHANESTESISTA, " & Chr(13)
   sql = sql & "          CHANESTESISTAANTIGO, " & Chr(13)
   sql = sql & "          ATUALIZACAO, " & Chr(13)
   sql = sql & "          VALORANTERIORPARTICULAR, " & Chr(13)
   sql = sql & "          valorantparticular, " & Chr(13)
   sql = sql & "          UCO_PROCEDIMENTO, " & Chr(13)
   sql = sql & "          UCO_PROCEDIMENTOANTIGO, " & Chr(13)
   sql = sql & "          DATAALTERACAO, " & Chr(13)
   sql = sql & "          TIPOALTERACAO) " & Chr(13)
   sql = sql & "       SELECT " & Chr(13)
   sql = sql & "          A.CONVENIO, " & Chr(13)
   sql = sql & "          A.PROCEDIMENTO, " & Chr(13)
   sql = sql & "          A.UNIDADE, " & Chr(13)
   sql = sql & "          A.TIPOREGISTRO, " & Chr(13)
   sql = sql & "          B.VALOR, " & Chr(13)
   sql = sql & "          A.VALOR, " & Chr(13)
   sql = sql & "          B.VALORANESTESISTA, " & Chr(13)
   sql = sql & "          A.VALORANESTESISTA, " & Chr(13)
   sql = sql & "          B.CH, " & Chr(13)
   sql = sql & "          A.CH, " & Chr(13)
   sql = sql & "          B.CHANESTESISTA, " & Chr(13)
   sql = sql & "          A.CHANESTESISTA, " & Chr(13)
   sql = sql & "          A.ATUALIZACAO, " & Chr(13)
   sql = sql & "          A.VALORANTERIORPARTICULAR, " & Chr(13)
   sql = sql & "          A.valorantparticular, " & Chr(13)
   sql = sql & "          B.UCO_PROCEDIMENTO, " & Chr(13)
   sql = sql & "          A.UCO_PROCEDIMENTO, " & Chr(13)
   sql = sql & "          GETDATE(), " & Chr(13)
   sql = sql & "          @TIPOOPERACAO " & Chr(13)
   sql = sql & "       FROM DELETED A LEFT JOIN INSERTED B ON A.CONVENIO   = B.CONVENIO " & Chr(13)
   sql = sql & "                       AND A.PROCEDIMENTO = B.PROCEDIMENTO " & Chr(13)
   sql = sql & "                      AND A.UNIDADE = B.UNIDADE " & Chr(13)
   sql = sql & "                      AND A.TIPOREGISTRO = B.TIPOREGISTRO " & Chr(13)
   sql = sql & "    End " & Chr(13)
   sql = sql & "    Else " & Chr(13)
   sql = sql & "     " & Chr(13)
   sql = sql & "    BEGIN " & Chr(13)
   sql = sql & "       INSERT INTO savelog..CONVENIOPROCEDIMENTO_LOG ( " & Chr(13)
   sql = sql & "          CONVENIO, " & Chr(13)
   sql = sql & "          PROCEDIMENTO, " & Chr(13)
   sql = sql & "          UNIDADE, " & Chr(13)
   sql = sql & "          TIPOREGISTRO, " & Chr(13)
   sql = sql & "          VALOR, " & Chr(13)
   sql = sql & "          VALORANTIGO, " & Chr(13)
   sql = sql & "          VALORANESTESISTA, " & Chr(13)
   sql = sql & "          VALORANESTESISTAANTIGO, " & Chr(13)
   sql = sql & "          CH, " & Chr(13)
   sql = sql & "          CHANTIGO, " & Chr(13)
   sql = sql & "          CHANESTESISTA, " & Chr(13)
   sql = sql & "          CHANESTESISTAANTIGO, " & Chr(13)
   sql = sql & "          ATUALIZACAO, " & Chr(13)
   sql = sql & "          VALORANTERIORPARTICULAR, " & Chr(13)
   sql = sql & "          valorantparticular, " & Chr(13)
   sql = sql & "          UCO_PROCEDIMENTO, " & Chr(13)
   sql = sql & "          UCO_PROCEDIMENTOANTIGO, " & Chr(13)
   sql = sql & "          DATAALTERACAO, " & Chr(13)
   sql = sql & "          TIPOALTERACAO) " & Chr(13)
   sql = sql & "       SELECT " & Chr(13)
   sql = sql & "          B.CONVENIO, " & Chr(13)
   sql = sql & "          B.PROCEDIMENTO, " & Chr(13)
   sql = sql & "          B.UNIDADE, " & Chr(13)
   sql = sql & "          B.TIPOREGISTRO, " & Chr(13)
   sql = sql & "          B.VALOR, " & Chr(13)
   sql = sql & "          A.VALOR, " & Chr(13)
   sql = sql & "          B.VALORANESTESISTA, " & Chr(13)
   sql = sql & "          A.VALORANESTESISTA, " & Chr(13)
   sql = sql & "          B.CH, " & Chr(13)
   sql = sql & "          A.CH, " & Chr(13)
   sql = sql & "          B.CHANESTESISTA, " & Chr(13)
   sql = sql & "          A.CHANESTESISTA, " & Chr(13)
   sql = sql & "          B.ATUALIZACAO, " & Chr(13)
   sql = sql & "          B.VALORANTERIORPARTICULAR, " & Chr(13)
   sql = sql & "          B.valorantparticular, " & Chr(13)
   sql = sql & "          B.UCO_PROCEDIMENTO, " & Chr(13)
   sql = sql & "          A.UCO_PROCEDIMENTO, " & Chr(13)
   sql = sql & "          GETDATE(), " & Chr(13)
   sql = sql & "          @TIPOOPERACAO " & Chr(13)
   sql = sql & "       FROM DELETED A INNER JOIN INSERTED B ON A.CONVENIO   = B.CONVENIO " & Chr(13)
   sql = sql & "                       AND A.PROCEDIMENTO = B.PROCEDIMENTO " & Chr(13)
   sql = sql & "                      AND A.UNIDADE = B.UNIDADE " & Chr(13)
   sql = sql & "                      AND A.TIPOREGISTRO = B.TIPOREGISTRO " & Chr(13)
   sql = sql & "    END " & Chr(13)
   
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_EXT ADD DATAENVIO DATETIME"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_EXT ADD REENVIAR BIT"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_INT ADD DATAENVIO DATETIME"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_INT ADD REENVIAR BIT"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_AMB ADD DATAENVIO DATETIME"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_AMB ADD REENVIAR BIT"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD INTEGRA_SLINE BIT"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETROS ADD CAMINHO_XML_SLINE VARCHAR(255)"
   
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_EXT_SEQUENCIA] ON [dbo].[LAU_MOVIM_EXT] " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "[Sequencia] Asc " & Chr(13)
   sql = sql & ")"
   
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_INT_SEQUENCIA] ON [dbo].[LAU_MOVIM_INT] "
   sql = sql & "( " & Chr(13)
   sql = sql & "[Sequencia] Asc " & Chr(13)
   sql = sql & ")"
   
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_AMB_SEQUENCIA] ON [dbo].[LAU_MOVIM_AMB] "
   sql = sql & "( " & Chr(13)
   sql = sql & "[Sequencia] Asc " & Chr(13)
   sql = sql & ")"
   
   Banco.Execute sql
   
   sql = "INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Resultados de Exames (NefroData)', 'mnuPar_Utl_Exp_Nef', ' ', 'mnuPar_Utl_Exp_Nef',"
   sql = sql & "        1, 1, '0104050600', 'mnuPar_Utl_Exp_Nef', 1"
   sql = sql & " FROM MENU"
   Banco.Execute sql
   
   sql = "UPDATE MENU SET NOMECAPTION = 'Resultado de Exames' WHERE NOMESUBNOVO = 'mnuPar_Utl_Exp_Nef'"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes052011()
   On Error GoTo Erro

   sql = " ALTER TABLE PRES_TMP_MEDICACAO ADD OBSERVACAO VARCHAR(250)"
   Banco.Execute sql

   sql = " ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT ADD MEDICACAO_ADMINISTRADA INT"
   Banco.Execute sql
   
   If Layout = 42 Then
      sql = ""
      sql = sql & "UPDATE PARAMETRO SET "
      sql = sql & "PRESCRICAO_PERIODO_INICIAL_MANHA = NULL,"
      sql = sql & "PRESCRICAO_PERIODO_FINAL_MANHA = NULL,"
      sql = sql & "PRESCRICAO_PERIODO_INICIAL_TARDE = NULL,"
      sql = sql & "PRESCRICAO_PERIODO_FINAL_TARDE = NULL,"
      sql = sql & "PRESCRICAO_PERIODO_INICIAL_NOITE = NULL,"
      sql = sql & "PRESCRICAO_PERIODO_FINAL_NOITE = NULL "
      Banco.Execute sql
   End If
   
   sql = " ALTER TABLE TER_LANCAMENTO_ITEM ADD DOC_IR CHAR(15)"
   Banco.Execute sql
   
   sql = " ALTER TABLE TER_PARAMETRO ADD IR_DOCUMENTO INT DEFAULT(0)"
   Banco.Execute sql
   
   sql = "CREATE TABLE PSI_TIPOPACIENTE("
   sql = sql & "CODIGO INT, "
   sql = sql & "DESCRICAO VARCHAR(30), "
   sql = sql & "CONSTRAINT PK_PSI_TIPOPACIENTE PRIMARY KEY (CODIGO))"
   
   Banco.Execute sql
   
   sql = "INSERT INTO PSI_TIPOPACIENTE (CODIGO, DESCRICAO) VALUES (1, 'AGUDO')"
   
   Banco.Execute sql
   
   sql = "INSERT INTO PSI_TIPOPACIENTE (CODIGO, DESCRICAO) VALUES (2, 'CRÔNICO')"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD PSI_TIPOPACIENTE INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPFLUXOCAIXA_TIPOBAIXA ADD IP VARCHAR(100)"
   Banco.Execute sql
      
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[FISALDOBANCODIARIO_TIPOBAIXA]("
   sql = sql & "    [BANCO] [int] NOT NULL,"
   sql = sql & "    [DATA] [datetime] NOT NULL,"
   sql = sql & "    [SALDO] [money] NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](255) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [IP] [varchar](255) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [TIPOBAIXA] [varchar](30) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "    [TIPO] [varchar](30) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "  CONSTRAINT [PK_FISALDOBANCODIARIO_TIPOBAIXA] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [BANCO] ASC,"
   sql = sql & "    [DATA] ASC,"
   sql = sql & "    [TIPOBAIXA] ASC,"
   sql = sql & "    [Tipo] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD TUSS_EXIBECODIGO_RELATORIO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_ALTERACAO_SALDO ADD SALDONOVO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_ALTERACAO_SALDO ADD SALDOUNITARIONOVO MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO ADD PSI_NATUREZAATENDIMENTO INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU(MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "       NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Relação de Atendimentos', 'mnuCRe_Ate', ' ', 'mnuCRe_Ate',"
   sql = sql & "    1, 4, '0404000000', 'mnuCRe_Ate', 1"
   sql = sql & " From Menu"
   Banco.Execute sql

   sql = " ALTER TABLE INTERNO ADD PROCSUS_CIHA INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE EXTERNO ADD PROCSUS_CIHA INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD PROCSUS_CIHA INT "
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD APAC_QUIMIO_ESTADIOUICC INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [SAVELOG]..[CONTAGEMESTOQUE]("
   sql = sql & "    DATA  DATETIME,"
   sql = sql & "    [PRODUTO] [int] NULL,"
   sql = sql & "    [CENTROCUSTO] [int] NULL,"
   sql = sql & "    [SALDOATUAL] [money] NULL,"
   sql = sql & "    [SALDOUNITARIOATUAL] [money] NULL,"
   sql = sql & "    [LOTE] [char](20) ,"
   sql = sql & "    [VALIDADE] [datetime] NULL,"
   sql = sql & "    [IP] [varchar](100) ,"
   sql = sql & "    [SALDONOVO] [money] NULL,"
   sql = sql & "    [SALDOUNITARIONOVO] [money] NULL,"
   sql = sql & "    Usuario VarChar(100),"
   sql = sql & "    ATUALIZACAO VarChar(100),"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "ALTER TABLE LISTAESPERA ADD CID2 CHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE LISTAESPERA ADD MEDICOSOLICITANTE SMALLINT"
   Banco.Execute sql
   
   sql = "ALTER TABLE LISTAESPERA ADD MEDICOAVALIADOR SMALLINT"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes062011()
   On Error GoTo Erro

   sql = " ALTER TABLE CONVENIODATA_PERIODO ADD ARQUIVOXMLENVIADO TEXT"
   Banco.Execute sql

   sql = " ALTER TABLE CONVENIODATA_PERIODO ADD ARQUIVOXMLRETORNO TEXT"
   Banco.Execute sql

   sql = " ALTER TABLE CONVENIODATA_PERIODO ADD DATAENVIOXML DATETIME"
   Banco.Execute sql

   sql = " ALTER TABLE CONVENIODATA_PERIODO ADD USUARIOENVIO VARCHAR(50)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE ENVIOXMLWEBSERVICE_LOG("
   sql = sql & " ANO INT,"
   sql = sql & " CONVENIO VARCHAR(50),"
   sql = sql & " MES VARCHAR(50),"
   sql = sql & " FATURA INT,"
   sql = sql & " ARQUIVOXMLENVIADO TEXT,"
   sql = sql & " ARQUIVOXMLRETORNO TEXT,"
   sql = sql & " DATAENVIOXML DATETIME,"
   sql = sql & " USUARIOENVIO VARCHAR(50))"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD APAC_QUIMIO_GRAHISTOPATOLOGICO_STR VARCHAR(100)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD APAC_QUIMIO_ESTADIO_STR VARCHAR(100)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD APAC_PROCEDIMENTO_SECUNDARIO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD APAC_QUANTIDADE_SECUNDARIO INT"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_EXTERNO_QUIMIOTERAPIA] ON [dbo].[Externo]"
   sql = sql & "( " & Chr(13)
   sql = sql & "   Quimioterapia Asc " & Chr(13)
   sql = sql & ")WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "ALTER TABLE externo ALTER COLUMN APAC_QUIMIO_ESQUEMA VARCHAR(255)"
   Banco.Execute sql
   
   sql = "ALTER TABLE SUS_TABELA_CIH DROP CONSTRAINT FK_SUS_TABELA_CIH_SUSInternos"
   Banco.Execute sql
   
   sql = " ALTER TABLE SUS_TABELA_CIH ALTER COLUMN CODIGO INT NOT NULL"
   Banco.Execute sql
   
   sql = " ALTER TABLE SUS_TABELA_CIH DROP CONSTRAINT PK_SUS_TABELA_CIH "
   Banco.Execute sql

   sql = " ALTER TABLE SUS_TABELA_CIH ADD CONSTRAINT PK_SUS_TABELA_CIH1 "
   sql = sql & " PRIMARY KEY NONCLUSTERED (CODIGOSUS, CODIGO, TABELA)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPRELACAONASCIDOS ADD ENDERECO VARCHAR(255)"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOVIM_PRES_amb ADD ORDEM INT"
   Banco.Execute sql

   sql = " ALTER TABLE MOVIM_PRES_EXT ADD ORDEM INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOVIM_PRES_INT ADD ORDEM INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE SUS_TABELA_CIH "
   sql = sql & " ADD TIPOTABELA INT NULL"
   Banco.Execute sql
   
   sql = "Update SUS_TABELA_CIH"
   sql = sql & " Set TipoTabela = 0"
   sql = sql & " Where Tabela <> 8"
   sql = sql & " AND TIPOTABELA IS NULL"
   Banco.Execute sql
   
   sql = "Update SUS_TABELA_CIH"
   sql = sql & " Set TipoTabela = 1"
   sql = sql & " Where Tabela = 8"
   sql = sql & " AND TIPOTABELA IS NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE SUS_TABELA_CIH"
   sql = sql & " ALTER COLUMN TIPOTABELA INT NOT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE SUS_TABELA_CIH"
   sql = sql & " DROP CONSTRAINT PK_SUS_TABELA_CIH1"
   Banco.Execute sql
   
   sql = "ALTER TABLE SUS_TABELA_CIH"
   sql = sql & " ADD CONSTRAINT PK_SUS_TABELA_CIH2 PRIMARY KEY NONCLUSTERED (CODIGOSUS, TABELA, CODIGO, TIPOTABELA)"
   Banco.Execute sql
   
   sql = "CREATE TABLE MOVIMENTACAOPACIENTE("
   sql = sql & " MOVIMENTACAOPACIENTE INT IDENTITY,"
   sql = sql & " REGISTRO INT,"
   sql = sql & " TIPO VARCHAR(20),"
   sql = sql & " DATAMOVIMENTACAO DATETIME,"
   sql = sql & " DATASOLICITACAO DATETIME,"
   sql = sql & " TIPOAMBULANCIA VARCHAR(20),"
   sql = sql & " MEDICO VARCHAR(150),"
   sql = sql & " ENFERMEIRO VARCHAR(150),"
   sql = sql & " LOCALTRANSFERENCIA VARCHAR(255),"
   sql = sql & " HIPOTESEDIAGNOSTICA VARCHAR(500),"
   sql = sql & " MOTIVOMOVIMENTACAO VARCHAR(500),"
   sql = sql & " atualizacao VarChar(100)"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_REGISTRO] ON MOVIMENTACAOPACIENTE"
   sql = sql & " ("
   sql = sql & " Registro"
   sql = sql & " )"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_DATAMOVIMENTACAO] ON MOVIMENTACAOPACIENTE"
   sql = sql & " ("
   sql = sql & " DATAMOVIMENTACAO"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPCONFERENCIASUSAIH ADD PROCEDIMENTOCONTA INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPCONFERENCIASUSAIH ADD PROCEDIMENTOCONTANOME VARCHAR(250)"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPCONFERENCIASUSAIH ADD DATAINTERNACAO DATETIME"
   Banco.Execute sql

   sql = " ALTER TABLE TMPCONFERENCIASUSAIH ADD DATAALTA DATETIME"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPCONFERENCIASUSAIH ADD DATAINTERNACAO DATETIME"
   Banco.Execute sql

   sql = " ALTER TABLE TMPCONFERENCIASUSAIH ADD DATAALTA DATETIME"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[CUSTOEVENTOFOLHA]( " & Chr(13)
   sql = sql & "[EVENTO] [int] NULL, " & Chr(13)
   sql = sql & "[DESCRICAO] [varchar](200) COLLATE Latin1_General_CI_AS NULL, " & Chr(13)
   sql = sql & "[ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL " & Chr(13)
   sql = sql & ") ON [PRIMARY]"
   
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[CUSTOPARAMETROEVENTOFOLHA]( " & Chr(13)
   sql = sql & "[TIPO] [int] NOT NULL, " & Chr(13)
   sql = sql & "[EVENTO] [int] NOT NULL, " & Chr(13)
   sql = sql & "[ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL, " & Chr(13)
   sql = sql & "CONSTRAINT [PK_CUSTOPARAMETROEVENTOFOLHA] PRIMARY KEY NONCLUSTERED " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "   [EVENTO] Asc " & Chr(13)
   sql = sql & ")WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY] " & Chr(13)
   sql = sql & ") ON [PRIMARY]"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE CUSTO_PARAMETRO " & Chr(13)
   sql = sql & "ADD SECAO_DEPARTAMENTO TINYINT " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE CUSTO_PARAMETRO " & Chr(13)
   sql = sql & "ADD EVENTO_SALARIO INT " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE CUSTO_PARAMETRO " & Chr(13)
   sql = sql & "ADD EVENTO_FERIASE13 INT " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE CUSTO_PARAMETRO " & Chr(13)
   sql = sql & "ADD EVENTO_FGTS INT " & Chr(13)
   Banco.Execute sql
   
   sql = "CREATE TABLE CUSTOCENTROCUSTO( " & Chr(13)
   sql = sql & "TIPO TINYINT, " & Chr(13)
   sql = sql & "CODIGO INT, " & Chr(13)
   sql = sql & "DESCRICAO VARCHAR(200), " & Chr(13)
   sql = sql & "ATUALIZACAO VARCHAR(100)) " & Chr(13)
   Banco.Execute sql
   
   sql = "CREATE TABLE CUSTOCENTROCUSTOASSOCIACAO( " & Chr(13)
   sql = sql & "TIPO TINYINT, " & Chr(13)
   sql = sql & "CENTROCUSTO INT, " & Chr(13)
   sql = sql & "CENTROCUSTOIMP INT, " & Chr(13)
   sql = sql & "ATUALIZAÇÃO VARCHAR(100), " & Chr(13)
   sql = sql & "CONSTRAINT [PK_CUSTOCENTROCUSTOASSOCIACAO] PRIMARY KEY NONCLUSTERED " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "   TIPO, CENTROCUSTO, CENTROCUSTOIMP ASC " & Chr(13)
   sql = sql & ")WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY] " & Chr(13)
   sql = sql & ") ON [PRIMARY] " & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[TMP_IMPORTACAO_PORTAL]("
   sql = sql & "    [PDC] [varchar](16) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [COTACAO] [int] NULL,"
   sql = sql & "    [PRODUTO] [int] NULL,"
   sql = sql & "    [QUANTIDADE] [int] NULL,"
   sql = sql & "    [CNPJ] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [VALORUNITARIO] [money] NULL,"
   sql = sql & "    [FORNECEDOR] [int] NULL"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   If Layout = 42 Then
      sql = " UPDATE PRODUTOTIPO SET ITEMSOUTILIZADOPRESCRICAO = 1"
      sql = sql & " WHERE CODIGO IN (10, 15)"
      Banco.Execute sql
   End If
   
   sql = ""
   sql = sql & " CREATE VIEW V_RECUPERA_CIHA_CONSOLIDADO" & Chr(13)
   sql = sql & " AS " & Chr(13)
   sql = sql & "     SELECT 5 AS TIPOREG, 'AMB' AS TIPO, A.DATAINTERNACAO," & Chr(13)
   sql = sql & "          ISNULL(G.CODIGOSUS_PROCEDIMENTO,0) AS PROCSUS," & Chr(13)
   sql = sql & "          COUNT(a.Registro) As Quantidade" & Chr(13)
   sql = sql & "     FROM   AMBULATORIAL A WITH(NOLOCK) INNER JOIN FICHAS B WITH(NOLOCK)                  ON A.FICHA = B.FICHA" & Chr(13)
   sql = sql & "                               INNER JOIN CONVENIOS D WITH(NOLOCK)               ON A.CONVENIO = D.CONVENIO" & Chr(13)
   sql = sql & "                               LEFT  JOIN COMPLEMENTAR E WITH(NOLOCK)            ON A.REGISTRO = E.REGISTRO" & Chr(13)
   sql = sql & "                               LEFT  JOIN DADOSGUIA F WITH(NOLOCK)               ON A.REGISTRO = F.REGISTRO" & Chr(13)
   sql = sql & "                                                                    AND F.TIPOREGISTRO=1" & Chr(13)
   sql = sql & "                               LEFT  JOIN SUS_PROCEDIMENTO G WITH(NOLOCK)        ON A.PROCSUS_CIHA =  G.CODIGOSUS_PROCEDIMENTO" & Chr(13)
   sql = sql & "    Where 1 = 1" & Chr(13)
   sql = sql & "    AND    D.TIPOCONVENIO <> 3" & Chr(13)
   sql = sql & "    AND    ISNULL(A.CANCELADO,0) = 0" & Chr(13)
   sql = sql & "    AND    A.REGISTRO>0" & Chr(13)
   sql = sql & "    AND EXISTS (SELECT CON.PROCEDIMENTO" & Chr(13)
   sql = sql & "                FROM SUS_PROCEDIMENTO_REGISTRO CON" & Chr(13)
   sql = sql & "                Where CON.Registro = 1" & Chr(13)
   sql = sql & "                 AND   LEFT(REPLICATE('0',10-LEN(LEFT(G.CODIGOSUS_PROCEDIMENTO,10))) + LTRIM(G.CODIGOSUS_PROCEDIMENTO),2) IN ('02', '03', '04', '05')" & Chr(13)
   sql = sql & "                AND   CON.PROCEDIMENTO = G.CODIGOSUS_PROCEDIMENTO )" & Chr(13)
   sql = sql & "   GROUP BY G.CODIGOSUS_PROCEDIMENTO, A.DATAINTERNACAO" & Chr(13)
   
   sql = sql & "    Union All" & Chr(13)
   
   sql = sql & "   SELECT 5 AS TIPOREG, 'AMB' AS TIPO, A.DATAINTERNACAO," & Chr(13)
   sql = sql & "         ISNULL(G.CODIGOSUS_PROCEDIMENTO,0) AS PROCSUS," & Chr(13)
   sql = sql & "         COUNT(a.Registro) As Quantidade" & Chr(13)
   sql = sql & "   FROM   EXTERNO      A WITH(NOLOCK) INNER JOIN FICHAS B WITH(NOLOCK)                  ON A.FICHA = B.FICHA" & Chr(13)
   sql = sql & "                              INNER JOIN CONVENIOS D WITH(NOLOCK)               ON A.CONVENIO = D.CONVENIO" & Chr(13)
   sql = sql & "                              LEFT  JOIN COMPLEMENTAR E WITH(NOLOCK)            ON A.REGISTRO = E.REGISTRO" & Chr(13)
   sql = sql & "                              LEFT  JOIN DADOSGUIA F WITH(NOLOCK)               ON A.REGISTRO = F.REGISTRO" & Chr(13)
   sql = sql & "                                                                 AND F.TIPOREGISTRO=3" & Chr(13)
   sql = sql & "                              LEFT  JOIN SUS_PROCEDIMENTO G WITH(NOLOCK)        ON A.PROCSUS_CIHA =  G.CODIGOSUS_PROCEDIMENTO" & Chr(13)
   sql = sql & "   Where 1 = 1" & Chr(13)
   sql = sql & "   AND    D.TIPOCONVENIO <> 3" & Chr(13)
   sql = sql & "   AND    ISNULL(A.CANCELADO,0) = 0" & Chr(13)
   sql = sql & "   AND    A.REGISTRO>0" & Chr(13)
   sql = sql & "   AND EXISTS (SELECT CON.PROCEDIMENTO" & Chr(13)
   sql = sql & "FROM SUS_PROCEDIMENTO_REGISTRO CON" & Chr(13)
   sql = sql & "             Where CON.Registro = 1" & Chr(13)
   sql = sql & "             AND   LEFT(REPLICATE('0',10-LEN(LEFT(G.CODIGOSUS_PROCEDIMENTO,10))) + LTRIM(G.CODIGOSUS_PROCEDIMENTO),2) IN ('02', '03', '04', '05')" & Chr(13)
   sql = sql & "             AND   CON.PROCEDIMENTO = G.CODIGOSUS_PROCEDIMENTO )" & Chr(13)
   sql = sql & "GROUP BY G.CODIGOSUS_PROCEDIMENTO, A.DATAINTERNACAO" & Chr(13)
   Banco.Execute sql
   
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes072011()
   On Error GoTo Erro
   
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Consulta Exames Anteriores', 'mnuMed_Lau_CEA', ' ', 'mnuMed_Lau_CEA',"
   sql = sql & "        1, 1, '0606060700', 'mnuMed_Lau_CEA', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = " ALTER TABLE USUARIO ADD PERMITE_CONFERIR_LAUDO INT   "
   Banco.Execute sql
   
   sql = " ALTER TABLE PRODUTOCONVENIO ADD MEDICBRAS VARCHAR(20)"
   Banco.Execute sql
   
   sql = " ALTER TABLE PRODUTOCONVENIO ADD CODLABOBRAS VARCHAR(20)"
   Banco.Execute sql

   sql = " ALTER TABLE PRODUTOCONVENIO ADD CODAPREBRAS VARCHAR(100)"
   Banco.Execute sql

   sql = " ALTER TABLE PRODUTOCONVENIO ADD QTDEMBALAGEM MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE SAVELOG..PRODUTOCONVENIO ADD VALOR_ANTIGO MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE SAVELOG..PRODUTOCONVENIO ADD VALORRECEBIDO_ANTIGO MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE SAVELOG..PRODUTOCONVENIO ADD DATAOPERACAO DATETIME"
   Banco.Execute sql
   
   sql = " ALTER TABLE SAVELOG..PRODUTOCONVENIO ADD TIPOOPERACAO INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE PRODUTOCONVENIO ADD VALORBRASINDICE MONEY "
   Banco.Execute sql
   
   sql = " ALTER TABLE PRODUTOBRASINDICE ADD TIPOPRODUTO INT"
   Banco.Execute sql
      
   If Layout = 21 Then
      sql = " UPDATE LEITOUNIDADE SET QUANTIDADELEITO = 12, LEITOATIVOS = 12 WHERE LEITOUNIDADE = 9"
      Banco.Execute sql
   End If
   
   sql = "CREATE TABLE LAU_COLETA( "
   sql = sql & " COLETA INT IDENTITY,"
   sql = sql & " TIPO VARCHAR(20),"
   sql = sql & " DATA DATETIME,"
   sql = sql & " ATUALIZACAO VARCHAR(50),"
   sql = sql & " CONSTRAINT [PK_LAU_COLETA] PRIMARY KEY NONCLUSTERED"
   sql = sql & " ("
   sql = sql & "    coleta Asc"
   sql = sql & " ))"
   Banco.Execute sql

   sql = "ALTER TABLE LAU_MOVIM_AMB"
   sql = sql & " ADD COLETA INT CONSTRAINT [FK_LAU_MOVIM_AMB] FOREIGN KEY(COLETA)"
   sql = sql & " References LAU_COLETA(coleta)"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_INT"
   sql = sql & " ADD COLETA INT CONSTRAINT [FK_LAU_MOVIM_INT] FOREIGN KEY(COLETA)"
   sql = sql & " References LAU_COLETA(coleta)"
   Banco.Execute sql

   sql = "ALTER TABLE LAU_MOVIM_EXT"
   sql = sql & " ADD COLETA INT CONSTRAINT [FK_LAU_MOVIM_EXT] FOREIGN KEY(COLETA)"
   sql = sql & " References LAU_COLETA(coleta)"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_DET_SERV"
   sql = sql & " ADD COLETA INT CONSTRAINT [FK_LAU_MOVIM_DET_SERV] FOREIGN KEY(COLETA)"
   sql = sql & " References LAU_COLETA(coleta)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AGENDAMENTOCONSULTA"
   sql = sql & " ADD DATAATUALIZACAO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE tmpATENDIMENTOANUAL"
   sql = sql & " ADD FAIXAETARIA VARCHAR(50)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[CUSTO_CC_EVENTO]("
   sql = sql & "    [CENTROCUSTO] [smallint] NOT NULL,"
   sql = sql & "    [ITEMCUSTO] [int] NOT NULL,"
   sql = sql & "    [EVENTO] [int] NOT NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_CUSTO_] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [CENTROCUSTO] ASC,"
   sql = sql & "    [ITEMCUSTO] ASC,"
   sql = sql & "    [EVENTO] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = " ALTER TABLE [dbo].[CUSTO_CC_EVENTO]  WITH CHECK ADD  CONSTRAINT [FK_CUSTO_CC_EVENTO_CentroCusto] FOREIGN KEY([CENTROCUSTO])"
   sql = sql & " References [dbo].[CentroCusto]([CentroCusto])"
   Banco.Execute sql
   
   sql = " ALTER TABLE [dbo].[CUSTO_CC_EVENTO]  WITH CHECK ADD  CONSTRAINT [FK_CUSTO_CC_EVENTO_CUSTO_ITEMCUSTO] FOREIGN KEY([ITEMCUSTO])"
   sql = sql & " References [dbo].[CUSTO_ITEMCUSTO]([ItemCusto])"
   Banco.Execute sql

   sql = " ALTER TABLE [dbo].[CUSTO_CC_EVENTO]  WITH CHECK ADD  CONSTRAINT [FK_CUSTO_CC_EVENTO_TER_EVENTO] FOREIGN KEY([EVENTO])"
   sql = sql & " References [dbo].[TER_EVENTO]([EVENTO])"
   Banco.Execute sql
   
   sql = ""
   sql = sql & "  INSERT INTO MENU("
   sql = sql & "      MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "      NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & "  SELECT MAX(MENU)+1,'Importação de Pedidos', 'mnuCom_Imp', ' ', 'mnuCom_Imp',"
   sql = sql & "         1, 2, '0513000000', 'mnuCom_Imp', 1"
   sql = sql & "  From Menu"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_COLETA"
   sql = sql & " ADD RESPONSAVEL SMALLINT"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_COLETA  WITH CHECK ADD  CONSTRAINT [FK_LAU_COLETA_MEDICO] FOREIGN KEY(RESPONSAVEL)"
   sql = sql & " References MEDICOS(Medico)"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_EXT] ON [dbo].[LAU_MOVIM_EXT]"
   sql = sql & " ("
   sql = sql & "  coleta Asc"
   sql = sql & " )"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_INT] ON [dbo].[LAU_MOVIM_INT]"
   sql = sql & " ("
   sql = sql & " coleta Asc"
   sql = sql & " )"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_AMB] ON [dbo].[LAU_MOVIM_AMB]"
   sql = sql & " ("
   sql = sql & " coleta Asc"
   sql = sql & " )"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_DET_SERV] ON [dbo].[LAU_MOVIM_DET_SERV]"
   sql = sql & " ("
   sql = sql & " coleta Asc"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = " ALTER TABLE SUS_PROCEDIMENTO ADD EXPORTA_COMPETENCIA INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " UPDATE SUS_PROCEDIMENTO SET EXPORTA_COMPETENCIA = 1 "
   sql = sql & " WHERE CODIGOSUS_PROCEDIMENTO IN (802010024, 802010032, 802010040, 309010071, 309010080, 309010098,"
   sql = sql & "       309010047, 309010055, 309010063, 0503040045)"
   Banco.Execute sql
   
   sql = " UPDATE SUS_PROCEDIMENTO SET EXPORTA_COMPETENCIA = 1 "
   sql = sql & " Where CODIGOSUS_PROCEDIMENTO >= 802010075 And CODIGOSUS_PROCEDIMENTO <= 802010164"
   Banco.Execute sql
   
   sql = " ALTER TABLE COTACAO1 ADD NUMEROPDC VARCHAR(12)"
   Banco.Execute sql

   sql = " ALTER TABLE COTACAO1 ADD DATAINICIAL DATETIME"
   Banco.Execute sql

   sql = " ALTER TABLE PEDIDOCOMPRA1 ADD DIVERGENTE INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CENTROCUSTO ADD CONT_CUSTO_SALARIO VARCHAR(20)"
   Banco.Execute sql
   
   sql = " ALTER TABLE CENTROCUSTO ADD CONT_CUSTO_FERIAS VARCHAR(20)"
   Banco.Execute sql
   
   sql = " ALTER TABLE CENTROCUSTO ADD CONT_CUSTO_FGTS VARCHAR(20)"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes082011()
   On Error GoTo Erro

   sql = ""
   sql = sql & " CREATE VIEW V_RECUPERA_CIHA_CONSOLIDADO_2 " & Chr(13)
   sql = sql & " AS" & Chr(13)
   sql = sql & "      SELECT 5 AS TIPOREG, 'AMB' AS TIPO, A.DATAINTERNACAO," & Chr(13)
   sql = sql & "           MAX(E.CODIGOSUS) AS PROCSUS, 1 AS QUANTIDADE, C.REGISTRO, C.SEQUENCIA" & Chr(13)
   sql = sql & "      FROM   AMBULATORIAL A WITH(NOLOCK) INNER JOIN FICHAS       B WITH(NOLOCK) ON A.FICHA    = B.FICHA" & Chr(13)
   sql = sql & "                                INNER JOIN MOVIM_AMB    C WITH(NOLOCK) ON A.REGISTRO = C.REGISTRO" & Chr(13)
   sql = sql & "                                                             AND C.TIPOLANCAMENTO IN (2,3)" & Chr(13)
   sql = sql & "                                INNER JOIN CONVENIOS    D WITH(NOLOCK)               ON A.CONVENIO = D.CONVENIO" & Chr(13)
   sql = sql & "                                LEFT  JOIN SUS_TABELA_CIH E ON C.PROCEDIMENTO = E.CODIGO" & Chr(13)
   sql = sql & "                                                    AND D.TABELA       = E.TABELA" & Chr(13)
   sql = sql & "  WHERE 1 = 1" & Chr(13)
   sql = sql & "  AND    D.TIPOCONVENIO <> 3" & Chr(13)
   sql = sql & "  AND    ISNULL(A.CANCELADO,0) = 0" & Chr(13)
   sql = sql & "  AND    A.REGISTRO>0" & Chr(13)
   sql = sql & "  AND EXISTS (SELECT CON.PROCEDIMENTO" & Chr(13)
   sql = sql & "              FROM SUS_PROCEDIMENTO_REGISTRO CON" & Chr(13)
   sql = sql & "              Where CON.Registro = 1 " & Chr(13)
   sql = sql & "              AND   LEFT(REPLICATE('0',10-LEN(LEFT(E.CODIGOSUS,10))) + LTRIM(E.CODIGOSUS),2) IN ('02', '03', '04', '05')" & Chr(13)
   sql = sql & "              AND   CON.PROCEDIMENTO = E.CODIGOSUS)" & Chr(13)
   sql = sql & "  GROUP BY A.DATAINTERNACAO, C.REGISTRO, C.SEQUENCIA" & Chr(13)
   sql = sql & "  Union All" & Chr(13)
   sql = sql & "    SELECT 5 AS TIPOREG, 'AMB' AS TIPO, A.DATAINTERNACAO," & Chr(13)
   sql = sql & "          MAX(E.CODIGOSUS) AS PROCSUS,  1 AS QUANTIDADE, C.REGISTRO, C.SEQUENCIA" & Chr(13)
   sql = sql & "    FROM   EXTERNO      A WITH(NOLOCK) INNER JOIN FICHAS    B WITH(NOLOCK) ON A.FICHA = B.FICHA" & Chr(13)
   sql = sql & "                               INNER JOIN MOVIM_EXT C WITH(NOLOCK) ON A.REGISTRO = C.REGISTRO" & Chr(13)
   sql = sql & "                                                             AND C.TIPOLANCAMENTO IN (2,3)" & Chr(13)
   sql = sql & "                               INNER JOIN CONVENIOS D WITH(NOLOCK)               ON A.CONVENIO = D.CONVENIO" & Chr(13)
   sql = sql & "                               LEFT  JOIN SUS_TABELA_CIH E ON C.PROCEDIMENTO = E.CODIGO" & Chr(13)
   sql = sql & "                                                   AND D.TABELA       = E.TABELA" & Chr(13)
   sql = sql & "    Where 1 = 1" & Chr(13)
   sql = sql & "    AND    D.TIPOCONVENIO <> 3" & Chr(13)
   sql = sql & "    AND    ISNULL(A.CANCELADO,0) = 0" & Chr(13)
   sql = sql & "    AND    A.REGISTRO>0" & Chr(13)
   sql = sql & "    AND EXISTS (SELECT CON.PROCEDIMENTO" & Chr(13)
   sql = sql & "              FROM SUS_PROCEDIMENTO_REGISTRO CON" & Chr(13)
   sql = sql & "              Where CON.Registro = 1 " & Chr(13)
   sql = sql & "              AND   LEFT(REPLICATE('0',10-LEN(LEFT(E.CODIGOSUS,10))) + LTRIM(E.CODIGOSUS),2) IN ('02', '03', '04', '05')" & Chr(13)
   sql = sql & "              AND   CON.PROCEDIMENTO = E.CODIGOSUS)" & Chr(13)
   sql = sql & "    GROUP BY A.DATAINTERNACAO, C.REGISTRO, C.SEQUENCIA" & Chr(13)
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD CONSULTAPREANESTESICA TINYINT"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_AMBULATORIAL_CONSULTAPREANESTESICA] ON [Ambulatorial] " & Chr(13)
   sql = sql & " ( " & Chr(13)
   sql = sql & "    [CONSULTAPREANESTESICA] ASC " & Chr(13)
   sql = sql & " )"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD TISS_RESUMO_EXPORTA_HONORARIO INT "
   Banco.Execute sql
   
   'MARILIA = UNIMED
   If Layout = 21 Then
      sql = " UPDATE CONVENIOS SET TISS_RESUMO_EXPORTA_HONORARIO = 1 WHERE CONVENIO = 1 "
      Banco.Execute sql
   End If
   
   'PRO-SAUDE = UNIMED E CASSI
   If Layout = 13 Then
      sql = " UPDATE CONVENIOS SET TISS_RESUMO_EXPORTA_HONORARIO = 1 WHERE CONVENIO IN ( 21 , 5) "
      Banco.Execute sql
   End If
      
   'CAFELANDIA = SAO LUCAS
   If Layout = 36 Then
      sql = " UPDATE CONVENIOS SET TISS_RESUMO_EXPORTA_HONORARIO = 1 WHERE CONVENIO = 15 "
      Banco.Execute sql
   End If
      
   sql = "ALTER TABLE EXTERNO ADD APAC_QUIMIO_METASTASE_STR VARCHAR(100)"
   Banco.Execute sql
      
   sql = ""
   sql = sql & " INSERT INTO MENU(MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "       NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Integração Folha de Pagamento', 'mnuLan_IFP', ' ', 'mnuLan_IFP',"
   sql = sql & "        1, 4, '0212000000', 'mnuLan_IFP', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql

   sql = ""
   sql = sql & " CREATE TABLE [dbo].[CUSTO_INTEGRACAO_FOLHA]("
   sql = sql & "    [ANO] [int] NOT NULL,"
   sql = sql & "    [MES] [int] NOT NULL,"
   sql = sql & "    [ITEMCUSTO] [int] NOT NULL,"
   sql = sql & "    [CENTROCUSTO] [int] NOT NULL,"
   sql = sql & "    [VALOR] [money] NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_CUSTO_INTEGRACAO_FOLHA] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [ANO] ASC,"
   sql = sql & "    [MES] ASC,"
   sql = sql & "    [ITEMCUSTO] ASC,"
   sql = sql & "    [CentroCusto] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_INT " & Chr(13)
   sql = sql & "ADD DATACONFERENCIA DATETIME " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_AMB " & Chr(13)
   sql = sql & "ADD DATACONFERENCIA DATETIME " & Chr(13)
   Banco.Execute sql
   If Layout = 21 Then
      sql = " ALTER TABLE FICHAS ALTER COLUMN NOMEPAI VARCHAR(250)"
      Banco.Execute sql
      
      sql = " ALTER TABLE FICHAS ALTER COLUMN NOMEMAE VARCHAR(250)"
      Banco.Execute sql
   
   
   End If
   
   
   sql = "ALTER TABLE LAU_MOVIM_DET_SERV " & Chr(13)
   sql = sql & "ADD DATACONFERENCIA DATETIME " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_EXT " & Chr(13)
   sql = sql & "ADD DATACONFERENCIA DATETIME " & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & "CREATE TABLE [dbo].[TMP_CIHA]("
   sql = sql & "   [TIPOEXPORTACAO] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "   [TIPOREGISTRO] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "   [REGISTRO] [int] NULL,"
   sql = sql & "   [NOME] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "   [PROCEDIMENTO] [int] NULL,"
   sql = sql & "   [QUANTIDADE] [int] NULL,"
   sql = sql & "   [IP] [varchar](100) COLLATE Latin1_General_CI_AS NULL"
   sql = sql & ") ON [PRIMARY]"
   Banco.Execute sql
   
   sql = " ALTER TABLE RECIBO_FINANCEIRO ADD VALORDESCONTO MONEY"
   Banco.Execute sql
   
   sql = "CREATE TABLE LAU_REGLAB( " & Chr(13)
   sql = sql & " REGLAB INT IDENTITY, " & Chr(13)
   sql = sql & " TIPO VARCHAR(20), " & Chr(13)
   sql = sql & " DATA DATETIME, " & Chr(13)
   sql = sql & " ATUALIZACAO VARCHAR(50), " & Chr(13)
   sql = sql & " CONSTRAINT [PK_LAU_REGLAB] PRIMARY KEY NONCLUSTERED " & Chr(13)
   sql = sql & " ( " & Chr(13)
   sql = sql & "    REGLAB Asc " & Chr(13)
   sql = sql & " ))"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_AMB " & Chr(13)
   sql = sql & " ADD REGLAB INT CONSTRAINT [FK_LAU_MOVIM_AMB_REGLAB] FOREIGN KEY(REGLAB) " & Chr(13)
   sql = sql & " References LAU_REGLAB(REGLAB) " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_INT " & Chr(13)
   sql = sql & " ADD REGLAB INT CONSTRAINT [FK_LAU_MOVIM_INT_REGLAB] FOREIGN KEY(REGLAB) " & Chr(13)
   sql = sql & " References LAU_REGLAB(REGLAB)"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_EXT " & Chr(13)
   sql = sql & " ADD REGLAB INT CONSTRAINT [FK_LAU_MOVIM_EXT_REGLAB] FOREIGN KEY(REGLAB) " & Chr(13)
   sql = sql & " References LAU_REGLAB(REGLAB)"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_DET_SERV " & Chr(13)
   sql = sql & " ADD REGLAB INT CONSTRAINT [FK_LAU_MOVIM_DET_SERV_REGLAB] FOREIGN KEY(REGLAB) " & Chr(13)
   sql = sql & " References LAU_REGLAB(REGLAB)"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_EXT_REGLAB] ON [dbo].[LAU_MOVIM_EXT]"
   sql = sql & " ("
   sql = sql & "  REGLAB Asc"
   sql = sql & " )"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_INT_REGLAB] ON [dbo].[LAU_MOVIM_INT]"
   sql = sql & " ("
   sql = sql & " REGLAB Asc"
   sql = sql & " )"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_AMB_REGLAB] ON [dbo].[LAU_MOVIM_AMB]"
   sql = sql & " ("
   sql = sql & " REGLAB Asc"
   sql = sql & " )"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_DET_SERV_REGLAB] ON [dbo].[LAU_MOVIM_DET_SERV]"
   sql = sql & " ("
   sql = sql & " REGLAB Asc"
   sql = sql & " )"
   Banco.Execute sql
      
   If Layout = 1 Then
      sql = ""
      sql = sql & " INSERT INTO MENU("
      sql = sql & "    MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
      sql = sql & "    NOMESUBAUX,NIVELVISIBILIDADE)"
      sql = sql & " SELECT MAX(MENU)+1,'Consulta Arquivo de Prontuários', 'mnuSam_Arq_CAr', ' ', 'mnuSam_Arq_CAr',"
      sql = sql & "    1, 1, '0301040000', 'mnuSam_Arq_CAr', 1"
      sql = sql & " From Menu"
      Banco.Execute sql
   End If
   
   sql = " ALTER TABLE PARAMETRO ADD SAME_CAMINHO_PRONT_INTERNO  VARCHAR(250)"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD SAME_CAMINHO_PRONT_EXTERNO  VARCHAR(250)"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD SAME_CAMINHO_PRONT_AMBULATORIAL  VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[LAU_COLETA] DROP CONSTRAINT [FK_LAU_COLETA_MEDICO]"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_COLETA " & Chr(13)
   sql = sql & "ALTER COLUMN RESPONSAVEL INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[LAU_COLETA]  WITH CHECK ADD  CONSTRAINT [FK_LAU_COLETA_ENFERMEIRA] FOREIGN KEY([RESPONSAVEL]) " & Chr(13)
   sql = sql & "REFERENCES [dbo].[ENFERMEIRA] ([ENFERMEIRA])"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO ADD RECLAMACAO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE EXTERNO ADD RECLAMACAO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD RECLAMACAO INT"
   Banco.Execute sql
   
   sql = "CREATE TABLE CONTROLEAMBULANCIA("
   sql = sql & " CONTROLEAMBULANCIA INT IDENTITY,"
   sql = sql & " SOLICITANTE INT,"
   sql = sql & " DATASOLICITACAO DATETIME,"
   sql = sql & " ATENDENTE VARCHAR(255),"
   sql = sql & " DATACHEGADA DATETIME,"
   sql = sql & " DATASAIDA DATETIME,"
   sql = sql & " PLACA VARCHAR(20),"
   sql = sql & " DEONDEVEIO VARCHAR(255),"
   sql = sql & " PARAONDEVAI VARCHAR(255),"
   sql = sql & " DATACHEGADAFAMILIA DATETIME,"
   sql = sql & " DATASAIDAFAMILIA DATETIME,"
   sql = sql & " TIPOREGISTRO TINYINT,"
   sql = sql & " REGISTRO INT,"
   sql = sql & " OBSERVACAO VARCHAR(500),"
   sql = sql & " ATUALIZACAO VARCHAR(200),"
   sql = sql & " CONSTRAINT [PK_CONTROLEAMBULANCIA] PRIMARY KEY NONCLUSTERED"
   sql = sql & " ("
   sql = sql & "    [CONTROLEAMBULANCIA] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = " ALTER TABLE CONTROLEAMBULANCIA ADD CONSTRAINT FK_CONTROLEAMBULANCIA_USUARIO FOREIGN KEY(SOLICITANTE)"
   sql = sql & " REFERENCES USUARIO(USUARIO)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "    MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "    NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Controle Ambulância', 'mnuRec_CAm', ' ', 'mnuRec_CAm',"
   sql = sql & "    1, 1, '0220000000', 'mnuRec_CAm', 1"
   sql = sql & " From Menu"
   Banco.Execute sql
   
   sql = "alter table custo_carro"
   sql = sql & " add MARCA VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CUSTO_CARRO"
   sql = sql & " ADD MODELO VARCHAR(250)"
   Banco.Execute sql
   
   sql = "alter TABLE CUSTO_KM"
   sql = sql & " ADD MOTORISTA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CUSTO_KM"
   sql = sql & " ADD DESTINO VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO "
   sql = sql & " ALTER COLUMN APAC_QUIMIO_CIDTRATAMENTO1 VARCHAR(250)"
   Banco.Execute sql
    
   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ALTER COLUMN APAC_QUIMIO_CIDTRATAMENTO2 VARCHAR(250)"
   Banco.Execute sql
    
   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ALTER COLUMN APAC_QUIMIO_CIDTRATAMENTO3 VARCHAR(250)"
   Banco.Execute sql
    
   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD APAC_QUIMIO_DESCRICAOTRATAMENTO1 VARCHAR(250)"
    Banco.Execute sql
    
   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD APAC_QUIMIO_DESCRICAOTRATAMENTO2 VARCHAR(250)"
   Banco.Execute sql
    
   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD APAC_QUIMIO_DESCRICAOTRATAMENTO3 VARCHAR(250)"
   Banco.Execute sql
    
   sql = "ALTER TABLE TMPPRESTACAOMORADORES"
   sql = sql & " ADD DATAATUALIZACAO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPPRESTACAOMORADORES"
   sql = sql & " ADD USUARIO VARCHAR(100)"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes092011()
   On Error GoTo Erro
   
   If Layout = 30 Then
      sql = ""
      sql = sql & " UPDATE MENU SET NOMECAPTION = 'Fechamento Parcial de Faturamento'"
      sql = sql & " WHERE NOMESUBNOVO = 'mnuFat_Rch'"
      Banco.Execute sql
   End If

   sql = " ALTER TABLE sus_CBO ADD CBOCONVENIO INT DEFAULT(0)"
   Banco.Execute sql

   sql = " ALTER TABLE SAME_ESTATISTICA ADD DESCRICAOUNIDADE VARCHAR(100)"
   Banco.Execute sql

   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Gerencial Anual Pronto Atendimento', 'mnuSam_Est_Pro', ' ', 'mnuSam_Est_Pro',"
   sql = sql & "        1, 1, '0302210000', 'mnuSam_Est_Pro', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql

   sql = ""
   sql = sql & " CREATE TABLE [dbo].[CENSO_MENSAL]("
   sql = sql & "    [ANO] [int] NOT NULL,"
   sql = sql & "    [MES] [int] NOT NULL,"
   sql = sql & "    [UNIDADE] [int] NOT NULL,"
   sql = sql & "    [DESCRICAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [LEITOS] [int] NULL,"
   sql = sql & "    [PD_SUS] [int] NULL,"
   sql = sql & "    [PD_OUTROS] [int] NULL,"
   sql = sql & "    [PD_PARTICULAR] [int] NULL,"
   sql = sql & "    [PD_UNIMED] [int] NULL,"
   sql = sql & "    [INT_SUS] [int] NULL,"
   sql = sql & "    [INT_OUTROS] [int] NULL,"
   sql = sql & "    [INT_PARTICULAR] [int] NULL,"
   sql = sql & "    [INT_UNIMED] [int] NULL,"
   sql = sql & "    [PD_TOTAL] [int] NULL,"
   sql = sql & "    [INT_TOTAL] [int] NULL,"
   sql = sql & "  CONSTRAINT [PK_CENSO_MENSAL] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [ANO] ASC,"
   sql = sql & "    [MES] ASC,"
   sql = sql & "    [Unidade] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = ""
   sql = sql & " CREATE TABLE [dbo].[CUSTO_FECHAMENTO_MENSAL]("
   sql = sql & "    [ANO] [int] NOT NULL,"
   sql = sql & "    [MES] [int] NOT NULL,"
   sql = sql & "    [CENTROCUSTO] [int] NOT NULL,"
   sql = sql & "    [VALOR] [money] NULL,"
   sql = sql & "    [DATAFECHAMENTO] [datetime] NULL,"
   sql = sql & "    [USUARIO] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_CUSTO_FECHAMENTO_MENSAL] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [ANO] ASC,"
   sql = sql & "    [MES] ASC,"
   sql = sql & "    [CentroCusto] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = ""
   sql = sql & " CREATE TABLE [dbo].[CUSTO_CC_FUNCIONARIO]("
   sql = sql & "    [ANO] [int] NOT NULL,"
   sql = sql & "    [MES] [int] NOT NULL,"
   sql = sql & "    [CENTROCUSTO] [smallint] NOT NULL,"
   sql = sql & "    [FUNCIONARIO] [int] NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_CUSTO_CC_FUNCIONARIO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [ANO] ASC,"
   sql = sql & "    [MES] ASC,"
   sql = sql & "    [CentroCusto] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = " ALTER TABLE AMB92 ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMB99 ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE BRASIL ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE BRASIL2 ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE SUSINTERNOS ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE SUSEXTERNOS ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CIEFAS2000 ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE IAMSPEINTERNOS ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE IAMSPEEXTERNOS ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE SULAMERICA ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE BRADESCO ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CAIXAECONOMICA ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERCLINICA ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CABESP ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMB90 ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CBHPM ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE SUS_PROCEDIMENTO ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CBHPM3 ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CBHPM4 ADD TUSS_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CBHPM5 ADD TUSS_CONVENIO INT"
   Banco.Execute sql

   sql = " ALTER TABLE IAMSPEINTERNOS ADD PERMITEQUANTIDADEMAIOR INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_TISS_OUTRASDESPESAS " & Chr(13)
   sql = sql & " ADD NOMEPACIENTE VARCHAR(200)"
   Banco.Execute sql
   
   sql = "alter table ambulatorial " & Chr(13)
   sql = sql & " add QUANTIDADESEMANAGESTACAO varchar(50)"
   Banco.Execute sql
   
   sql = "alter table ambulatorial " & Chr(13)
   sql = sql & " add PARTOINDICACAO varchar(255)"
   Banco.Execute sql
   
   sql = "alter table ambulatorial " & Chr(13)
   sql = sql & " add PARTONUMEROABORTOS int"
   Banco.Execute sql
   
   sql = "alter table ambulatorial " & Chr(13)
   sql = sql & " add PARTONUMEROCESARIANAS int"
   Banco.Execute sql
   
   sql = "alter table ambulatorial " & Chr(13)
   sql = sql & " add PARTONUMEROPARTOS int"
   Banco.Execute sql
   
   sql = "alter table ambulatorial " & Chr(13)
   sql = sql & " add PARTONUMEROGESTACOES int"
   Banco.Execute sql
   
   sql = "alter table interno " & Chr(13)
   sql = sql & " add QUANTIDADESEMANAGESTACAO varchar(50)"
   Banco.Execute sql
   
   sql = "alter table interno " & Chr(13)
   sql = sql & " add PARTOINDICACAO varchar(255)"
   Banco.Execute sql
   
   sql = "alter table interno " & Chr(13)
   sql = sql & " add PARTONUMEROABORTOS int"
   Banco.Execute sql
   
   sql = "alter table interno " & Chr(13)
   sql = sql & " add PARTONUMEROCESARIANAS int"
   Banco.Execute sql
   
   sql = "alter table interno " & Chr(13)
   sql = sql & " add PARTONUMEROPARTOS int"
   Banco.Execute sql
   
   sql = "alter table interno " & Chr(13)
   sql = sql & " add PARTONUMEROGESTACOES int"
   Banco.Execute sql
   
   sql = "alter table externo " & Chr(13)
   sql = sql & " add QUANTIDADESEMANAGESTACAO varchar(50)"
   Banco.Execute sql
   
   sql = "alter table externo " & Chr(13)
   sql = sql & " add PARTOINDICACAO varchar(255)"
   Banco.Execute sql
   
   sql = "alter table externo " & Chr(13)
   sql = sql & " add PARTONUMEROABORTOS int"
   Banco.Execute sql
   
   sql = "alter table externo " & Chr(13)
   sql = sql & " add PARTONUMEROCESARIANAS int"
   Banco.Execute sql
   
   sql = "alter table externo " & Chr(13)
   sql = sql & " add PARTONUMEROPARTOS int"
   Banco.Execute sql
   
   sql = "alter table externo " & Chr(13)
   sql = sql & " add PARTONUMEROGESTACOES int"
   Banco.Execute sql
   
   sql = "alter table cirurgia " & Chr(13)
   sql = sql & " add apa int"
   Banco.Execute sql
   
   sql = "alter table cirurgia " & Chr(13)
   sql = sql & " add agendado int"
   Banco.Execute sql
   
   sql = "alter table cirurgia " & Chr(13)
   sql = sql & " add indicao int"
   Banco.Execute sql

   If Layout = 39 Or Layout = 42 Then
      sql = ""
      sql = sql & " UPDATE MENU SET NOMECAPTION = 'Fechamento Parcial de Faturamento'"
      sql = sql & " WHERE NOMESUBNOVO = 'mnuFat_Rch'"
      Banco.Execute sql
   End If
   
   sql = "alter table TMPIMPRESSAOCARTEIRINHA"
   sql = sql & " add LINHA5 VARCHAR(255)"
   Banco.Execute sql
   
   sql = "alter table TMPIMPRESSAOCARTEIRINHA"
   sql = sql & " add LINHA6 VARCHAR(255)"
   Banco.Execute sql
   
   sql = "alter table TMPIMPRESSAOCARTEIRINHA"
   sql = sql & " add LINHA7 VARCHAR(255)"
   Banco.Execute sql
   
   sql = "alter table TMPIMPRESSAOCARTEIRINHA"
   sql = sql & " add LINHA8 VARCHAR(255)"
   Banco.Execute sql
   
   sql = "alter table TMPIMPRESSAOCARTEIRINHA"
   sql = sql & " add LINHA9 VARCHAR(255)"
   Banco.Execute sql
   
   sql = "alter table TMPIMPRESSAOCARTEIRINHA"
   sql = sql & " add LINHA10 VARCHAR(255)"
   Banco.Execute sql
   
   sql = "alter table TMPIMPRESSAOCARTEIRINHA"
   sql = sql & " add LINHA11 VARCHAR(255)"
   Banco.Execute sql
   
   sql = "alter table TMPIMPRESSAOCARTEIRINHA"
   sql = sql & " add LINHA12 VARCHAR(255)"
   Banco.Execute sql
   
   sql = "alter table tmpdevolucaocompra"
   sql = sql & " alter column lote char(20)"
   Banco.Execute sql
   
   sql = "alter table tmpdevolucaocompra_log"
   sql = sql & " alter column lote char(20)"
   Banco.Execute sql
      
   sql = " ALTER TABLE SUS_PROCEDIMENTO ADD GENERICO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CIRURGIA"
   sql = sql & " DROP COLUMN APG"
   Banco.Execute sql

   sql = "ALTER TABLE CIRURGIA"
   sql = sql & " DROP COLUMN GESTACAOALTORISCO"
   Banco.Execute sql

   sql = "ALTER TABLE CIRURGIA"
   sql = sql & " DROP COLUMN VDRL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPRELDADOSCIRURGIA"
   sql = sql & " DROP COLUMN APG"
   Banco.Execute sql

   sql = "ALTER TABLE TMPRELDADOSCIRURGIA"
   sql = sql & " DROP COLUMN GESTACAOALTORISCO "
   Banco.Execute sql

   sql = "ALTER TABLE TMPRELDADOSCIRURGIA"
   sql = sql & " DROP COLUMN VDRL "
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_INT_PROCEDIMENTO] ON [dbo].[LAU_MOVIM_INT] "
   sql = sql & " ("
   sql = sql & "    [Procedimento] Asc"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_AMB_PROCEDIMENTO] ON [dbo].[LAU_MOVIM_AMB] "
   sql = sql & " ("
   sql = sql & "    [Procedimento] Asc"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_EXT_PROCEDIMENTO] ON [dbo].[LAU_MOVIM_EXT] "
   sql = sql & " ("
   sql = sql & "    [Procedimento] Asc"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = " ALTER TABLE USUARIO ADD PERMITE_CONFERIR_LAUDO INT   "
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD REGLABLIBERADO BIT"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD REGLABLIBERADO BIT"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD REGLABLIBERADO BIT"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes102011()
   On Error GoTo Erro

   sql = " ALTER TABLE MOVIM_PACIENTE_AMB ADD HORASAIDA DATETIME"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO ADD ACOMPANHANTEASSISTEPARTO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CENTROCUSTO ADD SALAOBSERVACAO INT"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_INT_PROCEDIMENTO] ON [dbo].[LAU_MOVIM_INT] "
   sql = sql & " ("
   sql = sql & "    [Procedimento] Asc"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_AMB_PROCEDIMENTO] ON [dbo].[LAU_MOVIM_AMB] "
   sql = sql & " ("
   sql = sql & "    [Procedimento] Asc"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_LAU_MOVIM_EXT_PROCEDIMENTO] ON [dbo].[LAU_MOVIM_EXT] "
   sql = sql & " ("
   sql = sql & "    [Procedimento] Asc"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = " ALTER TABLE USUARIO ADD PERMITE_CONFERIR_LAUDO INT   "
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD REGLABLIBERADO BIT"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD REGLABLIBERADO BIT"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD REGLABLIBERADO BIT"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIMENTOAVULSO"
   sql = sql & " ADD FABRICANTE VARCHAR(255)"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIMENTOPRODUTODIVERSOS"
   sql = sql & " ADD FABRICANTE VARCHAR(255)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[SUS_TABELA_CIH_INSUMO]("
   sql = sql & "    [CODIGOSUS] [int] NOT NULL,"
   sql = sql & "    [CONVENIO] [int] NOT NULL,"
   sql = sql & "    [INSUMO] [int] NOT NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_SUS_TABELA_CIH_INSUMO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [CODIGOSUS] ASC,"
   sql = sql & "    [CONVENIO] ASC,"
   sql = sql & "    [Insumo] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE VIEW V_RECUPERA_CIHA_INSUMO " & Chr(13)
   sql = sql & " AS " & Chr(13)
   sql = sql & "      SELECT 5 AS TIPOREG, 'AMB' AS TIPO, A.DATAINTERNACAO," & Chr(13)
   sql = sql & "           MAX(E.CODIGOSUS) AS PROCSUS, SUM(C.QUANTIDADE) AS QUANTIDADE, C.REGISTRO, C.SEQUENCIA" & Chr(13)
   sql = sql & "      FROM   AMBULATORIAL A WITH(NOLOCK) INNER JOIN FICHAS       B WITH(NOLOCK) ON A.FICHA    = B.FICHA" & Chr(13)
   sql = sql & "                                INNER JOIN MOVIM_AMB    C WITH(NOLOCK) ON A.REGISTRO = C.REGISTRO" & Chr(13)
   sql = sql & "                                                             AND C.TIPOLANCAMENTO IN (4)" & Chr(13)
   sql = sql & "                                INNER JOIN CONVENIOS    D WITH(NOLOCK)               ON A.CONVENIO = D.CONVENIO" & Chr(13)
   sql = sql & "                                INNER JOIN SUS_TABELA_CIH_INSUMO E ON C.PROCEDIMENTO = E.INSUMO" & Chr(13)
   sql = sql & "                                                          AND A.CONVENIO   = E.CONVENIO" & Chr(13)
   sql = sql & "     Where 1 = 1" & Chr(13)
   sql = sql & "     AND    D.TIPOCONVENIO <> 3" & Chr(13)
   sql = sql & "     AND    ISNULL(A.CANCELADO,0) = 0" & Chr(13)
   sql = sql & "     AND    A.REGISTRO>0" & Chr(13)
   sql = sql & "     AND EXISTS (SELECT CON.PROCEDIMENTO" & Chr(13)
   sql = sql & "                 FROM SUS_PROCEDIMENTO_REGISTRO CON" & Chr(13)
   sql = sql & "                 Where CON.Registro = 1" & Chr(13)
   sql = sql & "                 AND   CON.PROCEDIMENTO = E.CODIGOSUS)" & Chr(13)
   sql = sql & "    GROUP BY A.DATAINTERNACAO, C.REGISTRO, C.SEQUENCIA" & Chr(13)
   
   sql = sql & "     Union All" & Chr(13)
   
   sql = sql & "    SELECT 5 AS TIPOREG, 'AMB' AS TIPO, A.DATAINTERNACAO," & Chr(13)
   sql = sql & "          MAX(E.CODIGOSUS) AS PROCSUS,  SUM(C.QUANTIDADE) AS QUANTIDADE, C.REGISTRO, C.SEQUENCIA" & Chr(13)
   sql = sql & "    FROM   EXTERNO      A WITH(NOLOCK) INNER JOIN FICHAS    B WITH(NOLOCK) ON A.FICHA = B.FICHA" & Chr(13)
   sql = sql & "                               INNER JOIN MOVIM_EXT C WITH(NOLOCK) ON A.REGISTRO = C.REGISTRO" & Chr(13)
   sql = sql & "                                                             AND C.TIPOLANCAMENTO IN (4)" & Chr(13)
   sql = sql & "                               INNER JOIN CONVENIOS D WITH(NOLOCK)               ON A.CONVENIO = D.CONVENIO" & Chr(13)
   sql = sql & "                                INNER JOIN SUS_TABELA_CIH_INSUMO E ON C.PROCEDIMENTO = E.INSUMO" & Chr(13)
   sql = sql & "                                                          AND A.CONVENIO   = E.CONVENIO" & Chr(13)
   sql = sql & "    Where 1 = 1" & Chr(13)
   sql = sql & "    AND    D.TIPOCONVENIO <> 3" & Chr(13)
   sql = sql & "    AND    ISNULL(A.CANCELADO,0) = 0" & Chr(13)
   sql = sql & "    AND    A.REGISTRO>0" & Chr(13)
   sql = sql & "    AND EXISTS (SELECT CON.PROCEDIMENTO" & Chr(13)
   sql = sql & "              FROM SUS_PROCEDIMENTO_REGISTRO CON" & Chr(13)
   sql = sql & "              Where CON.Registro = 1 " & Chr(13)
   sql = sql & "              AND   CON.PROCEDIMENTO = E.CODIGOSUS)" & Chr(13)
   sql = sql & "    GROUP BY A.DATAINTERNACAO, C.REGISTRO, C.SEQUENCIA" & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_INT"
   sql = sql & " ADD DATAIMPRESSAO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_INT"
   sql = sql & " ADD USUARIOIMPRESSAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_AMB"
   sql = sql & " ADD DATAIMPRESSAO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_AMB"
   sql = sql & " ADD USUARIOIMPRESSAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_EXT"
   sql = sql & " ADD DATAIMPRESSAO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_EXT"
   sql = sql & " ADD USUARIOIMPRESSAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE USUARIO"
   sql = sql & " ADD USUARIOUNIMED INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE USUARIO ADD PREPARAKITMEDICAMENTO INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU(MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "       NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Comparativo Faturamento / Custo', 'mnuCRe_Fat', ' ', 'mnuCRe_Fat',"
   sql = sql & "    1, 4, '0405000000', 'mnuCRe_Fat', 1"
   sql = sql & " From Menu"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD APAC_QUIMIO_LOCALTUMOR VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD APAC_QUIMIO_CIDTUMOR VARCHAR(10)"
   Banco.Execute sql
   
   sql = " ALTER TABLE LAU_MOVIM_INT ADD USUARIOCONFERENCIA VARCHAR(30)"
   Banco.Execute sql

   sql = " ALTER TABLE LAU_MOVIM_AMB ADD USUARIOCONFERENCIA VARCHAR(30)"
   Banco.Execute sql

   sql = " ALTER TABLE LAU_MOVIM_EXT ADD USUARIOCONFERENCIA VARCHAR(30)"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD HABILITACAO_SUS INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_PRESCRICAOELETRONICA_REPETIR "
   sql = sql & "ADD ORDEM INT"
   Banco.Execute sql
   
   sql = "CREATE TABLE TMP_PRESCRICAOELETRONICA_REPETIR_COMPMANUAL( " & Chr(13)
   sql = sql & "REGISTRO INT, " & Chr(13)
   sql = sql & "PRESCRICAO INT, " & Chr(13)
   sql = sql & "PRODUTO INT, " & Chr(13)
   sql = sql & "SEQUENCIA INT, " & Chr(13)
   sql = sql & "QUANTIDADE CHAR(10), " & Chr(13)
   sql = sql & "SEQUENCIAITEMASSOCIADO INT, " & Chr(13)
   sql = sql & "OBSERVACAO VARCHAR(255), " & Chr(13)
   sql = sql & "IP VARCHAR(255)) "
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes112011()
   On Error GoTo Erro

   sql = " ALTER TABLE CONTROLEPERTENCE ADD CENTROCUSTO INT"
   Banco.Execute sql

   sql = " ALTER TABLE PRESCRICAOELETRONICAPERIODO_INT ADD PRESCRICAOIMPRESSA INT"
   Banco.Execute sql

   sql = " ALTER TABLE PRESCRICAOELETRONICAPERIODO_AMB ADD PRESCRICAOIMPRESSA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PRESCRICAOELETRONICAPERIODO_EXT ADD PRESCRICAOIMPRESSA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PRODUTOCONVENIO ADD PERCENTUALADICIONAL MONEY"
   Banco.Execute sql

   sql = " ALTER TABLE INTERNO ADD DECLARACAO_OBITO_RN VARCHAR(30)"
   Banco.Execute sql

   If Layout = 19 Then
      sql = " UPDATE CONVENIOS SET "
      sql = sql & " CODIGOMATERIALTISS = 82010021, CODIGOMEDICAMENTOTISS = 82020027"
      sql = sql & " Where Convenio = 111"
      Banco.Execute sql
   End If
      
   sql = " ALTER TABLE INTERNO_DADOS_OBSTETRICO ADD DADOS_OB_PARTOS_ANTERIORES_ABORTO INT "
   Banco.Execute sql

   sql = " ALTER TABLE INTERNO_DADOS_OBSTETRICO ADD DADOS_OB_TIPOGESTACAO INT"
   Banco.Execute sql
   
   'SEMPRE INSERIR ESTA INSTRUCAO NAS NOVAS FUNCOES DOS MESES SEGUINTES, ATUALIZANDO SOMENTE O VALOR
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '112011' referente ao mês de novembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '112011'"
   Banco.Execute sql
      
   'ESPECIALIDADECIRURGICA
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes122011()
   On Error GoTo Erro

   sql = " ALTER TABLE ESPECIALIDADECIRURGICA ADD COTA INT "
   Banco.Execute sql

   sql = " ALTER TABLE CATEGORIAPROCEDIMENTO ADD COTA INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE SAME_ESTATISTICA ADD PROCEDIMENTONOME VARCHAR(250)"
   Banco.Execute sql
   
   sql = " ALTER TABLE SAME_ESTATISTICA ADD COTA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE SAME_ESTATISTICA ADD PROCEDIMENTO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD UTILIZA_ANSTISS INT "
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[PRODUTOSALDO]("
   sql = sql & "    [DATA] [datetime] NOT NULL,"
   sql = sql & "    [PRODUTO] [int] NOT NULL,"
   sql = sql & "    [CENTROCUSTO] [smallint] NOT NULL,"
   sql = sql & "    [SALDO] [money] NULL,"
   sql = sql & "    [SALDOUNITARIO] [money] NULL,"
   sql = sql & "    [CUSTO] [money] NULL,"
   sql = sql & "    [VENDA] [money] NULL,"
   sql = sql & "    [CUSTOMEDIO] [money] NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_PRODUTOSALDO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [DATA] ASC,"
   sql = sql & "    [PRODUTO] ASC,"
   sql = sql & "    [CentroCusto] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " ALTER TABLE [dbo].[PRODUTOSALDO]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTOSALDO_CentroCusto] FOREIGN KEY([CENTROCUSTO])"
   sql = sql & " References [dbo].[CentroCusto]([CentroCusto])"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " ALTER TABLE [dbo].[PRODUTOSALDO]  WITH CHECK ADD  CONSTRAINT [FK_PRODUTOSALDO_Produto] FOREIGN KEY([PRODUTO])"
   sql = sql & " References [dbo].[Produto]([Produto])"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE INDEX IX_PRODUTO ON PRODUTOSALDO ("
   sql = sql & " Produto , Data, CENTROCUSTO "
   sql = sql & " )"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TRIGGER TR_SALDO_PRODUTO"
   sql = sql & " ON PRODUTOCENTROCUSTO"
   sql = sql & " With ENCRYPTION"
   sql = sql & " FOR UPDATE, INSERT, DELETE"
   sql = sql & " AS"
   sql = sql & "    IF NOT EXISTS(SELECT DATA"
   sql = sql & "               From PRODUTOSALDO"
   sql = sql & "               WHERE DATA = DATEADD(d, -1, CONVERT(VARCHAR(10),GETDATE(),111)))"
   sql = sql & "    BEGIN"
   sql = sql & "       INSERT INTO PRODUTOSALDO (DATA, CENTROCUSTO, PRODUTO,"
   sql = sql & "             CUSTO, VENDA, CUSTOMEDIO, SALDO, SALDOUNITARIO, ATUALIZACAO )"
   sql = sql & "       SELECT DATEADD(d, -1, CONVERT(VARCHAR(10),GETDATE(),111)) AS DATA, A.CENTROCUSTO, A.PRODUTO,"
   sql = sql & "            ISNULL(BRASINDICECUSTO,0), ISNULL(BRASINDICEVENDA,0), ISNULL(CUSTOMEDIO1,0),"
   sql = sql & "            SUM(A.SALDO), SUM(A.SALDOUNITARIO),"
   sql = sql & "            GETDATE()"
   sql = sql & "       FROM PRODUTOCENTROCUSTO A INNER JOIN PRODUTO ON A.PRODUTO = PRODUTO.PRODUTO"
   sql = sql & "                           INNER JOIN CENTROCUSTO C ON A.CENTROCUSTO = C.CENTROCUSTO"
   sql = sql & "       WHERE 1 = 1"
   sql = sql & "       AND ISNULL(SOMENTEFATURAMENTO,0) = 0"
   sql = sql & "       AND PRODUTO.EXCLUIDO = 0"
   sql = sql & "       AND C.MOVIMENTASALDO = 1"
   sql = sql & "       AND ISNULL(BALANCO_RECEITA_DESPESA,0) = 0"
   sql = sql & "       AND A.AUTORIZADO = 1"
   sql = sql & "       AND ISNULL(A.SALDO,0) > 0"
   sql = sql & "       GROUP BY A.CENTROCUSTO, A.PRODUTO, PRODUTO.BRASINDICECUSTO,"
   sql = sql & "                Produto.BRASINDICEVENDA , Produto.BRASINDICEFABRICA, Produto.CUSTOMEDIO1"
   sql = sql & "   END "
   Banco.Execute sql
      
   sql = " ALTER TABLE DMED ADD BENEFICIARIO VARCHAR(250)"
   Banco.Execute sql

   sql = " ALTER TABLE DMED ADD CPFBENEFICIARIO VARCHAR(50)"
   Banco.Execute sql

   sql = " ALTER TABLE DMED ADD BENEFICIARIO_NASCIMENTO DATETIME"
   Banco.Execute sql
   
   sql = " ALTER TABLE CIRURGIA ADD ENFERMEIRORPA INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE CIRURGIA ADD LEITORPA CHAR(15)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[CIRURGIA_PATOLOGIA]("
   sql = sql & "    [SEQUENCIA] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & "    [REGISTRO] [int] NULL,"
   sql = sql & "    [DOCUMENTO] [int] NULL,"
   sql = sql & "    [TIPOREGISTRO] [int] NULL,"
   sql = sql & "    [PATOLOGIA] [int] NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & " CONSTRAINT [PK_CIRURGIA_PATOLOGIA] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [Sequencia] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = " "
   sql = sql & " ALTER TABLE [dbo].[CIRURGIA_PATOLOGIA]  WITH CHECK ADD  CONSTRAINT [FK_CIRURGIA_PATOLOGIA_PATOLOGIA] FOREIGN KEY([PATOLOGIA])"
   sql = sql & " References [dbo].[Patologia]([Patologia])"
   Banco.Execute sql

   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '122011'"
   Banco.Execute sql
      
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes012012()
   On Error GoTo Erro

   sql = "ALTER TABLE FILANCAMENTO "
   sql = sql & " add SERIE VARCHAR(5)"
   Banco.Execute sql

   sql = "ALTER TABLE MOVIMENTOAVULSO"
   sql = sql & " ADD SERIE VARCHAR(5)"
   Banco.Execute sql

   sql = "ALTER TABLE PEDIDOCOMPRAPRODUTONOTA1"
   sql = sql & " ADD SERIE VARCHAR(5)"
   Banco.Execute sql

   sql = "ALTER TABLE PEDIDOCOMPRANOTA1"
   sql = sql & " ADD SERIE VARCHAR(5)"
   Banco.Execute sql

   sql = "ALTER TABLE tmpPRODUTOCOMPRANOTA"
   sql = sql & " ADD SERIE VARCHAR(5)"
   Banco.Execute sql

   sql = "ALTER TABLE tmpPRODUTOCOMPRAPRODUTONOTA"
   sql = sql & " ADD SERIE VARCHAR(5)"
   Banco.Execute sql

   sql = "ALTER TABLE MOVIMENTOPRODUTODIVERSOS"
   sql = sql & " ADD SERIE VARCHAR(5)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPINTERNACAO "
   sql = sql & "ADD ULTIMASINTERNACOES VARCHAR(5000)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPINTERNACAO "
   sql = sql & "ADD LEITOINTERNACAO VARCHAR(20)"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD UTILIZA_ANSTISS INT "
   Banco.Execute sql
      
   If Layout = 19 Then  'porto seguro
      sql = " UPDATE CONVENIOS SET CODIGOMATERIALTISS = NULL, CODIGOMEDICAMENTOTISS = NULL WHERE CONVENIO = 111"
      Banco.Execute sql
   End If
   
   sql = " ALTER TABLE CONT_EXPORTACAO ADD DATAEXPORTACAO DATETIME   "
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO"
   sql = sql & " drop column CAMINHOBKPSOLICITACAOELEGIBILIDADE VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE PARAMETRO"
   sql = sql & " drop column CAMINHOBKPRESPOSTAELEGIBILIDADE VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE PARAMETRO"
   sql = sql & " drop column CAMINHOBKPSOLICITACAOPROCEDIMENTO VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE PARAMETRO"
   sql = sql & " drop column CAMINHOBKPRESPOSTAPROCEDIMENTO VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE PARAMETRO"
   sql = sql & " ADD SEQUENCIA BIGINT"
   Banco.Execute sql

   sql = "CREATE TABLE UNIMEDS("
   sql = sql & " UNIDADE VARCHAR(4) NOT NULL,"
   sql = sql & " DESCRICAO VARCHAR(255),"
   sql = sql & " FESP BIT,"
   sql = sql & " Convenio SmallInt"
   sql = sql & " )"
   Banco.Execute sql

   sql = "ALTER TABLE UNIMEDS ADD  CONSTRAINT [PK_UNIMEDS] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & " Unidade Asc"
   sql = sql & " )WITH (SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
   Banco.Execute sql

   sql = "ALTER TABLE UNIMEDS  ADD  CONSTRAINT [FK_Unimeds_Convenio] FOREIGN KEY([CONVENIO])"
   sql = sql & " References CONVENIOS(Convenio)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD RESPOSTAELEGIBILIDADE VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD RESPOSTAELEGIBILIDADE VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD RESPOSTAELEGIBILIDADE VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_PLANO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_PLANO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_PLANO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_NOMEPLANO VARCHAR(50)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_NOMEPLANO VARCHAR(50)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_NOMEPLANO VARCHAR(50)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_EMPRESACODIGO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_EMPRESACODIGO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_EMPRESACODIGO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_EMPRESARAZAO VARCHAR(255)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_EMPRESARAZAO VARCHAR(255)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_EMPRESARAZAO VARCHAR(255)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_SEQUENCIALTRANSACAO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_SEQUENCIALTRANSACAO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_SEQUENCIALTRANSACAO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_DATAREGISTROTRANSACAO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_DATAREGISTROTRANSACAO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_DATAREGISTROTRANSACAO VARCHAR(20)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_HORAREGISTROTRANSACAO VARCHAR(10)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_HORAREGISTROTRANSACAO VARCHAR(10)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_HORAREGISTROTRANSACAO VARCHAR(10)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_CODIGOGLOSA VARCHAR(10)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_CODIGOGLOSA VARCHAR(10)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_CODIGOGLOSA VARCHAR(10)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_DESCRICAOGLOSA VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_DESCRICAOGLOSA VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_DESCRICAOGLOSA VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO"
   sql = sql & " ADD ELE_OBSERVACAO VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO"
   sql = sql & " ADD ELE_OBSERVACAO VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIAL"
   sql = sql & " ADD ELE_OBSERVACAO VARCHAR(500)"
   Banco.Execute sql

   sql = "ALTER TABLE FICHAS"
   sql = sql & " ADD IDENTIFICADORBENEFICIARIO TEXT"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_Fichas_Unidade] ON [dbo].[Fichas]"
   sql = sql & " ("
   sql = sql & "  Unidade Asc"
   sql = sql & " )WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_Fichas_Carteirinha] ON [dbo].[Fichas]"
   sql = sql & " ("
   sql = sql & " Carteirinha Asc"
   sql = sql & " )WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
   Banco.Execute sql

   sql = "ALTER TABLE MEDICOS"
   sql = sql & " DROP CONSTRAINT DF_Medicos_CodUnimed"
   Banco.Execute sql

   sql = "ALTER TABLE MEDICOS"
   sql = sql & " ALTER COLUMN CODUNIMED BIGINT"
   Banco.Execute sql

   sql = "ALTER TABLE [dbo].[Medicos] ADD  CONSTRAINT [DF_Medicos_CodUnimed]  DEFAULT (0) FOR [CodUnimed]"
   Banco.Execute sql

   sql = "CREATE TABLE AUTORIZACAOPROCEDIMENTO("
   sql = sql & " AUTORIZACAOPROCEDIMENTO BIGINT IDENTITY(1,1),"
   sql = sql & " NUMEROCARTEIRINHA VARCHAR(20),"
   sql = sql & " NOME VARCHAR(150),"
   sql = sql & " NOMEPLANO VARCHAR(50),"
   sql = sql & " DATAREGISTROTRANSACAO DATETIME,"
   sql = sql & " DATAEMISSAOGUIA DATETIME,"
   sql = sql & " NUMEROGUIAPRESTADOR VARCHAR(20),"
   sql = sql & " NUMEROGUIAOPERADORA VARCHAR(20),"
   sql = sql & " CODIGOPRESTADORNAOPERADORA VARCHAR(15),"
   sql = sql & " NOMECONTRATADO VARCHAR(100),"
   sql = sql & " DATAAUTORIZACAO DATETIME,"
   sql = sql & " SENHAAUTORIZACAO VARCHAR(20),"
   sql = sql & " CODIGOPROCEDIMENTO VARCHAR(20),"
   sql = sql & " TIPOTABELA VARCHAR(5),"
   sql = sql & " DESCRICAOPROCEDIMENTO VARCHAR(100),"
   sql = sql & " QUANTIDADESOLICITADA MONEY,"
   sql = sql & " QUANTIDADEAUTORIZADA MONEY,"
   sql = sql & " STATUSSOLICITACAOPROCEDIMENTO INT,"
   sql = sql & " OBSERVACAO VARCHAR(250),"
   sql = sql & " NOMEARQUIVO VARCHAR(250),"
   sql = sql & " TIPOREGISTRO VARCHAR(15),"
   sql = sql & " REGISTRO BIGINT,"
   sql = sql & " CONSTRAINT PK_AUTORIZACAOPROCEDIMENTO PRIMARY KEY CLUSTERED(AUTORIZACAOPROCEDIMENTO)"
   sql = sql & " )"
   Banco.Execute sql

   sql = "ALTER TABLE INTERNOPROCEDIMENTO"
   sql = sql & " ADD AUTORIZACAOPROCEDIMENTO BIGINT"
   Banco.Execute sql

   sql = "ALTER TABLE AMBULATORIALPROCEDIMENTO"
   sql = sql & " ADD AUTORIZACAOPROCEDIMENTO BIGINT"
   Banco.Execute sql

   sql = "CREATE TABLE AUTORIZACAOPROCEDIMENTOGLOSAS("
   sql = sql & " AUTORIZACAOPROCEDIMENTOGLOSAS BIGINT IDENTITY(1,1),"
   sql = sql & " AUTORIZACAOPROCEDIMENTO BIGINT,"
   sql = sql & " CODIGOGLOSA INT,"
   sql = sql & " descricaoGlosa VarChar(501)"
   sql = sql & " CONSTRAINT PK_AUTORIZACAOPROCEDIMENTOGLOSAS PRIMARY KEY (AUTORIZACAOPROCEDIMENTOGLOSAS),"
   sql = sql & " CONSTRAINT FK_AUTORIZACAOPROCEDIMENTOGLOSAS_AUTORIZACAOPROCEDIMENTO FOREIGN KEY (AUTORIZACAOPROCEDIMENTO) REFERENCES AUTORIZACAOPROCEDIMENTO(AUTORIZACAOPROCEDIMENTO)"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = " alter table TMP_UNIDADEPACIENTE ADD QUANTIDADEEXCEDIDA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE FICHAS ADD IDADECONJUGE INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE tmpINTERNACAO ADD IDADECONJUGE INT"
   Banco.Execute sql
   
   
   
   sql = " ALTER TABLE CIRURGIA ADD RPA_DATAINICIAL DATETIME "
   Banco.Execute sql
   
   sql = " ALTER TABLE CIRURGIA ADD RPA_DATAFINAL DATETIME "
   Banco.Execute sql
   
   sql = " ALTER TABLE CIRURGIA ADD RPA_HORAINICIAL CHAR(5)"
   Banco.Execute sql

   sql = " ALTER TABLE CIRURGIA ADD RPA_HORAFINAL CHAR(5)"
   Banco.Execute sql
   
   sql = " ALTER TABLE tmpHONORARIOMEDICO ALTER COLUMN PROCEDIMENTONOME VARCHAR(250)"
   Banco.Execute sql

   sql = " ALTER TABLE tmpHONORARIOMEDICO ALTER COLUMN NOME VARCHAR(250)"
   Banco.Execute sql

   If Val(RecuperaCampo(1, "1", "HOSPITAL_PSIQUIATRA", "PARAMETRO")) = 1 Then
      sql = ""
      sql = sql & " INSERT INTO MENU(MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
      sql = sql & "       NOMESUBAUX,NIVELVISIBILIDADE)"
      sql = sql & " SELECT MAX(MENU)+1,'Tipo de Internação', 'mnuSam_Est_Tin', ' ', 'mnuSam_Est_Tin',"
      sql = sql & "    1, 1, '0302220000', 'mnuCRe_Fat', 1"
      sql = sql & " From Menu"
      Banco.Execute sql
      
      sql = "UPDATE MENU " & Chr(13)
      sql = sql & "SET NOMECAPTION = 'Aniversariantes' " & Chr(13)
      sql = sql & "WHERE NOMECAPTION = 'nascidos'"
      Banco.Execute sql
   End If
   
   sql = " ALTER TABLE INTERNO ADD LEITOINICIAL CHAR(15) "
   Banco.Execute sql
   
   sql = " ALTER TABLE ALTAS ADD TIPOOBITO INT"
   Banco.Execute sql
   
   sql = " UPDATE MENU SET ATIVADO = 0 WHERE NOMESUBNOVO = 'mnuSam_Est_Obi'"
   Banco.Execute sql
   
   sql = " alter table TMP_UNIDADEPACIENTE ADD LEITOEXTRA INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE SAME_ESTATISTICA_INTERNACAO ADD TIPO INT"
   Banco.Execute sql
   
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '012012'"
   Banco.Execute sql
      
   Exit Function
Erro:
   Resume Next
End Function


Public Function AtualizaMes022012()
   On Error GoTo Erro

   sql = " ALTER TABLE ALTAS ADD TIPOOBITO INT"
   Banco.Execute sql


      
   sql = " ALTER TABLE PRODUTO ADD OBSERVACAO_COTACAO VARCHAR(250)"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_tmpMOVIMENTO_DEVOLUCAO_PRODUTO] ON [dbo].[tmpMOVIMENTO_DEVOLUCAO] " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "   [Produto] Asc " & Chr(13)
   sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_tmpMOVIMENTO_DEVOLUCAO_LOTE] ON [dbo].[tmpMOVIMENTO_DEVOLUCAO] " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "   [Lote] Asc " & Chr(13)
   sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & Chr(13)
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_tmpMOVIMENTO_DEVOLUCAO_VALIDADELOTE] ON [dbo].[tmpMOVIMENTO_DEVOLUCAO] " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "   ValidadeLote Asc " & Chr(13)
   sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_tmpMOVIMENTO_DEVOLUCAO_IP] ON [dbo].[tmpMOVIMENTO_DEVOLUCAO] " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "   IP Asc " & Chr(13)
   sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
   Banco.Execute sql

   sql = "CREATE NONCLUSTERED INDEX [IX_PRODUTO_LOTE_ETIQUETA_SEQUENCIA] ON [dbo].[PRODUTO_LOTE_ETIQUETA] " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "   Sequencia Asc " & Chr(13)
   sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
   Banco.Execute sql
   sql = "CREATE TABLE [dbo].[AUTORIZACAOPROCEDIMENTOAMBULATORIAL]( " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTO] [bigint] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "      [NUMEROCARTEIRINHA] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [NOME] [varchar](150) NULL, " & Chr(13)
   sql = sql & "      [NOMEPLANO] [varchar](50) NULL, " & Chr(13)
   sql = sql & "      [DATAREGISTROTRANSACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "      [DATAEMISSAOGUIA] [datetime] NULL, " & Chr(13)
   sql = sql & "      [NUMEROGUIAPRESTADOR] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [NUMEROGUIAOPERADORA] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [CODIGOPRESTADORNAOPERADORA] [varchar](15) NULL, " & Chr(13)
   sql = sql & "      [NOMECONTRATADO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "      [DATAAUTORIZACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "      [SENHAAUTORIZACAO] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [VALIDADESENHA] [datetime] NULL, " & Chr(13)
   sql = sql & "      [DIASAUTORIZADO] [int] NULL, " & Chr(13)
   sql = sql & "       [DATAPROVAVELADMISHOSP] [datetime] NULL, " & Chr(13)
   sql = sql & "       [TIPOACOMODACAO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "      [CODIGOPROCEDIMENTO] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [TIPOTABELA] [varchar](5) NULL, " & Chr(13)
   sql = sql & "      [DESCRICAOPROCEDIMENTO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "      [QUANTIDADESOLICITADA] [money] NULL, " & Chr(13)
   sql = sql & "      [QUANTIDADEAUTORIZADA] [money] NULL, " & Chr(13)
   sql = sql & "      [STATUSSOLICITACAOPROCEDIMENTO] [int] NULL, " & Chr(13)
   sql = sql & "      [OBSERVACAO] [varchar](250) NULL, " & Chr(13)
   sql = sql & "      [NOMEARQUIVO] [varchar](250) NULL, " & Chr(13)
   sql = sql & "      [TIPOREGISTRO] [varchar](15) NULL, " & Chr(13)
   sql = sql & "      [REGISTRO] [bigint] NULL, " & Chr(13)
   sql = sql & "    CONSTRAINT [PK_AUTORIZACAOPROCEDIMENTOAMBULATORIAL] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "   ( " & Chr(13)
   sql = sql & "      [autorizacaoProcedimento] Asc " & Chr(13)
   sql = sql & "   )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & Chr(13)
   sql = sql & "   ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[AUTORIZACAOPROCEDIMENTOEXTERNO]( " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTO] [bigint] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "      [NUMEROCARTEIRINHA] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [NOME] [varchar](150) NULL, " & Chr(13)
   sql = sql & "      [NOMEPLANO] [varchar](50) NULL, " & Chr(13)
   sql = sql & "      [DATAREGISTROTRANSACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "      [DATAEMISSAOGUIA] [datetime] NULL, " & Chr(13)
   sql = sql & "      [NUMEROGUIAPRESTADOR] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [NUMEROGUIAOPERADORA] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [CODIGOPRESTADORNAOPERADORA] [varchar](15) NULL, " & Chr(13)
   sql = sql & "      [NOMECONTRATADO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "      [DATAAUTORIZACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "      [SENHAAUTORIZACAO] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [VALIDADESENHA] [datetime] NULL, " & Chr(13)
   sql = sql & "      [DIASAUTORIZADO] [int] NULL, " & Chr(13)
   sql = sql & "       [DATAPROVAVELADMISHOSP] [datetime] NULL, " & Chr(13)
   sql = sql & "       [TIPOACOMODACAO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "      [CODIGOPROCEDIMENTO] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [TIPOTABELA] [varchar](5) NULL, " & Chr(13)
   sql = sql & "      [DESCRICAOPROCEDIMENTO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "      [QUANTIDADESOLICITADA] [money] NULL, " & Chr(13)
   sql = sql & "      [QUANTIDADEAUTORIZADA] [money] NULL, " & Chr(13)
   sql = sql & "      [STATUSSOLICITACAOPROCEDIMENTO] [int] NULL, " & Chr(13)
   sql = sql & "      [OBSERVACAO] [varchar](250) NULL, " & Chr(13)
   sql = sql & "      [NOMEARQUIVO] [varchar](250) NULL, " & Chr(13)
   sql = sql & "      [TIPOREGISTRO] [varchar](15) NULL, " & Chr(13)
   sql = sql & "      [REGISTRO] [bigint] NULL, " & Chr(13)
   sql = sql & "    CONSTRAINT [PK_AUTORIZACAOPROCEDIMENTOEXTERNO] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "   ( " & Chr(13)
   sql = sql & "      [autorizacaoProcedimento] Asc " & Chr(13)
   sql = sql & "   )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & Chr(13)
   sql = sql & "   ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[AUTORIZACAOPROCEDIMENTOINTERNO]( " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTO] [bigint] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "      [NUMEROCARTEIRINHA] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [NOME] [varchar](150) NULL, " & Chr(13)
   sql = sql & "      [NOMEPLANO] [varchar](50) NULL, " & Chr(13)
   sql = sql & "      [DATAREGISTROTRANSACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "      [DATAEMISSAOGUIA] [datetime] NULL, " & Chr(13)
   sql = sql & "      [NUMEROGUIAPRESTADOR] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [NUMEROGUIAOPERADORA] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [CODIGOPRESTADORNAOPERADORA] [varchar](15) NULL, " & Chr(13)
   sql = sql & "      [NOMECONTRATADO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "      [DATAAUTORIZACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "      [SENHAAUTORIZACAO] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [VALIDADESENHA] [datetime] NULL, " & Chr(13)
   sql = sql & "      [DIASAUTORIZADO] [int] NULL, " & Chr(13)
   sql = sql & "       [DATAPROVAVELADMISHOSP] [datetime] NULL, " & Chr(13)
   sql = sql & "       [TIPOACOMODACAO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "      [CODIGOPROCEDIMENTO] [varchar](20) NULL, " & Chr(13)
   sql = sql & "      [TIPOTABELA] [varchar](5) NULL, " & Chr(13)
   sql = sql & "      [DESCRICAOPROCEDIMENTO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "      [QUANTIDADESOLICITADA] [money] NULL, " & Chr(13)
   sql = sql & "      [QUANTIDADEAUTORIZADA] [money] NULL, " & Chr(13)
   sql = sql & "      [STATUSSOLICITACAOPROCEDIMENTO] [int] NULL, " & Chr(13)
   sql = sql & "      [OBSERVACAO] [varchar](250) NULL, " & Chr(13)
   sql = sql & "      [NOMEARQUIVO] [varchar](250) NULL, " & Chr(13)
   sql = sql & "      [TIPOREGISTRO] [varchar](15) NULL, " & Chr(13)
   sql = sql & "      [REGISTRO] [bigint] NULL, " & Chr(13)
   sql = sql & "    CONSTRAINT [PK_AUTORIZACAOPROCEDIMENTOINTERNO] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "   ( " & Chr(13)
   sql = sql & "      [autorizacaoProcedimento] Asc " & Chr(13)
   sql = sql & "   )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & Chr(13)
   sql = sql & "   ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[AUTORIZACAOPROCEDIMENTOAMBULATORIALGLOSAS]( " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTOGLOSAS] [bigint] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTO] [bigint] NULL, " & Chr(13)
   sql = sql & "      [CODIGOGLOSA] [int] NULL, " & Chr(13)
   sql = sql & "      [DESCRICAOGLOSA] [varchar](501) NULL, " & Chr(13)
   sql = sql & "    CONSTRAINT [PK_AUTORIZACAOPROCEDIMENTOAMBULATORIALGLOSAS] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "   ( " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTOGLOSAS] Asc " & Chr(13)
   sql = sql & "   )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & Chr(13)
   sql = sql & "   ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[AUTORIZACAOPROCEDIMENTOAMBULATORIALGLOSAS]  WITH CHECK ADD  CONSTRAINT [FK_AUTORIZACAOPROCEDIMENTOAMBULATORIALGLOSAS_AUTORIZACAOPROCEDIMENTO] FOREIGN KEY([AUTORIZACAOPROCEDIMENTO]) " & Chr(13)
   sql = sql & "References [dbo].[AUTORIZACAOPROCEDIMENTOAMBULATORIAL]([autorizacaoProcedimento])"
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[AUTORIZACAOPROCEDIMENTOAMBULATORIALGLOSAS] CHECK CONSTRAINT [FK_AUTORIZACAOPROCEDIMENTOAMBULATORIALGLOSAS_AUTORIZACAOPROCEDIMENTO]"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[AUTORIZACAOPROCEDIMENTOEXTERNOGLOSAS]( " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTOGLOSAS] [bigint] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTO] [bigint] NULL, " & Chr(13)
   sql = sql & "      [CODIGOGLOSA] [int] NULL, " & Chr(13)
   sql = sql & "      [DESCRICAOGLOSA] [varchar](501) NULL, " & Chr(13)
   sql = sql & "    CONSTRAINT [PK_AUTORIZACAOPROCEDIMENTOEXTERNOGLOSAS] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "   ( " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTOGLOSAS] Asc " & Chr(13)
   sql = sql & "   )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & Chr(13)
   sql = sql & "   ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[AUTORIZACAOPROCEDIMENTOEXTERNOGLOSAS]  WITH CHECK ADD  CONSTRAINT [FK_AUTORIZACAOPROCEDIMENTOEXTERNOGLOSAS_AUTORIZACAOPROCEDIMENTO] FOREIGN KEY([AUTORIZACAOPROCEDIMENTO]) " & Chr(13)
   sql = sql & "References [dbo].[AUTORIZACAOPROCEDIMENTOEXTERNO]([autorizacaoProcedimento]) "
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[AUTORIZACAOPROCEDIMENTOEXTERNOGLOSAS] CHECK CONSTRAINT [FK_AUTORIZACAOPROCEDIMENTOEXTERNOGLOSAS_AUTORIZACAOPROCEDIMENTO]"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[AUTORIZACAOPROCEDIMENTOINTERNOGLOSAS]( " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTOGLOSAS] [bigint] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTO] [bigint] NULL, " & Chr(13)
   sql = sql & "      [CODIGOGLOSA] [int] NULL, " & Chr(13)
   sql = sql & "      [DESCRICAOGLOSA] [varchar](501) NULL, " & Chr(13)
   sql = sql & "    CONSTRAINT [PK_AUTORIZACAOPROCEDIMENTOINTERNOGLOSAS] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "   ( " & Chr(13)
   sql = sql & "      [AUTORIZACAOPROCEDIMENTOGLOSAS] Asc " & Chr(13)
   sql = sql & "   )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & Chr(13)
   sql = sql & "   ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[AUTORIZACAOPROCEDIMENTOINTERNOGLOSAS]  WITH CHECK ADD  CONSTRAINT [FK_AUTORIZACAOPROCEDIMENTOINTERNOGLOSAS_AUTORIZACAOPROCEDIMENTO] FOREIGN KEY([AUTORIZACAOPROCEDIMENTO]) " & Chr(13)
   sql = sql & "    References [dbo].[AUTORIZACAOPROCEDIMENTOINTERNO]([autorizacaoProcedimento]) "
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[AUTORIZACAOPROCEDIMENTOINTERNOGLOSAS] CHECK CONSTRAINT [FK_AUTORIZACAOPROCEDIMENTOINTERNOGLOSAS_AUTORIZACAOPROCEDIMENTO]"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO " & Chr(13)
   sql = sql & "ADD DATAREGISTROTRANSACAOSTATUSAUTO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL " & Chr(13)
   sql = sql & "ADD DATAREGISTROTRANSACAOSTATUSAUTO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO " & Chr(13)
   sql = sql & "ADD DATAREGISTROTRANSACAOSTATUSAUTO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO " & Chr(13)
   sql = sql & "ADD NOMEARQUIVOSTATUSAUTO VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL " & Chr(13)
   sql = sql & "ADD NOMEARQUIVOSTATUSAUTO VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO " & Chr(13)
   sql = sql & "ADD NOMEARQUIVOSTATUSAUTO VARCHAR(250)"
   Banco.Execute sql
   
   sql = " ALTER TABLE FICHAS ADD IDADECONJUGE INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE tmpINTERNACAO ADD IDADECONJUGE INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE TER_PARAMETRO ADD CALCULAISSDATA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE DADOSAIH ADD DATAFATURAMENTO_APRESENTACAO DATETIME "
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPRELDADOSCIRURGIA ADD MOTIVOATRASO VARCHAR(100)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPRELDADOSCIRURGIA ADD OBSERVACAO VARCHAR(255)"
   Banco.Execute sql
   
   sql = " ALTER TABLE FORNECEDORES ADD ATIVIDADE_ISS INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE FICREDOR ADD ATIVIDADE_ISS INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[ATIVIDADE_ISS]("
   sql = sql & "    [ATIVIDADE] [int] NOT NULL,"
   sql = sql & "    [NOME] [nvarchar](255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,"
   sql = sql & "    [INCIDENCIA] [float] NULL,"
   sql = sql & "    [CODIGOISS] [float] NULL,"
   sql = sql & "  CONSTRAINT [PK_ATIVIDADE_ISS] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [Atividade] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPFATURAMENTOSUSMEDICO " & Chr(13)
   sql = sql & "ADD QUANTIDADE MONEY "
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '022012'"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes032012()
   On Error GoTo Erro

   
   sql = " ALTER TABLE EXTERNO ADD ENCAIXE INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD ENCAIXE INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CIRURGIA ADD OUTRAPATOLOGIA VARCHAR(100)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD APAC_QUIMIO_CITO_HISTOLOGICO VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD DATAENTRADAOBSERVACAO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD HORAENTRADAOBSERVACAO DATETIME"
   Banco.Execute sql
   
   If Layout = 44 Then
      sql = " UPDATE MENU SET NOMECAPTION = 'Prontuário Eletrônico' where NOMESUBNOVO = 'mnuMed_EPr'"
      Banco.Execute sql
   End If
   
   sql = " ALTER TABLE AGENDAMENTOCONSULTA ADD DATAINCLUSAO DATETIME"
   Banco.Execute sql
   
   If Layout = 18 Then 'ARTUR NOG, ASSIMEDICA
      sql = " UPDATE CONVENIOS SET TISS_GUIA_OPERADORA_AMB = 1 WHERE CONVENIO = 47 "
      Banco.Execute sql
   End If
   
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Ficha Kardex', 'mnuPrd_Rel_Kar', ' ', 'mnuPrd_Rel_Kar',"
   sql = sql & "        1, 2, '0112110000', 'mnuPrd_Rel_Kar', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD NOMEFANTASIA VARCHAR(50)"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD HONORARIO_PA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE ALTAS ADD LIBERACAOLEITO VARCHAR(20)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU(MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "       NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Limpeza Leitos', 'mnuEnf_Lim', ' ', 'mnuEnf_Lim',"
   sql = sql & "    1, 1, '0511000000', 'mnuEnf_Lim', 1"
   sql = sql & " From Menu"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPPRODUTOGENERICO ADD CUSTOGENERICO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPPRODUTOGENERICO ADD VENDAGENERICO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE CENTROCUSTO ADD LIMPEZALEITO BIT"
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '032012'"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function


Public Function AtualizaMes042012()
   On Error GoTo Erro

   sql = " ALTER TABLE PARAMETRO ADD TIPOHOSPITAL INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD CORHOSPITAL VARCHAR(50)"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD NOMEFANTASIA VARCHAR(50)"
   Banco.Execute sql
   
   If Layout = 44 Then
      sql = "UPDATE MENU " & Chr(13)
      sql = sql & "SET NOMECAPTION = 'Aniversariantes' " & Chr(13)
      sql = sql & "WHERE NOMECAPTION = 'Relação Ministério'"
      Banco.Execute sql
   End If
   
   sql = ""
   sql = sql & "CREATE TABLE [dbo].[TUSS_CONVENIO_PROCEDIMENTO]("
   sql = sql & "   [TUSS] [int] NOT NULL,"
   sql = sql & "   [TABELA] [nchar](15) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "   [CONVENIO] [int] NOT NULL,"
   sql = sql & "   [CODIGO] [int] NOT NULL,"
   sql = sql & " CONSTRAINT [PK_TUSS_CONVENIO_PROCEDIMENTO] PRIMARY KEY CLUSTERED"
   sql = sql & "("
   sql = sql & "   [TUSS] ASC,"
   sql = sql & "   [TABELA] ASC,"
   sql = sql & "   [CONVENIO] ASC,"
   sql = sql & "   [Codigo] Asc"
   sql = sql & ")WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & ") ON [PRIMARY]"
   Banco.Execute sql
   
   sql = " ALTER TABLE PEDIDOCOMPRA1 ADD OBSERVACAO_PEDIDO VARCHAR(250)"
   Banco.Execute sql
      
   sql = " ALTER TABLE EXTERNO ADD APAC_QUIMIO_CONTINUIDADE_TRATAMENTO INT "
   Banco.Execute sql
      
   sql = " ALTER TABLE PRODUTO ADD OPM INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRODUTO ADD DESCRICAOETIQUETA VARCHAR(200)"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD TISS_CODIGOOPM INT"
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '042012'"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes052012()
   On Error GoTo Erro

   sql = " ALTER TABLE PARAMETRO ADD TIPOHOSPITAL INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD CORHOSPITAL VARCHAR(50)"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD NOMEFANTASIA VARCHAR(50)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[CONT_SEQUENCIA]("
   sql = sql & "    [ANO] [int] NOT NULL,"
   sql = sql & "    [MES] [int] NOT NULL,"
   sql = sql & "    [LOTE] [int] NOT NULL,"
   sql = sql & "    [SEQUENCIA] [int] NOT NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_CONT_SEQUENCIA] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [ANO] ASC,"
   sql = sql & "    [MES] ASC,"
   sql = sql & "    [LOTE] ASC,"
   sql = sql & "    [Sequencia] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[PRODUTOSALDOLOTE]("
   sql = sql & "    [DATA] [datetime] NOT NULL,"
   sql = sql & "    [PRODUTO] [int] NOT NULL,"
   sql = sql & "    [CENTROCUSTO] [int] NOT NULL,"
   sql = sql & "    [LOTE] [char](15) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "    [VALIDADELOTE] [datetime] NOT NULL,"
   sql = sql & "    [SALDO] [money] NULL,"
   sql = sql & "    [SALDOUNITARIO] [money] NULL,"
   sql = sql & "    [ATUALIZACAO] [datetime] NULL,"
   sql = sql & "  CONSTRAINT [PK_PRODUTOSALDOLOTE] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [DATA] ASC,"
   sql = sql & "    [PRODUTO] ASC,"
   sql = sql & "    [CENTROCUSTO] ASC,"
   sql = sql & "    [LOTE] ASC,"
   sql = sql & "    [ValidadeLote] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = " ALTER TABLE MEDICOS ADD MEDICO_PJ INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOVIM_INT ADD PLANTAO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOVIM_AMB ADD PLANTAO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE MOVIM_EXT ADD PLANTAO INT"
   Banco.Execute sql
   
   sql = " alter table tmpINTERNACAO ALTER COLUMN NACIONALIDADE VARCHAR(100)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD ATESTADO VARCHAR(2000)"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIALPROCEDIMENTO ADD QUANTIDADE_AUX INT  "
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPTISS_SPSADT ADD TIPOREGISTRO VARCHAR(30)"
   Banco.Execute sql

   sql = " ALTER TABLE TMPTISS_INTERNO ADD TIPOREGISTRO VARCHAR(30)"
   Banco.Execute sql

   sql = " ALTER TABLE EXTERNO ADD LOTEGUIA_IAM VARCHAR(20)"
   Banco.Execute sql

   sql = " ALTER TABLE INTERNO ADD LOTEGUIA_IAM VARCHAR(20)"
   Banco.Execute sql

   sql = " ALTER TABLE AMBULATORIAL ADD LOTEGUIA_IAM VARCHAR(20)"
   Banco.Execute sql

   sql = " ALTER TABLE EXTERNO ADD IAM_TIPOFECHAMENTO INT"
   Banco.Execute sql

   sql = " ALTER TABLE INTERNO ADD IAM_TIPOFECHAMENTO INT"
   Banco.Execute sql

   sql = " ALTER TABLE AMBULATORIAL ADD IAM_TIPOFECHAMENTO INT"
   Banco.Execute sql

   
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Conferência de Lote', 'mnuFat_Iam_Amb_Con', ' ', 'mnuFat_Iam_Amb_Con',"
   sql = sql & "        1, 1, '0722030100', 'mnuFat_Iam_Amb_Con', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_AMB_TMP_SUS ADD PLANTAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_INT_TMP_SUS ADD PLANTAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE USUARIO ADD GERENTE INT"
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '052012'"
   Banco.Execute sql
   
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes062012()
   On Error GoTo Erro

   sql = " ALTER TABLE PRODUTO ADD MEDICAMENTORESTRITO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD PERCENTUALMEDICAMENTORESTRITO MONEY"
   Banco.Execute sql
   
   If Layout = 42 Then
      sql = "UPDATE MENU SET ATIVADO = 1 WHERE NOMESUBNOVO = 'mnuFin_Cop_PPa'"
      Banco.Execute sql
   End If
      
   sql = " ALTER TABLE CONT_LANCAMENTO_EXTERNO ADD CC_DEBITO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONT_LANCAMENTO_EXTERNO ADD CC_CREDITO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE USUARIO ADD GERENTE INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PRODUTOCONVENIO ADD CONVLANCCONVENIO INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE PRODUTOCONVENIO ADD UNIDADEFAT CHAR(10)"
   Banco.Execute sql
      
   If Layout = 13 Then
      sql = " UPDATE CONVENIOS SET TISS_CODIGOOPM = 1 WHERE CONVENIO = 8   "
      Banco.Execute sql
   End If
   
   sql = " ALTER TABLE CONVENIOS ADD BRAS_TIPOPRECO_CONVENIO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD VALORTOTALCONTA_SADT MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPCHEQUE ADD VALORDESPESABANCARIA MONEY "
   Banco.Execute sql
   
   sql = " ALTER TABLE EXTERNO ADD VALORTOTALCONTA_SADT MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD IAM_AMBULATORIAL MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD IAM_SPSADT MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE EXTERNO ADD IAM_AMBULATORIAL MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE EXTERNO ADD IAM_SPSADT MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE EXTERNO ADD VALORTOTALCONTA_CONSULTA MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD VALORTOTALCONTA_CONSULTA MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE EXTERNO ADD IAM_CONSULTA MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD IAM_CONSULTA MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD LOTEGUIA_IAM_SPSADT VARCHAR(20)"
   Banco.Execute sql
   
   sql = " ALTER TABLE AMBULATORIAL ADD LOTEGUIA_IAM_CONSULTA VARCHAR(20)"
   Banco.Execute sql
   
   sql = " ALTER TABLE EXTERNO ADD LOTEGUIA_IAM_SPSADT VARCHAR(20)"
   Banco.Execute sql
   
   sql = " ALTER TABLE EXTERNO ADD LOTEGUIA_IAM_CONSULTA VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE USUARIO ADD NAOPERMITECANCELARREGISTROS INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE PARAMETRO ADD VALIDA_VALORTOTALNOTA_COMPRADIRETA INT"
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '062012'"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes072012()
   On Error GoTo Erro

   sql = " ALTER TABLE CONVENIOS ADD TISS_INSUMO_UTILIZA_FATORFATURAMENTO INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE INSUMOS ADD FATORFATURAMENTO MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARAMETRO ADD VALIDA_VALORTOTALNOTA_COMPRADIRETA INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE TMP_CONSUMO ("
   sql = sql & " DESCRICAOPRODUTO     VARCHAR(100),"
   sql = sql & " QUANTIDADE           MONEY,"
   sql = sql & " FORNECEDOR           VARCHAR(100),"
   sql = sql & " ENDERECOFORNECEDOR   VARCHAR(100),"
   sql = sql & " NOTA                 VARCHAR(20),"
   sql = sql & " IP                   VARCHAR(100))"
   Banco.Execute sql
   
   sql = "ALTER TABLE LEITOUNIDADE ADD QUANTIDADELEITOOUTROSCONVENIOS INT"
   Banco.Execute sql
      
   sql = " alter table EXTERNO ADD DATAULTIMATRANSFERENCIAPRONTUARIO DATETIME"
   Banco.Execute sql

   sql = " ALTER TABLE EXTERNO ADD HORAULTIMATRANSFERENCIAPRONTUARIO DATETIME"
   Banco.Execute sql

   sql = " ALTER TABLE EXTERNO ADD LOCALARMAZENAMENTO INT"
   Banco.Execute sql

   sql = " ALTER TABLE EXTERNO ADD USUARIOARMAZENAMENTO INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE PRODUTOCONVENIO ADD CODIGO_CONVENIO_TISS INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PRODUTOCONVENIO ADD CONVLANCCONVENIO INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE PRODUTOCONVENIO ADD UNIDADEFAT CHAR(10)"
   Banco.Execute sql
      
   If Layout = 47 Then
      sql = " UPDATE MENU SET ATIVADO = 0 WHERE MENU IN (130, 564)"
      Banco.Execute sql
   End If
   
   sql = "ALTER TABLE INTERNO ADD CCIH_HIPDIAGNOSTICORN VARCHAR(200) "
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD CCIH_HIPDIAGNOSTICOMAE VARCHAR(200)"
   Banco.Execute sql
   
   sql = "ALTER TABLE USUARIO ADD NAOPERMITECANCELARREGISTROS INT"
   Banco.Execute sql
      
   sql = " ALTER TABLE PARAMETRO ADD VALIDA_VALORTOTALNOTA_COMPRADIRETA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE USUARIO ADD GERENTE INT"
   Banco.Execute sql
   
   sql = "CREATE TABLE CCIH_GRUPO_CRITERIO( " & Chr(13)
   sql = sql & " CODIGO [int] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & " DESCRICAO VARCHAR(100), " & Chr(13)
   sql = sql & " ATUALIZACAO VARCHAR(100), " & Chr(13)
   sql = sql & " CONSTRAINT [PK_CCIH_GRUPO_CRITERIO] PRIMARY KEY NONCLUSTERED " & Chr(13)
   sql = sql & " ( " & Chr(13)
   sql = sql & "    CODIGO " & Chr(13)
   sql = sql & " ), " & Chr(13)
   sql = sql & " CONSTRAINT [UN_CCIH_GRUPO_CRITERIO] UNIQUE NONCLUSTERED " & Chr(13)
   sql = sql & " ( " & Chr(13)
   sql = sql & "    DESCRICAO " & Chr(13)
   sql = sql & " ))"
   
   Banco.Execute sql
   
   sql = "ALTER TABLE CCIH_BAR_CRITERIO ADD GRUPO INT "
   Banco.Execute sql

   sql = "ALTER TABLE [dbo].CCIH_BAR_CRITERIO  WITH NOCHECK ADD  CONSTRAINT [FK_CCIH_BAR_CRITERIO_CCIH_GRUPO_CITERIO] FOREIGN KEY(GRUPO) " & Chr(13)
   sql = sql & "REFERENCES [dbo].CCIH_GRUPO_CRITERIO(Codigo)"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[PRES_TUTOR_PRODUTOESPECIFICO]( " & Chr(13)
   sql = sql & "   [PRODUTO] [int] NULL, " & Chr(13)
   sql = sql & "   [TUTOR] [int] NULL, " & Chr(13)
   sql = sql & "   [ATUALIZACAO] [varchar](100) NULL " & Chr(13)
   sql = sql & " ) ON [PRIMARY] "
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD CAMINHORESPOSTADADOSBENE VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE UNIMEDS ADD SOLICITADADOSBENEFICIARIO BIT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD COPARTICIPACAO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD RESPFINANCEIROCPF VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD RESPFINANCEIRONOME VARCHAR(200)"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD RESPFINANCEIRONASCIMENTO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD RESPFINANCEIROEMAIL VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD COPARTICIPACAO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RESPFINANCEIROCPF VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RESPFINANCEIRONOME VARCHAR(200)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RESPFINANCEIRONASCIMENTO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RESPFINANCEIROEMAIL VARCHAR(250)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO ADD COPARTICIPACAO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD RESPFINANCEIROCPF VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD RESPFINANCEIRONOME VARCHAR(200)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD RESPFINANCEIRONASCIMENTO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD RESPFINANCEIROEMAIL VARCHAR(250)"
   Banco.Execute sql
      
   sql = " ALTER TABLE dadosguia ADD ATUALIZACAO VARCHAR(100)"
   Banco.Execute sql
      
   sql = " ALTER TABLE CONT_LANCAMENTO_EXTERNO ADD ADIANTAMENTORECEBIMENTO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE FILANCAMENTORECEBIMENTORECEBIDO ADD VALORADIANTAMENTO MONEY"
   Banco.Execute sql
      
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[FILANCAMENTOADIANTAMENTORECEBIMENTO]("
   sql = sql & "    [LANCAMENTO] [int] NOT NULL,"
   sql = sql & "    [ADIANTAMENTO] [int] NOT NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](155) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "  CONSTRAINT [PK_FILANCAMENTOADIANTAMENTORECEBIMENTO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [LANCAMENTO] ASC,"
   sql = sql & "    [Adiantamento] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[FIADIANTAMENTORECEBIMENTO]("
   sql = sql & " [ADIANTAMENTO] [int] IDENTITY(1,1) NOT NULL,"
   sql = sql & " [CODIGO] [int] NULL,"
   sql = sql & " [TIPOPACIENTE] [int] NULL,"
   sql = sql & " [TIPO] [int] NULL,"
   sql = sql & " [BANCO] [int] NULL,"
   sql = sql & " [TIPOBAIXA] [int] NULL,"
   sql = sql & " [DATA] [datetime] NULL,"
   sql = sql & " [VALOR] [money] NULL,"
   sql = sql & " [DOCUMENTO] [varchar](20) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & " [CENTROCUSTO] [int] NULL,"
   sql = sql & " [GRUPO] [int] NULL,"
   sql = sql & " [OBSERVACAO] [varchar](250) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & " [COMPENSADO] [int] NULL,"
   sql = sql & " [ATUALIZACAO] [varchar](100) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & " CONSTRAINT [PK_FIADIANTAMENTORECEBIMENTO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [Adiantamento] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Adiantamento', 'MnuFin_Cor_Adi', ' ', 'MnuFin_Cor_Adi',"
   sql = sql & "        1, 1, '0902100000', 'MnuFin_Cor_Adi', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql
   
   sql = " ALTER TABLE FILANCAMENTORECEBIMENTORECEBIDO ADD ADIANTAMENTO INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE FIADIANTAMENTORECEBIMENTO ADD DATACOMPENSADO DATETIME"
   Banco.Execute sql
   
   If Layout = 1 Then
      sql = " UPDATE CONVENIOS SET CODIGOTISS  = '923' WHERE CONVENIO = 29"
      Banco.Execute sql
   End If
   
   sql = " ALTER TABLE USUARIO ADD PERMISSAODEVOLUCAO INT"
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '072012'"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes082012()
   On Error GoTo Erro
   
   sql = " ALTER TABLE FILANCAMENTORECEBIMENTO ADD NOTAFISCAL INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD TISS_INSUMO_UTILIZA_FATORFATURAMENTO INT "
   Banco.Execute sql

   sql = "ALTER TABLE INTERNO ADD COPARTICIPACAO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD RESPFINANCEIROCPF VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD RESPFINANCEIRONOME VARCHAR(200)"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD RESPFINANCEIRONASCIMENTO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD RESPFINANCEIROEMAIL VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD COPARTICIPACAO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RESPFINANCEIROCPF VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RESPFINANCEIRONOME VARCHAR(200)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RESPFINANCEIRONASCIMENTO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RESPFINANCEIROEMAIL VARCHAR(250)"
   Banco.Execute sql

   sql = "ALTER TABLE EXTERNO ADD COPARTICIPACAO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD RESPFINANCEIROCPF VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD RESPFINANCEIRONOME VARCHAR(200)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD RESPFINANCEIRONASCIMENTO DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD RESPFINANCEIROEMAIL VARCHAR(250)"
   Banco.Execute sql

   sql = " ALTER TABLE PARTICULAR ADD PORTANES MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD PORTSALA MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD SEXO INT"
   Banco.Execute sql

   sql = " ALTER TABLE PARTICULAR ADD IDADEMIN INT"
   Banco.Execute sql

   sql = " ALTER TABLE PARTICULAR ADD IDADEMAX INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD VALORANTES MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD VALORDEPOIS MONEY"
   Banco.Execute sql

   sql = " ALTER TABLE PARTICULAR ADD DATAALT DATETIME"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD CODIGOFILME INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD CODIGOCONTRASTE INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD ANESTESIA INT"
   Banco.Execute sql

   sql = " ALTER TABLE PARTICULAR ADD CCIHCATEGORIA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD CIRURGIACATEGORIA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD ESPECIALIDADECIRURGICA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD CATEGORIAPROCEDIMENTO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD CIRURGIACATEGORIA INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD SIGLA CHAR(10)"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ADD PARTICULAR INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ALTER COLUMN VALOR MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ALTER COLUMN EspecProcedimento INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE PARTICULAR ALTER COLUMN CONVENIO INT"
   Banco.Execute sql

   sql = "ALTER TABLE USUARIO ADD PERMISSAOPREDEVOLUCAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE USUARIO ADD PERMISSAODEVOLUCAO INT"
   Banco.Execute sql
   
   sql = "CREATE TABLE PRE_DEVOLUCAO( " & Chr(13)
   sql = sql & "  ID_PRE_DEVOLUCAO INT IDENTITY, " & Chr(13)
   sql = sql & "  SITUACAO INT, " & Chr(13)
   sql = sql & "  TIPO_DEVOLUCAO VARCHAR(30), " & Chr(13)
   sql = sql & "  ATUALIZACAO VARCHAR(100), " & Chr(13)
   sql = sql & "  ATUALIZACAOCANCELADO VarChar(100) " & Chr(13)
   sql = sql & " ) " & Chr(13)
   sql = sql & "  ALTER TABLE PRE_DEVOLUCAO ADD  CONSTRAINT [PK_PRE_DEVOLUCAO] PRIMARY KEY NONCLUSTERED " & Chr(13)
   sql = sql & " ( " & Chr(13)
   sql = sql & "    ID_PRE_DEVOLUCAO Asc " & Chr(13)
   sql = sql & ")"
   Banco.Execute sql
   
   sql = "CREATE TABLE PRE_DEVOLUCAO_ITEM( " & Chr(13)
   sql = sql & "   ID_PRE_DEVOLUCAO INT, " & Chr(13)
   sql = sql & "   [SEQUENCIA] [int] NULL, " & Chr(13)
   sql = sql & "   [TIPO] [varchar](60) NULL, " & Chr(13)
   sql = sql & "   [REGISTRO] [int] NULL, " & Chr(13)
   sql = sql & "   [DATACONSUMO] [datetime] NULL, " & Chr(13)
   sql = sql & "   [DATAALTERACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "   [PRODUTO] [int] NULL, " & Chr(13)
   sql = sql & "   [DESCRICAO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "   [QUANTIDADEANTIGA] [money] NULL, " & Chr(13)
   sql = sql & "   [QUANTIDADENOVA] [money] NULL, " & Chr(13)
   sql = sql & "   [USUARIO] [varchar](100) NULL, " & Chr(13)
   sql = sql & "   [NOTA] [varchar](100) NULL, " & Chr(13)
   sql = sql & "   [IP] [varchar](100) NULL, " & Chr(13)
   sql = sql & "   [CENTROCUSTO] [int] NULL, " & Chr(13)
   sql = sql & "   [LOTE] [char](10) NULL, " & Chr(13)
   sql = sql & "   [VALIDADELOTE] [datetime] NULL, " & Chr(13)
   sql = sql & "   [PERIODODISPENSADO] [int] NULL, " & Chr(13)
   sql = sql & "   [ATUALIZACAO] [varchar](155) NULL, " & Chr(13)
   sql = sql & "   [KIT] [int] NULL, " & Chr(13)
   sql = sql & "   [NAOINSERE] [int] NULL, " & Chr(13)
   sql = sql & "   [KITCENTROCUSTO] [int] NULL, " & Chr(13)
   sql = sql & "   [USUARIOLANCADO] [int] NULL, " & Chr(13)
   sql = sql & "   [QUANTIDADEKIT] [int] NULL, " & Chr(13)
   sql = sql & "   [SEQUENCIAINC] [int] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "   [MADRUGADA] [int] NULL, " & Chr(13)
   sql = sql & "   CONSTRAINT [PK_PRE_DEVOLUCAO_ITEM] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "   ( " & Chr(13)
   sql = sql & "      [SEQUENCIAINC] Asc " & Chr(13)
   sql = sql & "   ), " & Chr(13)
   sql = sql & "   CONSTRAINT [FK_PRE_DEVOLUCAO_ITEM_PRE_DEVOLUCAO] FOREIGN KEY(ID_PRE_DEVOLUCAO) " & Chr(13)
   sql = sql & "   REFERENCES PRE_DEVOLUCAO (ID_PRE_DEVOLUCAO)) " & Chr(13)
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD PORTECBHPM INT"
   Banco.Execute sql
      
   If Layout = 30 Then
      sql = ""
      sql = sql & " INSERT INTO MENU("
      sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
      sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
      sql = sql & " SELECT MAX(MENU)+1,'Impressão de Nota Fiscal Avulsa', 'MnuFin_Cor_INF', ' ', 'MnuFin_Cor_INF',"
      sql = sql & "        1, 1, '0902110000', 'MnuFin_Cor_INF', 1"
      sql = sql & " FROM MENU "
      Banco.Execute sql
   End If
      
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[FINOTAFISCALRECEBER]("
   sql = sql & "    [NOTAFISCAL] [decimal](13, 2) NOT NULL,"
   sql = sql & "    [CODIGO] [int] NULL,"
   sql = sql & "    [TIPO] [int] NULL,"
   sql = sql & "    [TIPOPACIENTE] [int] NULL,"
   sql = sql & "    [DATA] [datetime] NULL,"
   sql = sql & "    [VALOR] [money] NULL,"
   sql = sql & "    [NOME] [varchar](155) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [CPF] [varchar](30) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [ENDERECO] [varchar](155) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [CIDADE] [varchar](155) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [UF] [varchar](2) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [CEP] [varchar](50) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [TELEFONE] [varchar](20) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](155) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [IMPRESSOATUALIZACAO] [varchar](155) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [descricao] [varchar](2000) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [IMPRESSO] [int] NULL,"
   sql = sql & "  CONSTRAINT [PK_FINOTAFISCALRECEBER] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & "    [NotaFiscal] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[Brasil3]( " & Chr(13)
   sql = sql & " [CodigoBrasil3] [int] NOT NULL, " & Chr(13)
   sql = sql & " [Descricao] [nvarchar](255) NULL, " & Chr(13)
   sql = sql & " [PortAnes] [tinyint] NULL, " & Chr(13)
   sql = sql & " [PortSala] [tinyint] NULL, " & Chr(13)
   sql = sql & " [Sexo] [tinyint] NULL, " & Chr(13)
   sql = sql & " [IdadeMin] [tinyint] NULL, " & Chr(13)
   sql = sql & " [IdadeMax] [tinyint] NULL, " & Chr(13)
   sql = sql & " [CA] [tinyint] NULL, " & Chr(13)
   sql = sql & " [Filme] [money] NULL, " & Chr(13)
   sql = sql & " [DataAlt] [datetime] NULL, " & Chr(13)
   sql = sql & " [Finci] [smallint] NULL, " & Chr(13)
   sql = sql & " [QtdAux] [smallint] NULL, " & Chr(13)
   sql = sql & " [Carencia] [smallint] NULL, " & Chr(13)
   sql = sql & " [CHantes] [money] NULL, " & Chr(13)
   sql = sql & " [CHdepois] [money] NULL, " & Chr(13)
   sql = sql & " [Espec] [int] NULL, " & Chr(13)
   sql = sql & " [Trava] [varbinary](8) NULL, " & Chr(13)
   sql = sql & " [ANESTESIA] [money] NULL, " & Chr(13)
   sql = sql & " [CODIGOFILME] [int] NULL, " & Chr(13)
   sql = sql & " [CCIHCATEGORIA] [int] NULL, " & Chr(13)
   sql = sql & " [CIRURGIACATEGORIA] [int] NULL, " & Chr(13)
   sql = sql & " [CODIGOCONTRASTE] [int] NULL, " & Chr(13)
   sql = sql & " [ESPECIALIDADECIRURGICA] [int] NULL, " & Chr(13)
   sql = sql & " [CATEGORIAPROCEDIMENTO] [int] NULL, " & Chr(13)
   sql = sql & " [QuantidadeAuxiliar] [int] NULL, " & Chr(13)
   sql = sql & " [FILMEAUX] [money] NULL, " & Chr(13)
   sql = sql & " [PORTEANEST] [int] NULL, " & Chr(13)
   sql = sql & " [NAOMULTIPLICAACOMODACAO] [int] NULL, " & Chr(13)
   sql = sql & " [UTILIZAVIDEO] [int] NULL, " & Chr(13)
   sql = sql & " [GUIAAUTORIZACAO] [int] NULL, " & Chr(13)
   sql = sql & " [filme290507] [money] NULL, " & Chr(13)
   sql = sql & " [PORTSALAAUX] [int] NULL, " & Chr(13)
   sql = sql & " [CONTACONTABIL] [nvarchar](15) NULL, " & Chr(13)
   sql = sql & " [TUSS] [int] NULL, " & Chr(13)
   sql = sql & " [CONTACONTABIL060710] [char](15) NULL, " & Chr(13)
   sql = sql & " [TUSS_CONVENIO] [int] NULL, " & Chr(13)
   sql = sql & " CONSTRAINT [PK_Brasil3] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & " ( " & Chr(13)
   sql = sql & " [CodigoBrasil3] Asc " & Chr(13)
   sql = sql & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] " & Chr(13)
   sql = sql & " ) ON [PRIMARY] " & Chr(13)
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD FONTE_REMUNERACAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT ADD CONTROLECHECAGEM VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_AMB ADD CONTROLECHECAGEM VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_EXT ADD CONTROLECHECAGEM VARCHAR(50)"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD TISS_INSUMO_UTILIZA_FATORFATURAMENTO INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE INSUMOS ADD FATORFATURAMENTO MONEY"
   Banco.Execute sql
   
   sql = " ALTER table TMPAGENDAMENTOCIRURGICOAG ALTER COLUMN PACIENTE1 VARCHAR(100)"
   Banco.Execute sql

   sql = " ALTER TABLE AGENDAMENTOCIRURGICO  ALTER COLUMN PACIENTE VARCHAR(100)"
   Banco.Execute sql
   
   sql = " ALTER TABLE AGENDAMENTOCIRURGICO ALTER COLUMN OBSERVACAO VARCHAR(250)"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPAGENDAMENTOCIRURGICOAG ADD MATERIAL VARCHAR(100)"
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '082012'"
   Banco.Execute sql

   Exit Function
Erro:
   Resume Next
End Function


Public Function AtualizaMes092012()
   On Error GoTo Erro
   
   
   'If Layout = 19 Then
   '   sql = " UPDATE CONVENIOS SET TISS_QUANTIDADECHARREGISTROANS = 8 WHERE CONVENIO = 123 "
   '   Banco.Execute sql
   'End If
   
   sql = " ALTER TABLE CONVENIOPROCEDIMENTO ADD VALOR_SP_CONVENIO MONEY"
   Banco.Execute sql

   sql = " ALTER TABLE CONVENIOPROCEDIMENTO ADD VALOR_SH_CONVENIO MONEY"
   Banco.Execute sql

   sql = " ALTER TABLE CONVENIOPROCEDIMENTO ADD VALOR_SADT_CONVENIO MONEY"
   Banco.Execute sql
   
   sql = "CREATE TABLE TESTEPEZINHOTIPO( " & Chr(13)
   sql = sql & " CODIGO INT IDENTITY, " & Chr(13)
   sql = sql & " DESCRICAO VARCHAR(20), " & Chr(13)
   sql = sql & " ATIVO INT) "
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOPROCEDIMENTO ADD VALOR_SA_CONVENIO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TESTE_CODIGO] [varchar](20) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TESTE_LOTE] [varchar](20) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TESTE_NUMERO] [varchar](20) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TESTE_TIPO] [int] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TESTE_COLETA_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TESTE_RECOLETA_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [DEPART_COLETA_AMOSTRA_RECEB_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [DEPART_COLETA_AMOSTRA_ENVIO_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [DEPART_COLETA_RESULTADO_RECEB_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [DEPART_COLETA_RESULTADO_ENVIO_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [DEPART_RECOLETA_AMOSTRA_RECEB_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [DEPART_RECOLETA_AMOSTRA_ENVIO_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TESTE_COLETA_FUNC] [varchar](50) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TESTE_RECOLETA_FUNC] [varchar](50) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [DEPART_RECOLETA_RESULTADO_RECEB_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [DEPART_RECOLETA_RESULTADO_ENVIO_DATA] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [IDADEGESTACIONALSEMANA] [int] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [IDADEGESTACIONALDIAS] [int] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TRANSFUSAO] [int] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [DATAULTIMATRANSFUSAO] [datetime] NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD [TESTE_RECOLETA_MOTIVO] [varchar](250) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_CODIGO]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_LOTE]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_NUMERO]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_TIPO]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_COLHIDO_DATA]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_ENCAMINHADO_DATA]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_RECEBIDO_DATA]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_ENTREGUE_DATA]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_COLHIDO_HORA]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_ENCAMINHADO_HORA]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_RECEBIDO_HORA]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_ENTREGUE_HORA]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_COLHIDO_FUNC]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_ENCAMINHADO_FUNC]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_RECEBIDO_FUNC]"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS DROP COLUMN [PKU_ENTREGUE_FUNC]"
   Banco.Execute sql
   
   sql = "CREATE TABLE AIH_COD_SOL_LIB ( " & Chr(13)
   sql = sql & "CODIGO INT IDENTITY, " & Chr(13)
   sql = sql & "REGISTRO INT, " & Chr(13)
   sql = sql & "NUMEROAIH VARCHAR(13), " & Chr(13)
   sql = sql & "COD_SOL_LIB INT)"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_AIH_COD_SOL_LIB_REGISTRO] ON AIH_COD_SOL_LIB(REGSITRO) "
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_AIH_COD_SOL_LIB_NUMEROAIH] ON AIH_COD_SOL_LIB(NUMEROAIH) "
   Banco.Execute sql
   
   sql = "ALTER TABLE AIH_COD_SOL_LIB  WITH CHECK ADD  CONSTRAINT [FK_AIH_COD_SOL_LIB_DADOSAIH] FOREIGN KEY([Registro], NUMEROAIH) REFERENCES DADOSAIH([Registro], numeroAIH)"
   Banco.Execute sql
   
   sql = "CREATE TABLE COD_SOL_LIB ( " & Chr(13)
   sql = sql & "COD_SOL_LIB VARCHAR(5) NOT NULL, " & Chr(13)
   sql = sql & "COMBINACOES_COD_SOL_LIB_SAVE VARCHAR(20) CONSTRAINT COMBINACOES_COD_SOL_LIB_SAVE_UNIQUE UNIQUE)"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_COD_SOL_LIB_COMBINACOES_COD_SOL_LIB_SAVE] ON COD_SOL_LIB(COMBINACOES_COD_SOL_LIB_SAVE)"
   Banco.Execute sql
   
   sql = "ALTER TABLE COD_SOL_LIB ADD  CONSTRAINT [PK_COD_SOL_LIB] PRIMARY KEY NONCLUSTERED(COD_SOL_LIB Asc)"
   Banco.Execute sql
   
   sql = "CREATE TABLE COD_SOL_LIB_SAVE ( " & Chr(13)
   sql = sql & "CODIGO INT NOT NULL, " & Chr(13)
   sql = sql & "Descricao VarChar(50))"
   Banco.Execute sql
   
   sql = "ALTER TABLE COD_SOL_LIB_SAVE ADD CONSTRAINT [PK_COD_SOL_LIB_SAVE] PRIMARY KEY NONCLUSTERED (Codigo Asc)"
   Banco.Execute sql
   
   sql = "ALTER TABLE DADOSAIH ADD JUSTIFICATIVALIBERACAO VARCHAR(50)"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[UTICOMBO]("
   sql = sql & " [CODIGO] [int] NOT NULL,"
   sql = sql & " [COMBO] [int] NOT NULL,"
   sql = sql & " [NOME] [varchar](50) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & " [ATUALIZACAO] [varchar](40) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & " CONSTRAINT [PK_UTICOMBO] PRIMARY KEY NONCLUSTERED"
   sql = sql & " ("
   sql = sql & "    [CODIGO] ASC,"
   sql = sql & "    [Combo] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[UTIPRINCIPAL]("
   sql = sql & "    [REGISTRO] [int] NOT NULL,"
   sql = sql & "    [DATA] [datetime] NOT NULL,"
   sql = sql & "    [CABECA] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_CABECA]  DEFAULT (0),"
   sql = sql & "    [PESCOCO] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_PESCOCO]  DEFAULT (0),"
   sql = sql & "    [MAOE] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_MAOE]  DEFAULT (0),"
   sql = sql & "    [MAOD] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_MAOD]  DEFAULT (0),"
   sql = sql & "    [ANTEBRACOE] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_ANTEBRACOE]  DEFAULT (0),"
   sql = sql & "    [ANTEBRACOD] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_ANTEBRACOD]  DEFAULT (0),"
   sql = sql & "    [BRACOE] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_BRACOE]  DEFAULT (0),"
   sql = sql & "    [BRACOD] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_BRACOD]  DEFAULT (0),"
   sql = sql & "    [TORAX] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_TORAX]  DEFAULT (0),"
   sql = sql & "    [ABDOMEM] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_ABDOMEM]  DEFAULT (0),"
   sql = sql & "    [BACIA] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_BACIA]  DEFAULT (0),"
   sql = sql & "    [COXAE] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_COXAE]  DEFAULT (0),"
   sql = sql & "    [COXAD] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_COXAD]  DEFAULT (0),"
   sql = sql & "    [JOELHOE] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_JOELHOE]  DEFAULT (0),"
   sql = sql & "    [JOELHOD] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_JOELHOD]  DEFAULT (0),"
   sql = sql & "    [PANTURRILHAE] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_PANTURRILHAE]  DEFAULT (0),"
   sql = sql & "    [PANTURRILHAD] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_PANTURRILHAD]  DEFAULT (0),"
   sql = sql & "    [PEE] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_PEE]  DEFAULT (0),"
   sql = sql & "    [PED] [int] NULL CONSTRAINT [DF_UTIPRINCIPAL_PED]  DEFAULT (0),"
   sql = sql & "    [ATUALIZACAO] [varchar](40) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "  CONSTRAINT [PK_UTIPRINCIPAL] PRIMARY KEY NONCLUSTERED"
   sql = sql & " ("
   sql = sql & "    [REGISTRO] ASC,"
   sql = sql & "    [data] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql


   sql = ""
   sql = sql & " CREATE TABLE [dbo].[UTIMEDICOPACIENTE]("
   sql = sql & "    [REGISTRO] [int] NULL,"
   sql = sql & "    [MEDICO] [int] NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](40) COLLATE Latin1_General_CI_AS NULL"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & " CREATE TABLE [dbo].[UTIEVOLUCAOCLINICA]("
   sql = sql & "    [REGISTRO] [int] NOT NULL,"
   sql = sql & "    [DATA] [datetime] NOT NULL,"
   sql = sql & "    [CONSCIENCIA] [int] NULL,"
   sql = sql & "    [MUCOSA] [int] NULL,"
   sql = sql & "    [VIAAEREA] [int] NULL,"
   sql = sql & "    [TRACAO] [int] NULL,"
   sql = sql & "    [PERIFERIA] [int] NULL,"
   sql = sql & "    [IMOBILIZACAO] [int] NULL,"
   sql = sql & "    [DRENOTORAX] [int] NULL,"
   sql = sql & "    [FIXADORORTOPEDICO] [int] NULL,"
   sql = sql & "    [CHOQUE] [int] NULL,"
   sql = sql & "    [PUPILA] [int] NULL,"
   sql = sql & "    [PAMINIMA] [int] NULL,"
   sql = sql & "    [PAMEDIA] [int] NULL,"
   sql = sql & "    [PAMAXIMA] [int] NULL,"
   sql = sql & "    [TEMPERATURA] [int] NULL,"
   sql = sql & "    [PULSOFC] [int] NULL,"
   sql = sql & "    [FREQUENCIARESPIRATORIA] [int] NULL,"
   sql = sql & "    [PESO] [int] NULL,"
   sql = sql & "    [DIURESE] [int] NULL,"
   sql = sql & "    [EVOLUCAO] [varchar](1024) COLLATE Latin1_General_CI_AS NULL,"
   sql = sql & "    [ABREOLHO] [int] NULL,"
   sql = sql & "    [RESPIRACAOVERBAL] [int] NULL,"
   sql = sql & "    [RESPIRACAOMOTORA] [int] NULL,"
   sql = sql & "    [ACESSOVENOSOCENTRAL] [int] NULL,"
   sql = sql & "    [SONDANASOGASTRICA] [int] NULL,"
   sql = sql & "    [SONDAOROGASTRICA] [int] NULL,"
   sql = sql & "    [SONDAVESICAL] [int] NULL,"
   sql = sql & "    [SONDAENTERAL] [int] NULL,"
   sql = sql & "    [TUBOORONASOTRAQUEAL] [int] NULL,"
   sql = sql & "    [CATETERPULMONAR] [int] NULL,"
   sql = sql & "    [CATETERHEMODIALISE] [int] NULL,"
   sql = sql & "    [CATETERDIALISEPERITONEAL] [int] NULL,"
   sql = sql & "    [CATETERPAMEDIA] [int] NULL,"
   sql = sql & "    [GASTROSTOMIA] [int] NULL,"
   sql = sql & "    [JEJUNOSTOMIA] [int] NULL,"
   sql = sql & "    [CISTOSTOMIA] [int] NULL,"
   sql = sql & "    [NEFROSTOMIA] [int] NULL,"
   sql = sql & "    [CANULATRAQUEAL] [int] NULL,"
   sql = sql & "    [GUEDEL] [int] NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](40) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "    [FLEBOTOMIA] [int] NULL CONSTRAINT [DF__UTIEVOLUC__FLEBO__1D48B800]  DEFAULT (0),"
   sql = sql & "    [INTRACATH] [int] NULL CONSTRAINT [DF__UTIEVOLUC__INTRA__1E3CDC39]  DEFAULT (0),"
   sql = sql & "    [ENTUBACAO] [int] NULL CONSTRAINT [DF__UTIEVOLUC__ENTUB__1F310072]  DEFAULT (0),"
   sql = sql & "    [EXTUBACAO] [int] NULL CONSTRAINT [DF__UTIEVOLUC__EXTUB__202524AB]  DEFAULT (0),"
   sql = sql & "    [TRAQUEOSTOMIA] [int] NULL CONSTRAINT [DF__UTIEVOLUC__TRAQU__211948E4]  DEFAULT (0),"
   sql = sql & "    [SONDAGASRTRICA] [int] NULL CONSTRAINT [DF__UTIEVOLUC__SONDA__220D6D1D]  DEFAULT (0),"
   sql = sql & "    [NPP] [int] NULL CONSTRAINT [DF__UTIEVOLUCAO__NPP__23019156]  DEFAULT (0),"
   sql = sql & "  CONSTRAINT [PK_UTIEVOLUCAOCLINICA] PRIMARY KEY NONCLUSTERED"
   sql = sql & " ("
   sql = sql & "    [REGISTRO] ASC,"
   sql = sql & "    [data] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql

   sql = ""
   sql = sql & " CREATE TABLE [dbo].[UTICUIDADOGERAL]("
   sql = sql & "    [REGISTRO] [int] NOT NULL,"
   sql = sql & "    [DATA] [datetime] NOT NULL,"
   sql = sql & "    [EQUIPAMENTOVENTILACAO] [int] NULL,"
   sql = sql & "    [MANTERSNG] [int] NULL,"
   sql = sql & "    [SINALVITAL] [int] NULL,"
   sql = sql & "    [MODO] [int] NULL,"
   sql = sql & "    [DRENO] [int] NULL,"
   sql = sql & "    [CANULATRAQUEAL] [int] NULL,"
   sql = sql & "    [FISIOTERAPIA] [int] NULL,"
   sql = sql & "    [VEZESDIAFISIO] [int] NULL,"
   sql = sql & "    [CABECEIRALEITO] [int] NULL,"
   sql = sql & "    [VIAORAL] [int] NULL,"
   sql = sql & "    [TIPO] [int] NULL,"
   sql = sql & "    [ENTERAL] [int] NULL,"
   sql = sql & "    [VOLUME] [int] NULL,"
   sql = sql & "    [VEZESDIANUT] [int] NULL,"
   sql = sql & "    [PARENTERAL] [int] NULL,"
   sql = sql & "    [APARTIRDAS] [int] NULL,"
   sql = sql & "    [MONITORIZACAOCARDIACA] [int] NULL,"
   sql = sql & "    [BOMBAINFUSORA] [int] NULL,"
   sql = sql & "    [O2] [int] NULL,"
   sql = sql & "    [VENTILACAO] [int] NULL,"
   sql = sql & "    [OXIMETRIA] [int] NULL,"
   sql = sql & "    [DIETAENTERAL] [int] NULL,"
   sql = sql & "    [DIETAPARENTERAL] [int] NULL,"
   sql = sql & "    [CAPNOGRAFO] [int] NULL,"
   sql = sql & "    [MEDIDAPIC] [int] NULL,"
   sql = sql & "    [BALAOINTRAAORTICO] [int] NULL,"
   sql = sql & "    [DIALISEPERITONEAL] [int] NULL,"
   sql = sql & "    [HEMODIALISE] [int] NULL,"
   sql = sql & "    [ULTRAFILTRO] [int] NULL,"
   sql = sql & "    [PAINVASIVA] [int] NULL,"
   sql = sql & "    [PRESSAOINTRAPULMONAR] [int] NULL,"
   sql = sql & "    [ATUALIZACAO] [varchar](40) COLLATE Latin1_General_CI_AS NOT NULL,"
   sql = sql & "  CONSTRAINT [PK_UTICUIDADOGERAL] PRIMARY KEY NONCLUSTERED"
   sql = sql & " ("
   sql = sql & "    [REGISTRO] ASC,"
   sql = sql & "    [data] Asc"
   sql = sql & " )WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]"
   sql = sql & " ) ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[CIRURGIA_INSUMO]( " & Chr(13)
   sql = sql & "  [CIRURGIA] [int] NOT NULL, " & Chr(13)
   sql = sql & "  CONVENIO SMALLINT NOT NULL, " & Chr(13)
   sql = sql & "  INSUMO INT NOT NULL, " & Chr(13)
   sql = sql & "  [ATUALIZACAO] [varchar](60) NULL, " & Chr(13)
   sql = sql & "  CONSTRAINT [PK_CIRURGIA_INSUMO] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "  [CIRURGIA], " & Chr(13)
   sql = sql & "  [CONVENIO], " & Chr(13)
   sql = sql & "  [Insumo]))"
   Banco.Execute sql
   
   sql = "ALTER TABLE [CIRURGIA_INSUMO]  WITH CHECK ADD  CONSTRAINT [FK_CIRURGIA_INSUMO_INSUMO] FOREIGN KEY(INSUMO, CONVENIO) REFERENCES INSUMOS(INSUMO, CONVENIO)"
   Banco.Execute sql
   
   sql = "ALTER TABLE [CIRURGIA_INSUMO]  WITH CHECK ADD  CONSTRAINT [FK_CIRURGIA_INSUMO_CIRURGIA] FOREIGN KEY(CIRURGIA) REFERENCES CIRURGIACADASTRO (CIRURGIA)"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[CIRURGIAINSUMO]( " & Chr(13)
   sql = sql & "  [REGISTRO] [int] NOT NULL, " & Chr(13)
   sql = sql & "  [INSUMO] [int] NOT NULL, " & Chr(13)
   sql = sql & "  [DOCUMENTO] [int] NOT NULL, " & Chr(13)
   sql = sql & "  [NOME] [varchar](50) NULL, " & Chr(13)
   sql = sql & "  [ATUALIZACAO] [varchar](40) NOT NULL, " & Chr(13)
   sql = sql & "  [TIPOREGISTRO] [int] NULL, " & Chr(13)
   sql = sql & "  [SEQUENCIA] [int] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "  [QUANTIDADEFATURADA] [int] NULL, " & Chr(13)
   sql = sql & "  [SEQUENCIAFATURAMENTO] [int] NULL, " & Chr(13)
   sql = sql & "  CONSTRAINT [PK_CIRURGIAINSUMO] PRIMARY KEY NONCLUSTERED " & Chr(13)
   sql = sql & "( " & Chr(13)
   sql = sql & "  [REGISTRO] ASC, " & Chr(13)
   sql = sql & "  [INSUMO] ASC, " & Chr(13)
   sql = sql & "  [Documento] Asc))"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_CIRURGIAINSUMO_REGISTRO] ON [dbo].CIRURGIAINSUMO (Registro) "
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_CIRURGIAINSUMO_DOCUMENTO] ON [dbo].CIRURGIAINSUMO (Documento)"
   Banco.Execute sql
   
   sql = "CREATE NONCLUSTERED INDEX [IX_CIRURGIAINSUMO_SEQUENCIA] ON [dbo].CIRURGIAINSUMO (Sequencia)"
   Banco.Execute sql
   
   'ADICIONA O MENU LOCAL DE CONSUMO
   sql = ""
   sql = "INSERT INTO MENU VALUES (" & Chr(13)
   sql = sql & "1147, 'Locais de Consumo', 'MnuCGLocalC', '', 'MnuPar_Cad_Ger_Lco', 1,1,'0101012100', 'MnuPar_Cad_Ger_Lco',3,NULL,NULL)" & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE CENTROCUSTO ADD UPA INT NULL DEFAULT 0"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO DROP COLUMN CAMINHOSOLICITACAOELEGIBILIDADE"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO DROP COLUMN CAMINHORESPOSTAELEGIBILIDADE"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO DROP COLUMN CAMINHOSOLICITACAOPROCEDIMENTO"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO DROP COLUMN CAMINHORESPOSTAPROCEDIMENTO"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO DROP COLUMN CAMINHOSOLICITACAODADOSBENE"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO DROP COLUMN CAMINHORESPOSTADADOSBENE"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD REALIZAELEGIBILIDADE INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD ELEGIBILIDADEVIA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPENTRADAHOSTNAME VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPENTRADAUSUARIO VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPENTRADAPASSWORD VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPENTRADAPORTA VARCHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPENTRADAPASTA VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPSAIDAHOSTNAME VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPSAIDAUSUARIO VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPSAIDAPASSWORD VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPSAIDAPORTA VARCHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD FTPSAIDAPASTA VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD INTRANETENTRADA VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD INTRANETSAIDA VARCHAR(250)"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE CENTROCUSTO ADD LOGO IMAGE NULL"
   Banco.Execute sql

   sql = " ALTER TABLE CONVENIOPROCEDIMENTOHONORARIO ADD VALOR_SH_CONVENIO MONEY"
   Banco.Execute sql

   sql = " ALTER TABLE CONVENIOPROCEDIMENTOHONORARIO ADD VALOR_SADT_CONVENIO MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOPROCEDIMENTOHONORARIO ADD VALOR_SA_CONVENIO MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOPROCEDIMENTOHONORARIO ADD VALOR_SP_CONVENIO MONEY"
   Banco.Execute sql

   sql = ""
   sql = "ALTER TABLE CENTROCUSTO ADD UTILIZARLOGO INT NULL DEFAULT 0"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE AGENDAMENTOCIRURGICO ADD UNIDADECARTEIRINHA VARCHAR(10) NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE CENTROCUSTO ADD CC_CIRURGICO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE FICHAS ADD POSSUI_PRONTUARIO INT"
   Banco.Execute sql
   
   sql = "CREATE TABLE TMPINDICADORESANALITICO( " & Chr(13)
   sql = sql & "INDICADOR INT, " & Chr(13)
   sql = sql & "INDICADORDESCRICAO VARCHAR(50), " & Chr(13)
   sql = sql & "TIPOREGISTRO VARCHAR(20), " & Chr(13)
   sql = sql & "REGISTRO INT, " & Chr(13)
   sql = sql & "NOME VARCHAR(150), " & Chr(13)
   sql = sql & "DATA DATETIME, " & Chr(13)
   sql = sql & "QUANTIDADE MONEY, " & Chr(13)
   sql = sql & "TIPOCONVENIOCODIGO INT, " & Chr(13)
   sql = sql & "TIPOCONVENIO VARCHAR(20), " & Chr(13)
   sql = sql & "IP VARCHAR(100))"
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '092012'"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes102012()
   On Error GoTo Erro
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN QUANTIDADE MONEY "
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_EMPRESA VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_CNES INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_CGC VARCHAR(15)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_INSCRICAOESTADUAL VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_INSCRICAOMUNICIPAL VARCHAR(15)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_CPFDIRETOR VARCHAR(15)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_CNSDIRETOR VARCHAR(15)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_CRMDIRETOR VARCHAR(15)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_RUA VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_NUMERO VARCHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_BAIRRO VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_CEP VARCHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_CIDADE VARCHAR(50)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_UF VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_NOMEADMINISTRADOR1 VARCHAR(100)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PRESTADOR_NOMEADMINISTRADOR2 VARCHAR(100)"
   Banco.Execute sql
   
   sql = "UPDATE MENU " & Chr(13)
   sql = sql & "SET NOMECAPTION = 'Integração com CIH e CIHA' " & Chr(13)
   sql = sql & "WHERE NOMECAPTION = 'integração com cih'"
   Banco.Execute sql
   
   sql = "ALTER VIEW V_RECUPERA_CIHA_INSUMO " & Chr(13)
   sql = sql & "AS " & Chr(13)
   sql = sql & "    SELECT 5 AS TIPOREG, 'AMB' AS TIPO, A.DATAINTERNACAO,  " & Chr(13)
   sql = sql & "         MAX(E.CODIGOSUS) AS PROCSUS, SUM(C.QUANTIDADE) AS QUANTIDADE, C.REGISTRO, C.SEQUENCIA,  " & Chr(13)
   sql = sql & "         ISNULL(D.FONTE_REMUNERACAO,0) AS FONTE_REMUNERACAO, " & Chr(13)
   sql = sql & "         ISNULL(D.REGISTROANS,'') AS REGISTROANS " & Chr(13)
   sql = sql & "    FROM   AMBULATORIAL A WITH(NOLOCK) INNER JOIN FICHAS       B WITH(NOLOCK) ON A.FICHA    = B.FICHA " & Chr(13)
   sql = sql & "                              INNER JOIN MOVIM_AMB    C WITH(NOLOCK) ON A.REGISTRO = C.REGISTRO " & Chr(13)
   sql = sql & "                                                           AND C.TIPOLANCAMENTO IN (4)     " & Chr(13)
   sql = sql & "                              INNER JOIN CONVENIOS    D WITH(NOLOCK)               ON A.CONVENIO = D.CONVENIO " & Chr(13)
   sql = sql & "                              INNER JOIN SUS_TABELA_CIH_INSUMO E ON C.PROCEDIMENTO = E.INSUMO " & Chr(13)
   sql = sql & "                                                        AND A.CONVENIO   = E.CONVENIO " & Chr(13)
   sql = sql & "   WHERE 1 = 1  " & Chr(13)
   sql = sql & "   AND    D.TIPOCONVENIO <> 3 " & Chr(13)
   sql = sql & "   AND    ISNULL(A.CANCELADO,0) = 0  " & Chr(13)
   sql = sql & "   AND    A.REGISTRO>0  " & Chr(13)
   sql = sql & "   AND EXISTS (SELECT CON.PROCEDIMENTO  " & Chr(13)
   sql = sql & "               FROM SUS_PROCEDIMENTO_REGISTRO CON  " & Chr(13)
   sql = sql & "               WHERE CON.REGISTRO = 1   " & Chr(13)
   sql = sql & "               AND   CON.PROCEDIMENTO = E.CODIGOSUS) " & Chr(13)
   sql = sql & "  GROUP BY A.DATAINTERNACAO, C.REGISTRO, C.SEQUENCIA, D.FONTE_REMUNERACAO, D.REGISTROANS " & Chr(13)
   sql = sql & " " & Chr(13)
   sql = sql & "    UNION ALL " & Chr(13)
   sql = sql & " " & Chr(13)
   sql = sql & "  SELECT 5 AS TIPOREG, 'AMB' AS TIPO, A.DATAINTERNACAO,  " & Chr(13)
   sql = sql & "        MAX(E.CODIGOSUS) AS PROCSUS,  SUM(C.QUANTIDADE) AS QUANTIDADE, C.REGISTRO, C.SEQUENCIA, " & Chr(13)
   sql = sql & "        ISNULL(D.FONTE_REMUNERACAO,0) AS FONTE_REMUNERACAO, " & Chr(13)
   sql = sql & "        ISNULL(D.REGISTROANS,'') AS REGISTROANS " & Chr(13)
   sql = sql & "  FROM   EXTERNO      A WITH(NOLOCK) INNER JOIN FICHAS    B WITH(NOLOCK) ON A.FICHA = B.FICHA " & Chr(13)
   sql = sql & "                             INNER JOIN MOVIM_EXT C WITH(NOLOCK) ON A.REGISTRO = C.REGISTRO " & Chr(13)
   sql = sql & "                                                           AND C.TIPOLANCAMENTO IN (4)     " & Chr(13)
   sql = sql & "                             INNER JOIN CONVENIOS D WITH(NOLOCK)               ON A.CONVENIO = D.CONVENIO " & Chr(13)
   sql = sql & "                              INNER JOIN SUS_TABELA_CIH_INSUMO E ON C.PROCEDIMENTO = E.INSUMO " & Chr(13)
   sql = sql & "                                                        AND A.CONVENIO   = E.CONVENIO " & Chr(13)
   sql = sql & "  WHERE 1 = 1  " & Chr(13)
   sql = sql & "  AND    D.TIPOCONVENIO <> 3 " & Chr(13)
   sql = sql & "  AND    ISNULL(A.CANCELADO,0) = 0  " & Chr(13)
   sql = sql & "  AND    A.REGISTRO>0  " & Chr(13)
   sql = sql & "  AND    ISNULL(A.PACIENTE_INTERNO_FATURADO,0) = 0  " & Chr(13)
   sql = sql & "  AND EXISTS (SELECT CON.PROCEDIMENTO  " & Chr(13)
   sql = sql & "            FROM SUS_PROCEDIMENTO_REGISTRO CON  " & Chr(13)
   sql = sql & "            WHERE CON.REGISTRO = 1   -- CONSOLIDADO " & Chr(13)
   sql = sql & "            AND   CON.PROCEDIMENTO = E.CODIGOSUS) " & Chr(13)
   sql = sql & "  GROUP BY A.DATAINTERNACAO, C.REGISTRO, C.SEQUENCIA, D.FONTE_REMUNERACAO, D.REGISTROANS "
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPINDICADOR ADD ANTERIOR MONEY"
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '102012'"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNOPROCEDIMENTO" & Chr(13)
   sql = sql & "ALTER COLUMN NOME VARCHAR(200) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CENTROCUSTO" & Chr(13)
   sql = sql & "ADD    ENDERECO VARCHAR(100) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CENTROCUSTO" & Chr(13)
   sql = sql & "ADD    NUMERO VARCHAR(10) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CENTROCUSTO" & Chr(13)
   sql = sql & "ADD    CEP VARCHAR(10) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CENTROCUSTO" & Chr(13)
   sql = sql & "ADD    CIDADE INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CENTROCUSTO" & Chr(13)
   sql = sql & "ADD    CNPJ VARCHAR(20) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CENTROCUSTO" & Chr(13)
   sql = sql & "ADD    INSCRICAOESTADUAL VARCHAR(20) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CENTROCUSTO" & Chr(13)
   sql = sql & "ADD    BAIRRO VARCHAR(50) NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE PARAMETRO" & Chr(13)
   sql = sql & "ADD FATURAMENTO_UTILIZADATARECEBIMENTO INT NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE REQUISICAOPRODUTO" & Chr(13)
   sql = sql & "ADD SELECIONADO INT DEFAULT 0 NULL"
   Banco.Execute sql
   
   sql = "ALTER TRIGGER TR_PRODUTO_LOTE_ETIQUETA " & Chr(13)
   sql = sql & "ON PRODUTOCENTROCUSTOLOTE " & Chr(13)
   sql = sql & "With ENCRYPTION " & Chr(13)
   sql = sql & "FOR INSERT " & Chr(13)
   sql = sql & "AS " & Chr(13)
   sql = sql & "   INSERT INTO PRODUTO_LOTE_ETIQUETA ( " & Chr(13)
   sql = sql & "          PRODUTO, LOTE, VALIDADELOTE ) " & Chr(13)
   sql = sql & "   SELECT DISTINCT PRODUTO, A.LOTE, A.VALIDADE " & Chr(13)
   sql = sql & "   FROM PRODUTOCENTROCUSTOLOTE A " & Chr(13)
   sql = sql & "   WHERE A.PRODUTO NOT IN  (SELECT PRODUTO " & Chr(13)
   sql = sql & "             FROM PRODUTO_LOTE_ETIQUETA B " & Chr(13)
   sql = sql & "             Where a.Produto = b.Produto " & Chr(13)
   sql = sql & "             AND   A.LOTE     = B.LOTE " & Chr(13)
   sql = sql & "             AND   A.VALIDADE = B.VALIDADELOTE)"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[TMPESTATISTICAALTASANO]( " & Chr(13)
   sql = sql & "  [REGISTRO] [int] NULL, " & Chr(13)
   sql = sql & "  [NOME] [varchar](200) NULL, " & Chr(13)
   sql = sql & "  [DATAINTERNACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "  [DATAALTA] [datetime] NULL, " & Chr(13)
   sql = sql & "  [UNIDADE] [int] NULL, " & Chr(13)
   sql = sql & "  [UNIDADEDESC] [varchar](200) NULL, " & Chr(13)
   sql = sql & "  [LEITO] [varchar](10) NULL, " & Chr(13)
   sql = sql & "  [IP] [varchar](200) NULL " & Chr(13)
   sql = sql & ") ON [PRIMARY] " & Chr(13)
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[TMPESTATISTICAINTERNACOESANO]( " & Chr(13)
   sql = sql & "  [REGISTRO] [int] NULL, " & Chr(13)
   sql = sql & "  [NOME] [varchar](200) NULL, " & Chr(13)
   sql = sql & "  [DATAINTERNACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "  [DATAALTA] [datetime] NULL, " & Chr(13)
   sql = sql & "  [UNIDADE] [int] NULL, " & Chr(13)
   sql = sql & "  [UNIDADEDESC] [varchar](200) NULL, " & Chr(13)
   sql = sql & "  [LEITO] [varchar](10) NULL, " & Chr(13)
   sql = sql & "  [IP] [varchar](200) NULL " & Chr(13)
   sql = sql & ") ON [PRIMARY] " & Chr(13)
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[TMPESTATISTICAPACIENTESDIA]( " & Chr(13)
   sql = sql & "  [REGISTRO] [int] NULL, " & Chr(13)
   sql = sql & "  [NOME] [varchar](200) NULL, " & Chr(13)
   sql = sql & "  [DATAINTERNACAO] [datetime] NULL, " & Chr(13)
   sql = sql & "  [DATAALTA] [datetime] NULL, " & Chr(13)
   sql = sql & "  [UNIDADE] [int] NULL, " & Chr(13)
   sql = sql & "  [UNIDADEDESC] [varchar](200) NULL, " & Chr(13)
   sql = sql & "  [LEITO] [varchar](10) NULL, " & Chr(13)
   sql = sql & "  [DATA] [datetime] NULL, " & Chr(13)
   sql = sql & "  [IP] [varchar](200) NULL " & Chr(13)
   sql = sql & ") ON [PRIMARY]"
   Banco.Execute sql
   
   sql = "CREATE TABLE CNES_SERVICO_CLASSIF(" & Chr(13)
   sql = sql & "CNES INT NOT NULL," & Chr(13)
   sql = sql & "SERVICO INT NOT NULL," & Chr(13)
   sql = sql & "CLASSIFICACAO INT NOT NULL," & Chr(13)
   sql = sql & "ATUALIZACAO VARCHAR(100)," & Chr(13)
   sql = sql & "CONSTRAINT [PK_CNES_SERVICO_CLASSIF] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "(" & Chr(13)
   sql = sql & "  CNES ASC," & Chr(13)
   sql = sql & "  SERVICO ASC," & Chr(13)
   sql = sql & "  CLASSIFICACAO" & Chr(13)
   sql = sql & "))"
   Banco.Execute sql
   
   '---------------------------------------------------------------------
   'Protocolo: 11226
   'Ativar menu Dispensação de Materiais/ Medicamentos
   'para o cliente Salto de Pirapora
   If Layout = 42 Then
      sql = ""
      sql = "UPDATE MENU SET MENU.ATIVADO = 1 " & Chr(13)
      sql = sql & " WHERE MENU.MENU = 338"
      Banco.Execute sql
   End If
   Exit Function
Erro:
   Resume Next
End Function
 
Public Function AtualizaMes112012()
On Error GoTo Erro
   
   sql = "ALTER TABLE AGENDAMENTOCIRURGICO" & Chr(13)
   sql = sql & "ADD UNIDADE VARCHAR(5) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE AGENDAMENTOCIRURGICO" & Chr(13)
   sql = sql & "ADD CARTEIRINHA VARCHAR(20) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE AGENDAMENTOCIRURGICO" & Chr(13)
   sql = sql & "ADD CARTEIRINHACOMP VARCHAR(50) NULL"
   Banco.Execute sql
   
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '112012'"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE CENTROCUSTO ADD TIPOCENTROCUSTO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD CNES VARCHAR(20)"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE INSUMOS_LOG ALTER COLUMN DESCRICAO VARCHAR(155) NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE INSUMOS_LOG ALTER COLUMN DESCRICAOANTIGA VARCHAR(155) NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE OUTRO_DEPENDENTE ADD FICHA_AUX INT NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE LAU_MOVIM_EXT ADD TROUXEMATERIAL INT NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE LAU_MOVIM_AMB ADD TROUXEMATERIAL INT NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE LAU_MOVIM_INT ADD TROUXEMATERIAL INT NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "ALTER TABLE LAU_MOVIM_DET_SERV ADD TROUXEMATERIAL INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD NUMERODECLARACAO BIGINT NULL"
   Banco.Execute sql
   
'   sql = "USE [Save]" & Chr(13)
'   sql = sql & "GO" & Chr(13)
'   sql = sql & "SET ANSI_NULLS ON" & Chr(13)
'   sql = sql & "GO" & Chr(13)
'   sql = sql & "SET QUOTED_IDENTIFIER ON" & Chr(13)
'   sql = sql & "GO" & Chr(13)
'   sql = sql & "CREATE VIEW [dbo].[V_RECUPERAFRETECOMPRA]" & Chr(13)
'   sql = sql & "AS SELECT   1 AS TIPO ,SUM(ISNULL(A.PRECOBRASINDICE*(A.QUANTIDADE-ISNULL(B.QUANTIDADE,0)),0)" & Chr(13)
'   sql = sql & ") AS VALORTOTAL, AVG(ISNULL(A.VALORFRETE,0))+AVG(ISNULL(A.OUTRADESPESA,0))-AVG(ISNULL(A.VALORDESCONTO,0)) AS VALORDESPESADESCONTO," & Chr(13)
'   sql = sql & "A.FORNECEDOR,A.DOCUMENTO,A.DATA ,A.CCORIGEM,0 AS PEDIDO" & Chr(13)
'   sql = sql & "FROM MOVIMENTOAVULSO A LEFT  JOIN MOVIMENTOAVULSO B     ON A.PRODUTO=B.PRODUTO" & Chr(13)
'   sql = sql & "AND A.CCORIGEM=B.CCORIGEM AND A.FORNECEDOR=B.FORNECEDOR AND A.DOCUMENTO=B.DOCUMENTO" & Chr(13)
'   sql = sql & "AND A.DATA=B.DATA AND B.TIPO=1" & Chr(13)
'   sql = sql & "WHERE A.TIPO=0" & Chr(13)
'   sql = sql & "GROUP BY  A.FORNECEDOR,A.DOCUMENTO,A.DATA,A.CCORIGEM" & Chr(13)
'   sql = sql & "UNION ALL" & Chr(13)
'   sql = sql & "SELECT   2 AS TIPO ,SUM(ISNULL(B.QUANTIDADE*B.VALORUNITARIO,0)) AS VALORTOTAL," & Chr(13)
'   sql = sql & "AVG(ISNULL(A.FRETE,0))+AVG(ISNULL(A.OUTRADESPESA,0))-AVG(ISNULL(A.DESCONTO,0)) AS VALORDESPESADESCONTO," & Chr(13)
'   sql = sql & "C.FABRICANTE,A.NOTAFISCAL,B.DATAENTREGA ,0,A.PEDIDOCOMPRA" & Chr(13)
'   sql = sql & "FROM PEDIDOCOMPRANOTA1 A INNER JOIN PEDIDOCOMPRAPRODUTONOTA1 B ON A.PEDIDOCOMPRA=B.PEDIDOCOMPRA" & Chr(13)
'   sql = sql & "AND A.NOTAFISCAL=B.NOTAFISCAL" & Chr(13)
'   sql = sql & "INNER JOIN PEDIDOCOMPRA1 C            ON A.PEDIDOCOMPRA=C.PEDIDOCOMPRA" & Chr(13)
'   sql = sql & "GROUP BY  C.FABRICANTE,A.NOTAFISCAL,B.DATAENTREGA,A.PEDIDOCOMPRA" & Chr(13)
'   sql = sql & "GO" & Chr(13)
'   Banco.Execute sql

   sql = "INSERT INTO MENU (MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & "SELECT MAX(MENU)+1,'Média por Período','MnuFSugestaoCompra2','','mnuCom_Sco_Com',1,2,'0503010000','mnuCom_Sco_Com',1 FROM MENU" & Chr(13)
   Banco.Execute sql
   
   sql = "INSERT INTO MENU (MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & "SELECT MAX(MENU)+1,'Estoque Mínimo','MnuFSugestaoCompra3','','mnuCom_Sco_Min',1,2,'0503020000','mnuCom_Sco_Min',1 FROM MENU" & Chr(13)
   Banco.Execute sql
   
   sql = "UPDATE MENU SET NOMECAPTION = 'Sugestão de Compras' WHERE MENU = 674 AND MODULO = 2" & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE CIR_PACIENTE_CIRURGIA   ADD BOX INT NULL"
   Banco.Execute sql
   
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes122012()
On Error GoTo Erro
   
   sql = "INSERT INTO MENU (MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & "SELECT MAX(MENU)+1,'Média por Período','MnuFSugestaoCompra2','','mnuCom_Sco_Com',1,2,'0503010000','mnuCom_Sco_Com',1 FROM MENU" & Chr(13)
   Banco.Execute sql
   
   sql = "INSERT INTO MENU (MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,NOMESUBAUX,NIVELVISIBILIDADE)" & Chr(13)
   sql = sql & "SELECT MAX(MENU)+1,'Estoque Mínimo','MnuFSugestaoCompra3','','mnuCom_Sco_Min',1,2,'0503020000','mnuCom_Sco_Min',1 FROM MENU" & Chr(13)
   Banco.Execute sql
   
   sql = "UPDATE MENU SET NOMECAPTION = 'Sugestão de Compras' WHERE MENU = 674 AND MODULO = 2" & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE CIR_PACIENTE_CIRURGIA   ADD BOX INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE1 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE2 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE3 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE4 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE5 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE6 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE7 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE8 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE9 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE10 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE11 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE12 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE13 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE14 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE15 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE16 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE17 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE18 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE19 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE20 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE21 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE22 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE23 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE24 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE7_2 MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ALTER COLUMN HORAQTDE8_2 MONEY"
   Banco.Execute sql
      
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '122012'"
   Banco.Execute sql
   
   sql = " ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT" & Chr(13)
   sql = sql & " ALTER COLUMN CONTROLECHECAGEM VARCHAR(155) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CIRURGIA ADD COMPLICACAO VARCHAR(255) NULL"
   Banco.Execute sql
   
   sql = ""
   sql = "CREATE TABLE [dbo].[MATERIAL_AGENDAMENTO](" & Chr(13)
   sql = sql & "[SEQUENCIA] [bigint] NOT NULL," & Chr(13)
   sql = sql & "[AGENDAMENTO] [bigint] NOT NULL," & Chr(13)
   sql = sql & "[CODIGOEQUIPAMENTO] [int] NULL," & Chr(13)
   sql = sql & "[PROPRIETARIO] [int] NULL," & Chr(13)
   sql = sql & "[PRODUTO] [int] NULL," & Chr(13)
   sql = sql & "[TIPO] [varchar](20) NOT NULL," & Chr(13)
   sql = sql & "CONSTRAINT [PK_SEQUENCIA] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "([SEQUENCIA] ASC, [AGENDAMENTO] ASC)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & Chr(13)
   sql = sql & ") ON [PRIMARY]" & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPAGENDAMENTOCIRURGICOAG ADD AGENDAMENTO BIGINT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPAGENDAMENTOCIRURGICOAG ADD CARTEIRINHA INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TER_PARAMETRO ADD EVENTO_IRRF_OUTROS INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TER_PARAMETRO ADD EVENTO_INSS_JURIDICO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PSICOTROPICOMOVIMENTO ADD TIPOLIVRO VARCHAR(30) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PEDIDOCOMPRA1 ADD ALTERACAOFORNECEDOR INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE SAME_ESTATISTICA_INTERNACAO ADD TIPOCONVENIO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE SAME_ESTATISTICA_INTERNACAO ADD CONVENIO INT NULL"
   Banco.Execute sql
   
   If Layout = 1 Or Layout = 39 Then
      sql = "UPDATE MENU SET ATIVADO = 1 WHERE NOMESUBNOVO = 'mnuPrd_PCC_AEs'"
      Banco.Execute sql
   Else
      sql = "UPDATE MENU SET ATIVADO = 0 WHERE NOMESUBNOVO = 'mnuPrd_PCC_AEs'"
      Banco.Execute sql
   End If
   
   sql = "UPDATE MENU SET ATIVADO = 0 WHERE NOMESUBNOVO = 'mnuPrd_PCC_ASL'"
   Banco.Execute sql
   
   sql = "ALTER TMPLAU_MODELOVARIAVEL ALTER COLUMN RESULTADO VARCHAR(255) NULL"
   Banco.Execute sql
   
   sql = "alter table movim_int add PERCIAMSPEHOSPITAL money"
   Banco.Execute sql
   
   sql = "alter table movim_int add PERCIAMSPEPROFISSIONAL money"
   Banco.Execute sql
   
   sql = "alter table movim_int add PERCIAMSPEEXAME money"
   Banco.Execute sql
   
   sql = "alter table movim_amb add PERCIAMSPEHOSPITAL money"
   Banco.Execute sql
   
   sql = "alter table movim_amb add PERCIAMSPEPROFISSIONAL money"
   Banco.Execute sql
   
   sql = "alter table movim_amb add PERCIAMSPEEXAME money"
   Banco.Execute sql
   
   sql = "alter table movim_ext add PERCIAMSPEHOSPITAL money"
   Banco.Execute sql
   
   sql = "alter table movim_ext add PERCIAMSPEPROFISSIONAL money"
   Banco.Execute sql
   
   sql = "alter table movim_ext add PERCIAMSPEEXAME money"
   Banco.Execute sql
Exit Function
Erro:
   Resume Next
End Function
   
Public Function AtualizaMes012013()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '012013'"
   Banco.Execute sql
   
   sql = "ALTER TABLE PSICOTROPICOMOVIMENTO ADD TIPOLIVRO VARCHAR(50) NULL"
   Banco.Execute sql
   
   'RAPHAEL 02/01/2012 17:41
   'LIBERADO ESSE MENU, APENAS PARA FERNANDÓPOLIS, PARA CONTAGEM DO ESTOQUE
   If Layout = 12 Then
      sql = "UPDATE MENU SET ATIVADO = 1 WHERE NOMESUBNOVO = 'mnuPrd_PCC_ASL'"
      Banco.Execute sql
   End If
   
   If Layout = 1 Then
      sql = "UPDATE MENU SET ATIVADO = 1 WHERE NOMESUBNOVO = 'mnuPrd_PCC_AEs'"
      Banco.Execute sql
   End If
   
   sql = "UPDATE MENU SET ATIVADO = 0 WHERE NOMESUBNOVO = 'mnuCom_Sco_Min'"
   Banco.Execute sql
   
   sql = "SET ANSI_NULLS ON" & Chr(13)
   sql = sql & "GO" & Chr(13)
   sql = sql & "SET QUOTED_IDENTIFIER ON" & Chr(13)
   sql = sql & "GO" & Chr(13)
   sql = sql & "SET ANSI_PADDING ON" & Chr(13)
   sql = sql & "GO" & Chr(13)
   sql = sql & "CREATE TABLE [dbo].[FINOTAFISCAL](" & Chr(13)
   sql = sql & "   [NOTAFISCAL] [decimal](13, 2) NOT NULL," & Chr(13)
   sql = sql & "   [NOME] [varchar](155) NULL," & Chr(13)
   sql = sql & "   [CPF] [varchar](30) NULL," & Chr(13)
   sql = sql & "   [ENDERECO] [varchar](155) NULL," & Chr(13)
   sql = sql & "   [CIDADE] [varchar](155) NULL," & Chr(13)
   sql = sql & "   [UF] [varchar](2) NULL," & Chr(13)
   sql = sql & "   [CEP] [varchar](50) NULL," & Chr(13)
   sql = sql & "   [TELEFONE] [varchar](20) NULL," & Chr(13)
   sql = sql & "   [ATUALIZACAO] [varchar](155) NULL," & Chr(13)
   sql = sql & "   [DESCRICAO] [varchar](255) NULL," & Chr(13)
   sql = sql & "   [IMPRESSO] [int] NULL," & Chr(13)
   sql = sql & "   [IMPRESSOATUALIZACAO] [varchar](255) NULL," & Chr(13)
   sql = sql & "   [DATACANCELADO] [datetime] NULL," & Chr(13)
   sql = sql & "   [USUARIOCANCELADO] [varchar](100) NULL," & Chr(13)
   sql = sql & " CONSTRAINT [PK_FINOTAFISCAL] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "([NOTAFISCAL] ASC" & Chr(13)
   sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & Chr(13)
   sql = sql & ") ON [PRIMARY]" & Chr(13)
   sql = sql & "GO" & Chr(13)
   sql = sql & "SET ANSI_PADDING OFF" & Chr(13)
   sql = sql & "GO" & Chr(13)
   Banco.Execute sql
   
   sql = "INSERT INTO MENU (MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,NOMESUBAUX,NIVELVISIBILIDADE) " & Chr(13)
   sql = sql & "SELECT MAX(MENU)+1,'Tempo de Atendimento por Ala Hospitalar','MnuRec_Rel_TAH','','MnuRec_Rel_TAH',1,1,'0213150000','MnuRec_Rel_TAH',1 FROM MENU" & Chr(13)
   Banco.Execute sql
   
   If Layout = 10 Then
           sql = "INSERT INTO MENU (MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,NOMESUBAUX,NIVELVISIBILIDADE) " & Chr(13)
           sql = sql & "SELECT MAX(MENU)+1,'Relatorio de Diárias','MnuRec_Rel_Dia','','MnuRec_Rel_Dia',1,1,'0213160000','MnuRec_Rel_Dia',1 FROM MENU" & Chr(13)
           Banco.Execute sql
   End If
   
   sql = "ALTER TABLE REQUISICAOPRODUTO ADD USUARIOSELECIONA INT NULL" & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD ATESTADO VARCHAR(2000) NULL"
   Banco.Execute sql
   
   sql = "CREATE TABLE PEDIDOCOMPRA1EXCLUIDO( " & Chr(13)
   sql = sql & "PEDIDOCOMPRA INT, " & Chr(13)
   sql = sql & "CENTROCUSTO INT, " & Chr(13)
   sql = sql & "FABRICANTE INT, " & Chr(13)
   sql = sql & "ATUALIZACAO VARCHAR(100), " & Chr(13)
   sql = sql & "COTACAO INT, " & Chr(13)
   sql = sql & "ATUALIZACAOEXCLUSAO VARCHAR(100), " & Chr(13)
   sql = sql & "DATAOPERACAO DATETIME) " & Chr(13)
   Banco.Execute sql

   sql = "CREATE TABLE PEDIDOCOMPRAPRODUTO1EXCLUIDO( " & Chr(13)
   sql = sql & "PEDIDOCOMPRA INT, " & Chr(13)
   sql = sql & "PRODUTO INT, " & Chr(13)
   sql = sql & "QUANTIDADE INT, " & Chr(13)
   sql = sql & "ATUALIZACAO VARCHAR(100), " & Chr(13)
   sql = sql & "ATUALIZACAOEXCLUSAO VARCHAR(100), " & Chr(13)
   sql = sql & "DATAOPERACAO DATETIME)"
   Banco.Execute sql
         
    'Samuel / Adição de colunas para auxílio no form de contagem de estoque
   sql = "ALTER TABLE TMP_ALTERACAO_SALDO "
   sql = sql & "ADD GRUPO INT NULL, "
   sql = sql & "TIPO CHAR(2) NULL, " & Chr(13)
   sql = sql & "FAMILIA INT NULL, " & Chr(13)
   sql = sql & "PRATELEIRA VARCHAR(100) NULL, " & Chr(13)
   sql = sql & "PRODUTONOME VARCHAR(100) NULL"
   Banco.Execute sql
   
   'MOVIM_INT
   sql = "ALTER TABLE MOVIM_INT_TMP_SUS ADD PLANTAO INT NULL"
   Banco.Execute sql
   
   'Samuel / Cria tabela para guardar os filtros utilizados pelo usuário na contagem do estoque
   sql = "CREATE TABLE TMP_ALTERACAO_SALDO_FILTROS (" & Chr(13)
   sql = sql & "TIPOFILTRO  CHAR(1) NOT NULL, " & Chr(13)
   sql = sql & "VALORFILTRO VARCHAR(100) NOT NULL, " & Chr(13)
   sql = sql & " IP VarChar(100) not null)"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_INT_TMP_SUS ADD PERCIAMSPEPROFISSIONAL MONEY NULL"
   Banco.Execute sql
         
   'Samuel / Campo complemento estava permitindo apenas a 20 caracteres
   sql = "ALTER TABLE FICHAS " & Chr(13)
   sql = sql & "ALTER COLUMN COMPLEMENTO VARCHAR(255) NOT NULL"
   Banco.Execute sql
   

   sql = "ALTER TABLE MOVIM_INT_TMP_SUS ADD PERCIAMSPEHOSPITAL MONEY NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO " & Chr(13)
   sql = sql & "ALTER COLUMN COMPLEMENTOPASSAGEM VARCHAR(255) NULL"
   Banco.Execute sql
   

   sql = "ALTER TABLE MOVIM_INT_TMP_SUS ADD PERCIAMSPEEXAME MONEY NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE RESPONSAVEL " & Chr(13)
   sql = sql & "ALTER COLUMN COMPLEMENTO VARCHAR(255) NULL"
   Banco.Execute sql
   
   'MOVIM_AMB
   sql = "ALTER TABLE MOVIM_AMB_TMP_SUS ADD PLANTAO INT NULL"
   Banco.Execute sql
   
   'Samuel / Atribui vazio para o campo CID2 da tabela Complemento quando
   'este estiver com valor diferente do tamanho do codigo CID
   sql = "UPDATE COMPLEMENTAR SET CID2 = '' " & Chr(13)
   sql = sql & "Where Len(CID2) > 4 " & Chr(13)
   sql = sql & "Or (Len(CID2) < 3 And Len(CID2) > 0)"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_AMB_TMP_SUS ADD PERCIAMSPEPROFISSIONAL MONEY NULL"
   Banco.Execute sql
   
   sql = "UPDATE DADOSAIH SET CID2 = '' " & Chr(13)
   sql = sql & "Where Len(CID2) > 4 " & Chr(13)
   sql = sql & "Or (Len(CID2) < 3 And Len(CID2) > 0)"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_AMB_TMP_SUS ADD PERCIAMSPEHOSPITAL MONEY NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE REQUISICAOPRODUTO ADD DATASELECAO DATETIME"
   Banco.Execute sql
      
   sql = "ALTER TABLE MOVIM_AMB_TMP_SUS ADD PERCIAMSPEEXAME MONEY NULL"
   Banco.Execute sql
   
   'MOVIM_EXT
   sql = "ALTER TABLE MOVIM_EXT_TMP_SUS ADD PLANTAO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_EXT_TMP_SUS ADD PERCIAMSPEPROFISSIONAL MONEY NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_EXT_TMP_SUS ADD PERCIAMSPEHOSPITAL MONEY NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_EXT_TMP_SUS ADD PERCIAMSPEEXAME MONEY NULL"
   Banco.Execute sql
   
   If Layout = 23 Then
      sql = "UPDATE MENU SET NOMECAPTION = 'Prontuário Eletrônico' WHERE NOMESUBNOVO = 'mnuMed_EPr'"
      Banco.Execute sql
   End If
   

   sql = "CREATE TABLE [dbo].[SAME_MOVIMENTACAO_PRONTUARIO_EXT](" & Chr(13)
   sql = sql & "[SEQUENCIA] [int] IDENTITY(1,1) NOT NULL," & Chr(13)
   sql = sql & "[REGISTRO] [int] NULL," & Chr(13)
   sql = sql & "[LOCALORIGEM] [int] NULL," & Chr(13)
   sql = sql & "[USUARIOORIGEM] [int] NULL," & Chr(13)
   sql = sql & "[DATA] [datetime] NULL," & Chr(13)
   sql = sql & "[HORA] [char](5) NULL," & Chr(13)
   sql = sql & "[LOCALDESTINO] [int] NULL," & Chr(13)
   sql = sql & "[USUARIODESTINO] [int] NULL," & Chr(13)
   sql = sql & "[TIPOARMAZENAMENTO] [int] NULL," & Chr(13)
   sql = sql & "[ATUALIZACAO] [varchar](50) NULL," & Chr(13)
   sql = sql & "[ANESTESIA] [int] NULL," & Chr(13)
   sql = sql & "[PEDIATRIA] [int] NULL," & Chr(13)
   sql = sql & "[OBS_AVALIACAO] [varchar](155) NULL," & Chr(13)
   sql = sql & "[OBSERVACAO] [varchar](155) NULL," & Chr(13)
   sql = sql & "CONSTRAINT [PK_SAME_MOVIMENTACAO_PRONTUARIO_EXT] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & "([SEQUENCIA] ASC" & Chr(13)
   sql = sql & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]" & Chr(13)
   sql = sql & ") ON [PRIMARY]" & Chr(13)
   sql = sql & "SET ANSI_PADDING OFF" & Chr(13)
   sql = sql & "ALTER TABLE [dbo].[SAME_MOVIMENTACAO_PRONTUARIO_EXT]  WITH NOCHECK ADD  CONSTRAINT [FK_SAME_MOVIMENTACAO_PRONTUARIO_EXTERNO] FOREIGN KEY([REGISTRO])" & Chr(13)
   sql = sql & "REFERENCES [dbo].[Externo] ([Registro])" & Chr(13)
   sql = sql & "ALTER TABLE [dbo].[SAME_MOVIMENTACAO_PRONTUARIO_EXT] CHECK CONSTRAINT [FK_SAME_MOVIMENTACAO_PRONTUARIO_EXTERNO]" & Chr(13)
   sql = sql & "ALTER TABLE [dbo].[SAME_MOVIMENTACAO_PRONTUARIO_EXT]  WITH CHECK ADD  CONSTRAINT [FK_SAME_MOVIMENTACAO_PRONTUARIO_USUARIO_EXT] FOREIGN KEY([USUARIOORIGEM])" & Chr(13)
   sql = sql & "REFERENCES [dbo].[USUARIO] ([USUARIO])" & Chr(13)
   sql = sql & "ALTER TABLE [dbo].[SAME_MOVIMENTACAO_PRONTUARIO_EXT] CHECK CONSTRAINT [FK_SAME_MOVIMENTACAO_PRONTUARIO_USUARIO_EXT]" & Chr(13)
   sql = sql & "ALTER TABLE [dbo].[SAME_MOVIMENTACAO_PRONTUARIO_EXT]  WITH CHECK ADD  CONSTRAINT [FK_SAME_MOVIMENTACAO_PRONTUARIO_USUARIO1_EXT] FOREIGN KEY([USUARIODESTINO])" & Chr(13)
   sql = sql & "REFERENCES [dbo].[USUARIO] ([USUARIO])" & Chr(13)
   sql = sql & "ALTER TABLE [dbo].[SAME_MOVIMENTACAO_PRONTUARIO_EXT] CHECK CONSTRAINT [FK_SAME_MOVIMENTACAO_PRONTUARIO_USUARIO1_EXT]" & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_PRESCRICAOELETRONICA_REPETIR ADD USUARIOINCLUSAO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_PRESCRICAOELETRONICA_REPETIR ADD OBSERVACAO VARCHAR(255) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RETORNO INT NULL"
   Banco.Execute sql
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes022013()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '022013'"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD RETORNO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARTICULAR ADD QUANTIDADEAUXILIAR INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARTICULAR ADD UTILIZAVIDEO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARTICULAR ADD PORTEANEST TINYINT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD LOGO IMAGE NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_TISS_OUTRASDESPESAS ADD CONVENIO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PATRI_REQUISICAOMANUTENCAOPREDIAL ALTER COLUMN DESCRICAOFALHAAPRESENTADA VARCHAR(8000) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_TISS_HONORARIOINDIVIDUAL ADD CONVENIO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_TISS_OUTRASDESPESAS ADD CONVENIO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE SAVELOG..AMBULATORIALLOG ADD RETAGUADA INT NULL"
   Banco.Execute sql
   
   If Layout = 1 Or Layout = 30 Then
      sql = "UPDATE MENU SET ATIVADO = 1 WHERE NOMESUBNOVO = 'mnuPrd_PCC_AEs'"
      Banco.Execute sql
   End If
   
   sql = "ALTER TABLE tmpREGISTROGERAL ALTER COLUMN PROCEDIMENTONOME VARCHAR(500) NULL"
   Banco.Execute sql
   
      'Relatório de Consumo Faturado por Centro de Custo
   sql = "INSERT INTO MENU(MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,NOMESUBAUX,NIVELVISIBILIDADE) "
   sql = sql & "SELECT MAX(MENU)+1,'Consumo Faturado por Centro de Custo', 'mnuFat_Rel_FCC', ' ', 'mnuFat_Rel_FCC',1, 1, '0723130000', 'mnuFat_Rel_FCC', 1 FROM MENU"
   Banco.Execute sql
   
   
    'View V_FATURAMENTO
    sql = ""
    sql = sql & " CREATE VIEW V_FATURAMENTO AS "
    'INTERNO
    sql = sql & " SELECT "
    sql = sql & "   'INTERNO' AS TIPOATENDIMENTO,"
    sql = sql & "   B.FICHA, "
    sql = sql & "   A.REGISTRO,"
    sql = sql & "   D.FATURA,"
    sql = sql & "   B.NOME,"
    sql = sql & "   G.CONVENIO AS CD_CONVENIO,"
    sql = sql & "   G.DESCRICAO AS CONVENIO,"
    sql = sql & "   C.TIPOLANCAMENTO,"
    sql = sql & "   C.PROCEDIMENTO,"
    sql = sql & "   C.PROCEDIMENTONOME,"
    sql = sql & "   C.QUANTIDADE,"
    sql = sql & "   C.VALORUNITARIO,"
    sql = sql & "   ISNULL(C.QUANTIDADE,0) * ISNULL(C.VALORUNITARIO,0) AS VALORTOTAL,"
    sql = sql & "   CAST(D.ANO AS CHAR(4))+'/'+ (CASE WHEN D.MES < 10 THEN '0' ELSE '' END)+CAST(D.MES AS CHAR(2)) AS COMPETENCIA,"
    sql = sql & "   D.DATAINICIAL,"
    sql = sql & "   D.DATAFINAL,"
    sql = sql & "   E.CENTROCUSTO AS CD_CC_ATENDIMENTO,"
    sql = sql & "   E.DESCRICAO AS CC_ATENDIMENTO,"
    sql = sql & "   F.CENTROCUSTO AS CD_CC_PRODUCAO,"
    sql = sql & "   F.DESCRICAO AS CC_PRODUCAO"
    sql = sql & " FROM INTERNO A WITH (NOLOCK)"
    sql = sql & " INNER JOIN FICHAS B WITH (NOLOCK) ON B.FICHA = A.FICHA"
    sql = sql & " INNER JOIN MOVIM_INT C WITH (NOLOCK) ON C.REGISTRO = A.REGISTRO"
    sql = sql & " INNER JOIN CONVENIODATA_PERIODO D WITH (NOLOCK) ON D.FATURA = A.FATURACONTA AND A.CONVENIO = D.CONVENIO"
    sql = sql & " INNER JOIN CONVENIOS G WITH (NOLOCK) ON G.CONVENIO = A.CONVENIO "
    sql = sql & " LEFT JOIN CENTROCUSTO E WITH (NOLOCK) ON A.CENTROCUSTO = E.CENTROCUSTO"
    sql = sql & " LEFT JOIN CENTROCUSTO F WITH (NOLOCK) ON C.CENTROCUSTO = F.CENTROCUSTO"
    'AMBULATORIAL
    sql = sql & " UNION ALL"
    sql = sql & " SELECT "
    sql = sql & "   'AMBULATORIAL' AS TIPOATENDIMENTO,"
    sql = sql & "   B.FICHA,"
    sql = sql & "   A.REGISTRO,"
    sql = sql & "   D.FATURA,   "
    sql = sql & "   B.NOME,"
    sql = sql & "   G.CONVENIO AS CD_CONVENIO,"
    sql = sql & "   G.DESCRICAO AS CONVENIO,    "
    sql = sql & "   C.TIPOLANCAMENTO,"
    sql = sql & "   C.PROCEDIMENTO,"
    sql = sql & "   C.PROCEDIMENTONOME,"
    sql = sql & "   C.QUANTIDADE,"
    sql = sql & "   C.VALORUNITARIO,"
    sql = sql & "   ISNULL(C.QUANTIDADE,0) * ISNULL(C.VALORUNITARIO,0) AS VALORTOTAL,   "
    sql = sql & "   CAST(D.ANO AS CHAR(4))+'/'+ (CASE WHEN D.MES < 10 THEN '0' ELSE '' END)+CAST(D.MES AS CHAR(2)) AS COMPETENCIA,"
    sql = sql & "   D.DATAINICIAL,"
    sql = sql & "   D.DATAFINAL,"
    sql = sql & "   E.CENTROCUSTO AS CD_CC_ATENDIMENTO,"
    sql = sql & "   E.DESCRICAO AS CC_ATENDIMENTO,"
    sql = sql & "   F.CENTROCUSTO AS CD_CC_PRODUCAO,"
    sql = sql & "   F.DESCRICAO AS CC_PRODUCAO"
    sql = sql & " FROM AMBULATORIAL A WITH (NOLOCK)"
    sql = sql & " INNER JOIN FICHAS B WITH (NOLOCK) ON B.FICHA = A.FICHA"
    sql = sql & " INNER JOIN MOVIM_AMB C WITH (NOLOCK) ON C.REGISTRO = A.REGISTRO"
    sql = sql & " INNER JOIN CONVENIODATA_PERIODO D WITH (NOLOCK) ON D.FATURA = A.FATURACONTA AND A.CONVENIO = D.CONVENIO"
    sql = sql & " INNER JOIN CONVENIOS G WITH (NOLOCK) ON G.CONVENIO = A.CONVENIO "
    sql = sql & " LEFT JOIN CENTROCUSTO E WITH (NOLOCK) ON A.CENTROCUSTO = E.CENTROCUSTO"
    sql = sql & " LEFT JOIN CENTROCUSTO F WITH (NOLOCK) ON C.CENTROCUSTO = F.CENTROCUSTO"
    sql = sql & " UNION ALL"
    'EXTERNO
    sql = sql & " SELECT "
    sql = sql & "   'EXTERNO' AS TIPOATENDIMENTO,"
    sql = sql & "   B.FICHA, "
    sql = sql & "   A.REGISTRO,"
    sql = sql & "   D.FATURA,   "
    sql = sql & "   B.NOME,"
    sql = sql & "   G.CONVENIO AS CD_CONVENIO,"
    sql = sql & "   G.DESCRICAO AS CONVENIO,    "
    sql = sql & "   C.TIPOLANCAMENTO,"
    sql = sql & "   C.PROCEDIMENTO,"
    sql = sql & "   C.PROCEDIMENTONOME,"
    sql = sql & "   C.QUANTIDADE,"
    sql = sql & "   C.VALORUNITARIO,"
    sql = sql & "   ISNULL(C.QUANTIDADE,0) * ISNULL(C.VALORUNITARIO,0) AS VALORTOTAL,"
    sql = sql & "   CAST(D.ANO AS CHAR(4))+'/'+ (CASE WHEN D.MES < 10 THEN '0' ELSE '' END)+CAST(D.MES AS CHAR(2)) AS COMPETENCIA,"
    sql = sql & "   D.DATAINICIAL,"
    sql = sql & "   D.DATAFINAL,"
    sql = sql & "   E.CENTROCUSTO AS CD_CC_ATENDIMENTO,"
    sql = sql & "   E.DESCRICAO AS CC_ATENDIMENTO,"
    sql = sql & "   F.CENTROCUSTO AS CD_CC_PRODUCAO,"
    sql = sql & "   F.DESCRICAO AS CC_PRODUCAO"
    sql = sql & " FROM EXTERNO A WITH (NOLOCK)"
    sql = sql & " INNER JOIN FICHAS B WITH (NOLOCK) ON B.FICHA = A.FICHA"
    sql = sql & " INNER JOIN MOVIM_EXT C WITH (NOLOCK) ON C.REGISTRO = A.REGISTRO"
    sql = sql & " INNER JOIN CONVENIODATA_PERIODO D WITH (NOLOCK) ON D.FATURA = A.FATURACONTA AND A.CONVENIO = D.CONVENIO"
    sql = sql & " INNER JOIN CONVENIOS G WITH (NOLOCK) ON G.CONVENIO = A.CONVENIO"
    sql = sql & " LEFT JOIN CENTROCUSTO E WITH (NOLOCK) ON A.CENTROCUSTO = E.CENTROCUSTO"
    sql = sql & " LEFT JOIN CENTROCUSTO F WITH (NOLOCK) ON C.CENTROCUSTO = F.CENTROCUSTO"
    Banco.Execute sql

   'Relatório de Faturas Emitidas
   sql = "INSERT INTO MENU(MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,NOMESUBAUX,NIVELVISIBILIDADE) "
   sql = sql & "SELECT MAX(MENU)+1,'Faturas Emitidas', 'mnuFat_Rel_FEM', ' ', 'mnuFat_Rel_FEM',1, 1, '0723140000', 'mnuFat_Rel_FEM', 1 FROM MENU"
   Banco.Execute sql
    
    'View V_ATENDIMENTO_LANCAMENTO
    sql = ""
    sql = sql & " CREATE VIEW V_ATENDIMENTO_LANCAMENTO AS"
    sql = sql & " SELECT 'I' AS TIPOANTEDIMENTO, I.CONVENIO, I.REGISTRO, I.FATURACONTA AS FATURA, "
    sql = sql & " MI.TIPOLANCAMENTO, MI.PROCEDIMENTO AS CD_PRODUTO, MI.PROCEDIMENTONOME AS PRODUTO,  "
    sql = sql & " MI.QUANTIDADE, MI.VALORUNITARIO, CONVERT(VARCHAR,MI.DATA,111) AS DATA, "
    sql = sql & " SUBSTRING(CONVERT(VARCHAR,MI.HORA,108),1,5) AS HORA"
    sql = sql & " FROM INTERNO I WITH (NOLOCK)"
    sql = sql & " INNER JOIN MOVIM_INT MI WITH (NOLOCK) ON"
    sql = sql & " I.REGISTRO = MI.REGISTRO "
    sql = sql & " UNION ALL"
    sql = sql & " SELECT 'A' AS TIPOATENDIMENTO, A.CONVENIO, A.REGISTRO, A.FATURACONTA AS FATURA, "
    sql = sql & " MA.TIPOLANCAMENTO, MA.PROCEDIMENTO AS CD_PRODUTO, MA.PROCEDIMENTONOME AS PRODUTO,  "
    sql = sql & " MA.QUANTIDADE, MA.VALORUNITARIO, CONVERT(VARCHAR,MA.DATA,111) AS DATA, "
    sql = sql & " SUBSTRING(CONVERT(VARCHAR,MA.HORA,108),1,5) AS HORA"
    sql = sql & " FROM AMBULATORIAL A WITH (NOLOCK)"
    sql = sql & " INNER JOIN MOVIM_AMB MA WITH (NOLOCK) ON"
    sql = sql & " A.REGISTRO = MA.REGISTRO "
    sql = sql & " UNION ALL"
    sql = sql & " SELECT 'E' AS TIPOATENDIMENTO, E.CONVENIO, E.REGISTRO, E.FATURACONTA AS FATURA, "
    sql = sql & " ME.TIPOLANCAMENTO, ME.PROCEDIMENTO AS CD_PRODUTO, ME.PROCEDIMENTONOME AS PRODUTO,  "
    sql = sql & " ME.QUANTIDADE, ME.VALORUNITARIO, CONVERT(VARCHAR,ME.DATA,111) AS DATA, "
    sql = sql & " SUBSTRING(CONVERT(VARCHAR,ME.HORA,108),1,5) AS HORA"
    sql = sql & " FROM EXTERNO E WITH (NOLOCK)"
    sql = sql & " INNER JOIN MOVIM_EXT ME WITH (NOLOCK) ON"
    sql = sql & " E.REGISTRO = ME.REGISTRO "
    Banco.Execute sql
   
    'View V_FATURAMENTO_FATURAS_EMITIDAS
    sql = sql & " CREATE VIEW V_FATURAMENTO_FATURAS_EMITIDAS AS"
    sql = sql & " SELECT "
    sql = sql & "   A.CONVENIO AS CD_CONVENIO,"
    sql = sql & "   B.DESCRICAO AS CONVENIO,"
    sql = sql & "   A.ANO AS ANO,"
    sql = sql & "   A.MES AS MES,"
    sql = sql & "   CAST(A.ANO AS CHAR(4)) +'/' + CAST(A.MES AS CHAR(2)) AS COMPETENCIA,"
    sql = sql & "   A.FATURA,"
    sql = sql & "   A.USUARIO,"
    sql = sql & "   A.ATUALIZACAO,"
    sql = sql & "   A.NOMEFATURA,"
    sql = sql & "   CONVERT(VARCHAR,A.DATAINICIAL,111) AS DATAFATINICIAL,"
    sql = sql & "   CONVERT(VARCHAR,A.DATAFINAL,111) AS DATAFATFINAL,"
    sql = sql & "   SUBSTRING(A.ATUALIZACAO,1,10) AS DATAEMISSAO,"
    sql = sql & "   SUBSTRING(A.ATUALIZACAO,12,5) AS HORAEMISSAO,"
    sql = sql & "   CONVERT(VARCHAR,A.DATAVENCIMENTO,111) AS DATAVENCIMENTO,"
    sql = sql & "   ROUND(SUM(C.VALORUNITARIO * C.QUANTIDADE),2) AS VALORTOTAL"
    sql = sql & " FROM CONVENIODATA_PERIODO A WITH (NOLOCK)"
    sql = sql & " INNER JOIN CONVENIOS B WITH (NOLOCK) ON"
    sql = sql & " A.CONVENIO = B.CONVENIO"
    sql = sql & " INNER JOIN V_ATENDIMENTO_LANCAMENTO C WITH (NOLOCK) ON"
    sql = sql & " C.CONVENIO = A.CONVENIO"
    sql = sql & " AND C.FATURA = A.FATURA"
    sql = sql & " GROUP BY A.CONVENIO,B.DESCRICAO,A.ANO, A.MES, A.FATURA,   A.USUARIO, A.ATUALIZACAO, A.NOMEFATURA,"
    sql = sql & " CONVERT(VARCHAR,A.DATAINICIAL,111),CONVERT(VARCHAR,A.DATAFINAL,111),"
    sql = sql & " SUBSTRING(A.ATUALIZACAO,1,10), SUBSTRING(A.ATUALIZACAO,12,5), CONVERT(VARCHAR,A.DATAVENCIMENTO,111)"
    Banco.Execute sql
    
   sql = " ALTER TABLE MOVIMENTOCC ADD USUARIO_ATENDEU INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIALPROCEDIMENTO ADD PROCEDIMENTONOMETUSS VARCHAR(255) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIALPROCEDIMENTO ADD PROCEDIMENTOTUSS INT NULL"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO_NASCIDOS ADD TIPOOBITO INT"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO_NASCIDOS ADD DECLARACAOOBITO VARCHAR(20)"
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO_NASCIDOS ADD DATAOBITO DATETIME"
   Banco.Execute sql
   
   sql = " ALTER TABLE FIEMISSAOCHEQUEAVULSO"
   sql = sql & " ALTER COLUMN NOMINAL VARCHAR(100) NULL"
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPREGISTROGERAL "
   sql = sql & " ALTER COLUMN PROCEDIMENTONOME VARCHAR(255) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRESCRICAOELETRONICAPERIODO_AMB "
   sql = sql & " ADD MEDICO_AUTORIZADOR INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRESCRICAOELETRONICAPERIODO_EXT "
   sql = sql & " ADD MEDICO_AUTORIZADOR INT NULL"
   Banco.Execute sql
   
   sql = "IF  EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__CONVENIOS__REGIS__3CB4DFB3]') AND type = 'D')"
   sql = sql & "BEGIN" & Chr(13)
   sql = sql & "ALTER TABLE [dbo].[Convenios] DROP CONSTRAINT [DF__CONVENIOS__REGIS__3CB4DFB3]" & Chr(13)
   sql = sql & "END"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ALTER COLUMN REGISTROANS VARCHAR(30) NULL"
   Banco.Execute sql
   
   'CRIA OS CAMPOS NECESSÁRIOS PARA A DISPENSAÇÃO ITEM X ITEM
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_EXT ADD REGISTRO_SEQUENCIA BIGINT NOT NULL DEFAULT(0)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_EXT DROP CONSTRAINT PK_PREELETPROCEDIMENTOENFERMAGEM_EXT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_EXT" & Chr(13)
   sql = sql & "ADD CONSTRAINT PK_PREELETPROCEDIMENTOENFERMAGEM_EXT PRIMARY KEY CLUSTERED (REGISTRO_SEQUENCIA, REGISTRO, SEQUENCIA, PROCEDIMENTO, DATA, HORA)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_AMB ADD REGISTRO_SEQUENCIA BIGINT NOT NULL DEFAULT(0)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_AMB DROP CONSTRAINT PK_PREELETPROCEDIMENTOENFERMAGEM_AMB"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_AMB" & Chr(13)
   sql = sql & "ADD CONSTRAINT PK_PREELETPROCEDIMENTOENFERMAGEM_AMB PRIMARY KEY CLUSTERED (REGISTRO_SEQUENCIA, REGISTRO, SEQUENCIA, PROCEDIMENTO, DATA, HORA)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT ADD REGISTRO_SEQUENCIA BIGINT NOT NULL DEFAULT(0)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT DROP CONSTRAINT PK_PREELETPROCEDIMENTOENFERMAGEM_INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT" & Chr(13)
   sql = sql & "ADD CONSTRAINT PK_PREELETPROCEDIMENTOENFERMAGEM_INT PRIMARY KEY CLUSTERED (REGISTRO_SEQUENCIA, REGISTRO, SEQUENCIA, PROCEDIMENTO, DATA, HORA)"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_EMPRESTIMO ADD FABRICANTE VARCHAR(200) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETROETIQUETA"
   sql = sql & " ADD CRF INT NULL "
   Banco.Execute sql
   
   '-----------------------------------------------------------------------
   'Samuel / Atribui vazio para o campo CID2 da tabela Complemento e DADOSAIH
   'quando este estiver com valor diferente do tamanho do codigo CID
   '-----------------------------------------------------------------------
   sql = "UPDATE COMPLEMENTAR SET CID2 = '' " & Chr(13)
   sql = sql & "Where Len(CID2) > 4 " & Chr(13)
   sql = sql & "Or (Len(CID2) < 3 And Len(CID2) > 0)"
   Banco.Execute sql
   
   sql = "UPDATE DADOSAIH SET CID2 = '' " & Chr(13)
   sql = sql & "Where Len(CID2) > 4 " & Chr(13)
   sql = sql & "Or (Len(CID2) < 3 And Len(CID2) > 0)"
   Banco.Execute sql
   '-----------------------------------------------------------------------
   
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes032013()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '032013'"
   Banco.Execute sql
   
   'RAPHAEL 13/03/2013 11:50
   'COMENTADA A FUNÇÃO, MAS, PRECISA SER EXECUTADA DIRETO NA BASE DO CLIENTE
   'POR CONTA DA VELOCIDADE DA CONSULTA SER DEMORADA...
'   sql = "CREATE NONCLUSTERED INDEX IX_PRODUTOLOG ON SAVELOG..PRODUTOLOG(DATAOPERACAO, PRODUTO, CENTROCUSTO)"
'   Banco.Execute sql

   If Layout = 1 Then 'RELATÓRIO DE LOG
      sql = "INSERT INTO MENU SELECT MAX(MENU + 1), 'Relatório de Log Controle', 'mnuInf_Log', '2013-03-21', 'mnuInf_Log', 1, 1, '0818000000', 'Inf_Log', 1, NULL, NULL FROM MENU"
      Banco.Execute sql
   End If
   
   sql = "ALTER TABLE INSUMOS ADD QTDEMAXINTERNO MONEY"
   Banco.Execute sql

   sql = "ALTER TABLE INSUMOS ADD QTDEMAXAMBULATORIAL MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE INSUMOS ADD QTDEMAXEXTERNO MONEY"
   Banco.Execute sql
   
   '-----------------------------------------------------------------------
   'Samuel / Atribui vazio para o campo CID2 da tabela Complemento e DADOSAIH
   'quando este estiver com valor diferente do tamanho do codigo CID
   '-----------------------------------------------------------------------
   sql = "UPDATE COMPLEMENTAR SET CID2 = '' " & Chr(13)
   sql = sql & "Where Len(CID2) > 4 " & Chr(13)
   sql = sql & "Or (Len(CID2) < 3 And Len(CID2) > 0)"
   Banco.Execute sql
   
   sql = "UPDATE DADOSAIH SET CID2 = '' " & Chr(13)
   sql = sql & "Where Len(CID2) > 4 " & Chr(13)
   sql = sql & "Or (Len(CID2) < 3 And Len(CID2) > 0)"
   Banco.Execute sql
   '-----------------------------------------------------------------------
   'Samuel / Parametro adicionado na tabela de usuário
   'pois quando o cliente utiliza pré-devolução
   'serve como exceção
   'Solicitação de Araras
   sql = " ALTER TABLE USUARIO ADD PERMITEDEVOLUCAODIRETA INT NULL"
   Banco.Execute sql
   
   sql = "CREATE TABLE MOTIVO_RECOLETA( " & Chr(13)
   sql = sql & "CODIGO INT, " & Chr(13)
   sql = sql & "DESCRICAO VARCHAR(20), " & Chr(13)
   sql = sql & "CONSTRAINT [PK_MOTIVO_RECOLETA] PRIMARY KEY CLUSTERED (Codigo))"
   Banco.Execute sql
   
   sql = "CREATE TABLE [INTERNO_NASCIDO_MOTIVO_RECOLETA]( " & Chr(13)
   sql = sql & "[REGISTRO] [int] NOT NULL, " & Chr(13)
   sql = sql & "[NASCIDO] [int] NOT NULL, " & Chr(13)
   sql = sql & "[CODIGO] [int] NOT NULL, " & Chr(13)
   sql = sql & "CONSTRAINT [PK_INTERNO_NASCIDO_MOTIVO_RECOLETA] PRIMARY KEY CLUSTERED ([REGISTRO], [NASCIDO], [Codigo]))"
   Banco.Execute sql
   
   sql = "ALTER TABLE [INTERNO_NASCIDO_MOTIVO_RECOLETA] ADD CONSTRAINT [FK_INTERNO_NASCIDO_MOTIVO_RECOLETA_INTERNO_NASCIDOS] FOREIGN KEY([REGISTRO], [NASCIDO]) " & Chr(13)
   sql = sql & "REFERENCES [INTERNO_NASCIDOS]([registro], [NASCIDO])"
   Banco.Execute sql
   
   sql = "ALTER TABLE [INTERNO_NASCIDO_MOTIVO_RECOLETA] ADD CONSTRAINT [FK_INTERNO_NASCIDO_MOTIVO_RECOLETA_MOTIVO_RECOLETA] FOREIGN KEY([CODIGO]) " & Chr(13)
   sql = sql & "REFERENCES [MOTIVO_RECOLETA]([Codigo])"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (1, '17OH')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (2, 'BIO')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (3, 'G6PD')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (4, 'GALT/GAOS')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (5, 'HB')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (6, 'IRT')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (7, 'LEUCIN')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (8, 'NTSH/NT4')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (9, 'PKU')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (10, 'TOXO M')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (11, 'AAAC')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (12, 'PREM')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (13, 'TF')"
   Banco.Execute sql
   
   sql = "INSERT [MOTIVO_RECOLETA] ([CODIGO], [DESCRICAO]) VALUES (14, 'TRANS')"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD OBSERVACAO VARCHAR(250)"
   Banco.Execute sql
   
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes042013()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '042013'"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD URGENCIA INT"
   Banco.Execute sql
   
   sql = "CREATE TABLE [CIRURGIACOMPLEMENTO]("
   sql = sql & " [CIRURGIACOMPLEMENTO] [int] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & " [DESCRICAO] [varchar](100) NULL, " & Chr(13)
   sql = sql & " [ATUALIZACAO] [varchar](100) NULL, " & Chr(13)
   sql = sql & " [DESATIVADO] [bit] NULL, " & Chr(13)
   sql = sql & " CONSTRAINT [PK_CIRURGIACOMPLEMENTO] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & " ([CIRURGIACOMPLEMENTO]))"
   Banco.Execute sql

   sql = "ALTER TABLE COTACAOITEM1 ALTER COLUMN MARCA VARCHAR(255) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE MATERIAL_AGENDAMENTO ADD COMPLEMENTO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCAPALOTE ADD CEPPACIENTE VARCHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCAPALOTE ADD ENDERECOPACIENTE VARCHAR(255)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCAPALOTE ADD ENDERECOCOMPLEMENTOPACIENTE VARCHAR(255)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCAPALOTE ADD NUMEROPACIENTE VARCHAR(255)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCAPALOTE ADD BAIRROPACIENTE VARCHAR(255)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCAPALOTE ADD TELEFONEPACIENTE VARCHAR(15)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCAPALOTE ADD EMAILPACIENTE VARCHAR(255)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_PRESCRICAOELETRONICA_REPETIR ADD INTERVALO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ADD VIA VARCHAR(30)"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_AMB ADD REGISTROCONSULTA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD UNIDADEORIGEM VARCHAR(100)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD GUIAPRINCIPAL VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD GUIAPRINCIPAL VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD GUIAPRINCIPAL VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRODUTOSIMPRO ADD QUANTIDADEFRACAO MONEY NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRODUTOSIMPRO ADD TIPOEMBALAGEM VARCHAR(3) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRODUTOSIMPRO ADD TIPOFRACAO VARCHAR(4) NULL"
   Banco.Execute sql

   sql = "ALTER TABLE PARAMETRO ADD ELEGIBILIDADEUNIDAUTORIZADAS VARCHAR(250)"
   Banco.Execute sql
   
   'Samuel CI  12761
   '-------------------------------------------------------
   sql = "CREATE TABLE UF ( " & Chr(13)
   sql = sql & " ID_UF  INT PRIMARY KEY," & Chr(13)
   sql = sql & " SIGLA  CHAR(2) NOT NULL," & Chr(13)
   sql = sql & " DESCRICAO VARCHAR(50) NOT NULL," & Chr(13)
   sql = sql & " REGIAO VarChar(50))" & Chr(13)
   Banco.Execute sql
   '-------------------------------------------------------
   
   'Samuel CI  12761
   '-------------------------------------------------------
    sql = "INSERT INTO UF VALUES(0,'AC','ACRE','NORTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(1,'AL','ALAGOAS','NORDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(2 ,'AP','AMAPA','NORTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(3 ,'AM','AMAZONAS','NORTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(4 ,'BA','BAHIA','NORDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(5 ,'CE','CEARA','NORDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(6 ,'DF','DISTRITO FEDERAL','CENTRO OESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(7 ,'ES','ESPIRITO SANTO','SUDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(9 ,'GO','GOIAS','CENTRO OESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(10,'MA','MARANHAO','NORDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(11,'MG','MINAS GERAIS','SUDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(12,'MS','MATO GROSSO DO SUL','CENTRO OESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(13,'MT','MATO GROSSO','CENTRO OESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(14,'PA','PARA','NORTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(15,'PB','PARAIBA','NORDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(16,'PE','PERNAMBUCO','NORDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(17,'PI','PIAUI','NORDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(18,'PR','PARANA','SUL')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(19,'RJ','RIO DE JANEIRO','SUDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(20,'RN','RIO GRANDE DO NORTE','NORDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(21,'RO','RONDONIA','NORTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(22,'RR','RORAIMA','NORTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(23,'RS','RIO GRANDE DO SUL','SUL')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(24,'SC','SANTA CATARINA','SUL')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(25,'SE','SERGIPE','NORDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(26,'SP','SAO PAULO','SUDESTE')"
    Banco.Execute sql
    sql = "INSERT INTO UF VALUES(27,'TO','TOCANTINS','NORTE')"
    Banco.Execute sql
   '-------------------------------------------------------
   
   'Samuel  CI 12805
   '-------------------------------------------------------
   sql = " alter table TMPSAME_MOVIMENTACAO_PRONTUARIO "
   sql = sql & " add USUARIOTRANSFERENCIA VARCHAR(50) NULL"
   Banco.Execute sql
   
   sql = " ALTER TABLE LEITOUNIDADE ADD QUANTIDADELEITO INT NULL "
   Banco.Execute sql
   '-------------------------------------------------------
   
   'Samuel CI 12823
   '-------------------------------------------------------
   sql = " alter table COTACAO1 alter column observacao varchar(2000) null "
   Banco.Execute sql
   
   sql = " alter table COTACAO1 alter column observacaopedido varchar(2000) null"
   Banco.Execute sql
   '-------------------------------------------------------
   
   sql = "ALTER TABLE PRODUTOCONVENIO ADD CONSTRAINT DF_PRODUTOCONVENIO_CONVLANCCONVENIO DEFAULT 1 FOR CONVLANCCONVENIO"
   Banco.Execute sql
   
   sql = "UPDATE PRODUTOCONVENIO SET CONVLANCCONVENIO = 1 WHERE CONVLANCCONVENIO IS NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD PLANTAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIAL ADD PLANTAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNO ADD PLANTAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD VALORIZAPLANTAODOMINGO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD VALORIZAPLANTAOSABADO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD VALORIZAPLANTAOFERIADO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD VALORIZAPLANTAODIASSEMANA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD HORAINICIALPLANTAOSEMANA DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD HORAFINALPLANTAOSEMANA DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PERCENTUALPLANTAO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD PROCEDIMENTOSPLANTAO VARCHAR(2000)"
   Banco.Execute sql
   
   If Layout = 47 Then
      sql = "INSERT INTO MENU SELECT MAX(MENU) + 1, 'Brasindice/ABCFarma', 'mnuPrd_Tbl_Atp_BraABC', '2013-04-26 SA', 'mnuPrd_Tbl_Atp_BraABC', 1, 2, '0111020100', 'mnuPrd_Tbl_Atp_BraABC', 1, NULL, NULL, NULL, NULL FROM MENU"
      Banco.Execute sql
      
      sql = "INSERT INTO MENU SELECT MAX(MENU) + 1, 'SIMPRO', 'mnuPrd_Tbl_Atp_Spo', '2013-04-26 SA', 'mnuPrd_Tbl_Atp_Spo', 1, 2, '0111020200', 'mnuPrd_Tbl_Atp_Spo', 1, NULL, NULL, NULL, NULL FROM MENU"
      Banco.Execute sql
   End If
   
'Samuel CI:12795 Criação de parametro que habilita
'valorização de insumo por plano
'--------------------------------------------------------------------------
   sql = "alter table parametro add insumo_valoriza_por_plano int null"
   Banco.Execute sql
   
   sql = "alter table planoconvenio drop CONSTRAINT [PK_PLANOCONVENIO]"
   Banco.Execute sql
   
   sql = "alter table planoconvenio alter column convenio smallint not null"
   Banco.Execute sql
   
   sql = " alter table planoconvenio add CONSTRAINT [PK_PLANOCONVENIO] PRIMARY KEY CLUSTERED([PLANO] ASC, [CONVENIO] ASC) "
   Banco.Execute sql

    sql = " create table convenioplanoinsumo( " & Chr(13)
    sql = sql & "   convenio smallint not null," & Chr(13)
    sql = sql & "   plano int not null," & Chr(13)
    sql = sql & "   insumo int not null," & Chr(13)
    sql = sql & "   valor money default(0) null," & Chr(13)
    sql = sql & "   atualizacao varchar(100) null," & Chr(13)
    sql = sql & "   ativo int default(1)," & Chr(13)
    sql = sql & "   constraint pk_convenioplanoinsumo primary key (convenio,plano,insumo)," & Chr(13)
    sql = sql & "   constraint fk_convenioplanoinsumo_convenio foreign key (convenio) References convenios(convenio)," & Chr(13)
    sql = sql & "   constraint fk_convenioplanoinsumo_plano foreign key (plano,convenio) References planoconvenio(plano,convenio)," & Chr(13)
    sql = sql & "   constraint fk_convenioplanoinsumo_insumo foreign key (insumo,convenio) References insumos(insumo,convenio)" & Chr(13)
    sql = sql & " )" & Chr(13)
    Banco.Execute sql
'--------------------------------------------------------------------------

   If Layout = 1 Then
      sql = "INSERT INTO MENU SELECT MAX(MENU) + 1, 'Motivos Auditoria', 'mnuPar_Cad_Fat_MAu', " & Chr(13)
      sql = sql & "'2013-04-11 SA', 'mnuPar_Cad_Fat_MAu', 1, 1, '0101020900', 'mnuPar_Cad_Fat_MAu', 1, " & Chr(13)
      sql = sql & "NULL, NULL FROM MENU"
      Banco.Execute sql
   End If
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes052013()
On Error GoTo Erro
   'Samuel CI:12795 Criação de parametro que habilita
'valorização de insumo por plano
'--------------------------------------------------------------------------
   sql = "alter table parametro add insumo_valoriza_por_plano int null"
   Banco.Execute sql
   
   sql = "alter table planoconvenio drop CONSTRAINT [PK_PLANOCONVENIO]"
   Banco.Execute sql
   
   sql = "alter table planoconvenio alter column convenio smallint not null"
   Banco.Execute sql
   
   sql = " alter table planoconvenio add CONSTRAINT [PK_PLANOCONVENIO] PRIMARY KEY CLUSTERED([PLANO] ASC, [CONVENIO] ASC) "
   Banco.Execute sql

    sql = " create table convenioplanoinsumo( " & Chr(13)
    sql = sql & "   convenio smallint not null," & Chr(13)
    sql = sql & "   plano int not null," & Chr(13)
    sql = sql & "   insumo int not null," & Chr(13)
    sql = sql & "   valor money default(0) null," & Chr(13)
    sql = sql & "   atualizacao varchar(100) null," & Chr(13)
    sql = sql & "   ativo int default(1)," & Chr(13)
    sql = sql & "   constraint pk_convenioplanoinsumo primary key (convenio,plano,insumo)," & Chr(13)
    sql = sql & "   constraint fk_convenioplanoinsumo_convenio foreign key (convenio) References convenios(convenio)," & Chr(13)
    sql = sql & "   constraint fk_convenioplanoinsumo_plano foreign key (plano,convenio) References planoconvenio(plano,convenio)," & Chr(13)
    sql = sql & "   constraint fk_convenioplanoinsumo_insumo foreign key (insumo,convenio) References insumos(insumo,convenio)" & Chr(13)
    sql = sql & " )" & Chr(13)
    Banco.Execute sql
'--------------------------------------------------------------------------

   If Layout = 1 Then
      sql = "INSERT INTO MENU SELECT MAX(MENU) + 1, 'Motivos Auditoria', 'mnuPar_Cad_Fat_MAu', " & Chr(13)
      sql = sql & "'2013-04-11 SA', 'mnuPar_Cad_Fat_MAu', 1, 1, '0101020900', 'mnuPar_Cad_Fat_MAu', 1, " & Chr(13)
      sql = sql & "NULL, NULL FROM MENU"
      Banco.Execute sql
   End If
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '052013'"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD CANCELADA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD CANCELADA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD CANCELADA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNOPROCEDIMENTO ADD QUANTIDADE_AUX MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIALPROCEDIMENTO ADD QUANTIDADE_AUX MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNOPROCEDIMENTO ADD QUANTIDADE_AUX MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNOPROCEDIMENTO ADD SEQUENCIA INT IDENTITY"
   Banco.Execute sql
   
   sql = "ALTER TABLE AMBULATORIALPROCEDIMENTO ADD SEQUENCIA INT IDENTITY"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNOPROCEDIMENTO ADD SEQUENCIA INT IDENTITY"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[CONTRATOSUNIMED](" & Chr(13)
   sql = sql & " [EMPRESA] [varchar](4) NOT NULL, & Chr(13)"
   sql = sql & " [UNIDADE] [varchar](4) NOT NULL, & Chr(13)"
   sql = sql & " [CONVENIO] [smallint] NOT NULL, & Chr(13)"
   sql = sql & " CONSTRAINT [PK_CONTRATOSUNIMED] PRIMARY KEY CLUSTERED & Chr(13)"
   sql = sql & " ( & Chr(13)"
   sql = sql & "  [EMPRESA] ASC, & Chr(13)"
   sql = sql & " [unidade] Asc & Chr(13)"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[CONTRATOSUNIMED]  ADD CONSTRAINT [FK_CONTRATOSUNIMED_Convenio] FOREIGN KEY([CONVENIO]) References [dbo].[Convenios]([Convenio])"
   Banco.Execute sql

   sql = "ALTER TABLE [dbo].[CONTRATOSUNIMED]  ADD  CONSTRAINT [FK_CONTRATOSUNIMED_UNIDADE] FOREIGN KEY([UNIDADE])References [dbo].[UNIMEDS]([unidade])"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[CONTRATOSUNIMEDCOPARTICIPACAO]("
   sql = sql & " [EMPRESA] [varchar](4) NOT NULL,"
   sql = sql & " [UNIDADE] [varchar](4) NOT NULL,"
   sql = sql & " [DESCRICAO] [varchar](250) NOT NULL,"
   sql = sql & " CONSTRAINT [PK_CONTRATOSUNIMEDCOPARTICIPACAO] PRIMARY KEY CLUSTERED"
   sql = sql & " ("
   sql = sql & " [EMPRESA] ASC,"
   sql = sql & " [unidade] Asc"
   sql = sql & " )"
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[CONTRATOSUNIMEDCOPARTICIPACAO] ADD  CONSTRAINT [FK_CONTRATOSUNIMEDCOPARTICIPACAO_UNIDADE] FOREIGN KEY([UNIDADE]) REFERENCES [dbo].[UNIMEDS] ([UNIDADE])"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD OBSERVACAOSOLICITACAO VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD GUIAPRINCIPAL VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD CANCELADA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD INDICACAOCLINICA VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD CID VARCHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD CARATERINTERNACAO VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD TIPOINTERNACAO VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD REGIMEINTERNACAO VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOAMBULATORIAL ADD DIASSOLICITADOS MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD OBSERVACAOSOLICITACAO VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD GUIAPRINCIPAL VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD CANCELADA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD INDICACAOCLINICA VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD CID VARCHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD CARATERINTERNACAO VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD TIPOINTERNACAO VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD REGIMEINTERNACAO VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOEXTERNO ADD DIASSOLICITADOS MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD OBSERVACAOSOLICITACAO VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD GUIAPRINCIPAL VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD CANCELADA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD INDICACAOCLINICA VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD CID VARCHAR(10)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD CARATERINTERNACAO VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD TIPOINTERNACAO VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD REGIMEINTERNACAO VARCHAR(2)"
   Banco.Execute sql
   
   sql = "ALTER TABLE AUTORIZACAOPROCEDIMENTOINTERNO ADD DIASSOLICITADOS MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCAPALOTE ALTER COLUMN TELEFONEPACIENTE VARCHAR(25) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE SAVELOG..FICHAS_LOG ALTER COLUMN COMPLEMENTO VARCHAR(255) NOT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE tmpINTERNACAO ADD CRMMEDICO INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PLANOCONVENIO ADD CODIGOCARTEIRINHA VARCHAR(7) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRESCRICAOELETRONICAPERIODO_INT ADD EMERGENCIA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRESCRICAOELETRONICAPERIODO_EXT ADD EMERGENCIA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRESCRICAOELETRONICAPERIODO_AMB ADD EMERGENCIA INT"
   Banco.Execute sql
      
    sql = "ALTER TABLE TMPINDICADOR ADD QUANTIDADEENFERMAGEM varchar(255) NULL"
    Banco.Execute sql
    
    sql = "ALTER TABLE tmpINTERNACAO ADD CPFRESPONSAVEL varchar(30) NULL"
    Banco.Execute sql
    
    sql = "ALTER TABLE tmpINTERNACAO ADD RGRESPONSAVEL varchar(30) NULL"
    Banco.Execute sql

   
   
   'Samuel: Atender protocolo 12926
   '------------------------------------------------------------------------
   sql = "ALTER TABLE TMPINDICADOR ADD LEGENDA VARCHAR(255) NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPINDICADOR ADD TOTALREALIZADA INT NULL"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPINDICADOR ADD TOTAL INT NULL"
   Banco.Execute sql
   '------------------------------------------------------------------------

    sql = "CREATE TABLE [dbo].[LOG_PASSAGEM]("
    sql = sql & "   [SEQUENCIA] [int] IDENTITY(1,1) NOT NULL,"
    sql = sql & "   [REGISTRO] [int] NULL,"
    sql = sql & "   [FICHA] [int] NULL,"
    sql = sql & "   [NOME] [varchar](250) NULL,"
    sql = sql & "   [CONVENIO] [int] NULL,"
    sql = sql & "   [NOMECONVENIO] [varchar](250) NULL,"
    sql = sql & "   [UNIDADE] [nchar](10) NULL,"
    sql = sql & "   [CARTEIRINHA] [varchar](50) NULL,"
    sql = sql & "   [GUIA] [varchar](50) NULL,"
    sql = sql & "   [DATAINTERNACAO] [datetime] NULL,"
    sql = sql & "   [HORAINTERNACAO] [datetime] NULL,"
    sql = sql & "   [DATAALTA] [datetime] NULL,"
    sql = sql & "   [MEDICO] [int] NULL,"
    sql = sql & "   [NOMEMEDICO] [varchar](250) NULL,"
    sql = sql & "   [DATAFATURAMENTO] [datetime] NULL,"
    sql = sql & "   [CENTROCUSTO] [int] NULL,"
    sql = sql & "   [DESCRICAOCENTROCUSTO] [varchar](250) NULL,"
    sql = sql & "   [CANCELADO] [int] NULL,"
    sql = sql & "   [CID] [nchar](10) NULL,"
    sql = sql & "   [TRAVALANC] [int] NULL,"
    sql = sql & "   [PROCSUS] [int] NULL,"
    sql = sql & "   [TIPOOPERACAO] [varchar](50) NULL,"
    sql = sql & "   [DATAOPERACAO] [datetime] NULL,"
    sql = sql & "   [ATUALIZACAO] [varchar](250) NULL,"
    sql = sql & "   [IP] [varchar](250) NULL,"
    sql = sql & "  CONSTRAINT [PK_LOG_PASSAGEM] PRIMARY KEY CLUSTERED "
    sql = sql & " ("
    sql = sql & "   [SEQUENCIA] ASC"
    sql = sql & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
    sql = sql & " ) ON [PRIMARY]"
    Banco.Execute sql
    
    sql = " CREATE TABLE [dbo].[LOG_ESTRUTURA]("
    sql = sql & "   [SEQUENCIA] [int] NOT NULL,"
    sql = sql & "   [TIPOLOG] [int] NOT NULL,"
    sql = sql & "   [CAMPOBANCO] [varchar](50) NULL,"
    sql = sql & "   [DESCRICAO] [varchar](50) NULL,"
    sql = sql & "   [TAMANHO] [int] NULL"
    sql = sql & "  ) ON [PRIMARY]"
    Banco.Execute sql
    
    sql = " CREATE TABLE [dbo].[LOG_MOVIMENTACAO]("
    sql = sql & "   [SEQUENCIA] [int] IDENTITY(1,1) NOT NULL,"
    sql = sql & "   [TIPOLANCAMENTO] [varchar](50) NULL,"
    sql = sql & "   [TIPOOPERACAO] [varchar](50) NULL,"
    sql = sql & "   [REGISTRO] [int] NULL,"
    sql = sql & "   [DATAOPERACAO] [datetime] NULL,"
    sql = sql & "   [USUARIO] [varchar](50) NULL,"
    sql = sql & "   [PROCEDIMENTO] [int] NULL,"
    sql = sql & "   [PROCEDIMENTONOME] [varchar](250) NULL,"
    sql = sql & "   [QUANTIDADE] [money] NULL,"
    sql = sql & "   [QUANTIDADEDISPENSADA] [money] NULL,"
    sql = sql & "   [VALOR] [money] NULL,"
    sql = sql & "   [CENTROCUSTO] [int] NULL,"
    sql = sql & "   [DESCRICAOCENTROCUSTO] [varchar](250) NULL,"
    sql = sql & "   [DATACONSUMO] [datetime] NULL,"
    sql = sql & "   [SEMBAIXA] [int] NULL,"
    sql = sql & "   [IP] [varchar](250) NULL,"
    sql = sql & "   CONSTRAINT [PK_LOG_MOVIMENTACAO] PRIMARY KEY CLUSTERED "
    sql = sql & "  ("
    sql = sql & "   [SEQUENCIA] ASC"
    sql = sql & "  )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
    sql = sql & "  ) ON [PRIMARY]"
    Banco.Execute sql
    
    sql = " CREATE TABLE [dbo].[PREELETPROCEDIMENTOENFERMAGEM_INT12022010]("
    sql = sql & "   [REGISTRO] [int] NOT NULL,"
    sql = sql & "   [SEQUENCIA] [smallint] NOT NULL,"
    sql = sql & "   [PROCEDIMENTO] [int] NOT NULL,"
    sql = sql & "   [DATA] [datetime] NOT NULL,"
    sql = sql & "   [HORA] [char](5) NOT NULL,"
    sql = sql & "   [CONFERIDO] [int] NOT NULL,"
    sql = sql & "   [NAOIMPRIMEPRODUTO] [int] NULL,"
    sql = sql & "   [NUMERO_JUNCAO] [int] NULL,"
    sql = sql & "   [NUMERO_LINHA] [int] NULL,"
    sql = sql & "   [QUANTIDADEINDIVIDUAL] [money] NULL,"
    sql = sql & "   [PERIODO] [int] NULL,"
    sql = sql & "   [MEDICACAO_ADMINISTRADA] [int] NULL,"
    sql = sql & "   [QUANTIDADECONVERTIDA] [money] NULL,"
    sql = sql & "   [ATUALIZACAO] [varchar](155) NULL"
    sql = sql & "  ) ON [PRIMARY]"
    Banco.Execute sql
    
    sql = " ALTER TABLE TMPCAPALOTE"
    sql = sql & "ALTER COLUMN TELEFONEPACIENTE varchar(20)"
    Banco.Execute sql
        
'-------------------------------------------------------------------------------------
    'Samuel CI:11309 2013/05/24 17:02
    'Parametrização para permitir seleção de convênio para unidades Unimed
    
    sql = " select * from menu where NOMESUBAUX = 'mnuPar_Cad_Ger_Uni' "
    Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
    
    If tbl.EOF = True Then
        sql = "INSERT INTO MENU( MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
        sql = sql & " NOMESUBAUX,NIVELVISIBILIDADE)"
        sql = sql & " SELECT MAX(MENU)+1,'Unimeds', 'MnuCGUni', ' ', 'mnuPar_Cad_Ger_Uni',"
        sql = sql & " 1, 1, '0101012200', 'mnuPar_Cad_Ger_Uni', 3 FROM MENU"
        Banco.Execute sql
    End If
    tbl.Close
    
    sql = " CREATE TABLE [dbo].[CONTRATOSUNIMED]("
    sql = sql & " [EMPRESA] [varchar](4) NOT NULL,"
    sql = sql & " [UNIDADE] [varchar](4) NOT NULL,"
    sql = sql & " [CONVENIO] [smallint] NOT NULL,"
    sql = sql & "  CONSTRAINT [PK_CONTRATOSUNIMED] PRIMARY KEY CLUSTERED"
    sql = sql & " ("
    sql = sql & " [EMPRESA] ASC,"
    sql = sql & " [unidade] Asc"
    sql = sql & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    sql = sql & " ) ON [PRIMARY]"
    Banco.Execute sql
    
    sql = " ALTER TABLE [dbo].[CONTRATOSUNIMED]"
    sql = sql & " ADD  CONSTRAINT [FK_CONTRATOSUNIMED_Convenio]"
    sql = sql & " FOREIGN KEY([CONVENIO]) REFERENCES [dbo].[Convenios] ([Convenio])"
    Banco.Execute sql

    sql = " ALTER TABLE [dbo].[CONTRATOSUNIMED]"
    sql = sql & " ADD  CONSTRAINT [FK_CONTRATOSUNIMED_UNIDADE]"
    sql = sql & " FOREIGN KEY([UNIDADE]) REFERENCES [dbo].[UNIMEDS] ([UNIDADE])"
    Banco.Execute sql
    
    sql = " alter table parametro"
    sql = sql & " add SELECIONACONVENIO BIT NULL"
    Banco.Execute sql
'-------------------------------------------------------------------------------------

    sql = " ALTER TABLE TMPCAPALOTE"
    sql = sql & "ALTER COLUMN TELEFONEPACIENTE varchar(50)"
    Banco.Execute sql
    
    sql = "ALTER TABLE parametro ADD insumo_valoriza_por_plano INT"
    Banco.Execute sql
    
    sql = "ALTER TABLE Externo  ADD PLANTAO INT"
    Banco.Execute sql
    
    sql = "ALTER TABLE Interno ADD PLANTAO INT"
    Banco.Execute sql
    
    sql = "ALTER TABLE Ambulatorial ADD PLANTAO INT"
    Banco.Execute sql
    
    sql = " ALTER TABLE PRODUTOSIMPRO_ASSOCIADO ADD EMBALAGEM VARCHAR(5)"
    Banco.Execute sql
    
    sql = " ALTER TABLE PRODUTOSIMPRO_ASSOCIADO ADD QUANTIDADE MONEY "
    Banco.Execute sql
    
Exit Function
Erro:
   Resume Next
End Function


Public Function AtualizaMes062013()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '062013'"
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_EMPRESTIMO ADD FABRICANTE VARCHAR(255)"
   Banco.Execute sql
   
   'RAPHAEL 18/06/2013 10:51
   'RELATÓRIO DE MEDICAMENTOS ADMINISTRADOS
   sql = "INSERT INTO MENU"
   sql = sql & "SELECT MAX(MENU) + 1, 'Medicamentos Administrados', 'mnuMed_Rel_Mad', '2013-06-18 SA', 'mnuMed_Rel_Mad', 1, 1,"
   sql = sql & "'0605030000', 'mnuMed_Rel_Mad', 1, NULL, NULL FROM MENU"
   Banco.Execute sql
   
   sql = "ALTER TABLE SUS_PROCEDIMENTO ADD CIHAENVIARCOMOAMBULATORIALINTERNOHOSPDIA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CIHA_INTERNO_CONS ADD IP VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE CIHA_INTERNO_IND ADD IP VARCHAR(250)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPTISS_SPSADT ADD VALIDADEGUIA DATETIME"
   Banco.Execute sql
   
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes072013()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '072013'"
   Banco.Execute sql
   
   sql = " CREATE TABLE CIHA_INTERNO_IND ( REGISTRO INT, IP VarChar(250) )"
   Banco.Execute sql
   
   sql = " CREATE TABLE CIHA_INTERNO_CONS ( REGISTRO INT, IP VarChar(250) )"
   Banco.Execute sql
   
   sql = " ALTER TABLE PRODUTOCONVENIO ADD [UNIDADE_FATURAMENTO_XML] [char](4) NULL "
   Banco.Execute sql
   
   sql = " ALTER TABLE PRODUTOCONVENIO ADD [FATOR_CONVERSOR_XML] [money] NULL "
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIOS ADD TISS_EXPORTA_CODIGO_INSUMO_CONVENIO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPCONTARECEBIMENTO ADD SALDODEVEDOR MONEY"
   Banco.Execute sql
   
   
   sql = " ALTER TABLE PRE_CUIDADO ADD DATA DATETIME NULL "
   Banco.Execute sql
   
   sql = " ALTER TABLE PRE_CUIDADO ADD USUARIO INT NULL "
   Banco.Execute sql
   
   sql = " ALTER TABLE INTERNO ADD DATAULTIMOPARTO DATETIME NULL "
   Banco.Execute sql

   sql = " ALTER TABLE AMBULATORIAL ADD PSI_ORDEMJUDICIAL BIT NULL "
   Banco.Execute sql

   sql = " ALTER TABLE EXTERNO ADD PSI_ORDEMJUDICIAL BIT NULL "
   Banco.Execute sql
    
   sql = " ALTER TABLE PRESCRICAOELETRONICAPERIODO_INT ADD USUARIOINCLUSAOPRESCRICAO INT NULL "
   Banco.Execute sql

    sql = " ALTER TABLE PRESCRICAOELETRONICAPERIODO_AMB ADD USUARIOINCLUSAOPRESCRICAO INT NULL "
    Banco.Execute sql

    sql = " ALTER TABLE PRESCRICAOELETRONICAPERIODO_EXT ADD USUARIOINCLUSAOPRESCRICAO INT NULL "
    Banco.Execute sql
        
    sql = " ALTER TABLE PACIENTE_RECEITA ADD TIPO INT NOT NULL DEFAULT(0) "
    Banco.Execute sql
    
    sql = " ALTER TABLE [PACIENTE_RECEITA] DROP CONSTRAINT [PK_PACIENTE_RECEITA] "
    Banco.Execute sql
    
    sql = " ALTER TABLE [PACIENTE_RECEITA] ADD  CONSTRAINT [PK_PACIENTE_RECEITA] PRIMARY KEY CLUSTERED ( "
    sql = sql & " [REGISTRO] ASC, [Tipo] Asc "
    sql = sql & " )WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]"
    Banco.Execute sql

    sql = " ALTER TABLE TMPREL_PRONTUARIO_ANAMNESE ADD USUARIO VARCHAR(100) NULL "
    Banco.Execute sql
    
    sql = " ALTER TABLE TMPREL_PRONTUARIO_ANAMNESE ADD DATA DATETIME NULL "
    Banco.Execute sql
    
    sql = " ALTER TABLE TMPREL_PRONTUARIO_EXAME ADD USUARIO VARCHAR(100) NULL "
    Banco.Execute sql
    
    sql = " ALTER TABLE TMPREL_PRONTUARIO_EXAME ADD DATA DATETIME NULL "
    Banco.Execute sql
    
    sql = " ALTER TABLE INTERNOPROCEDIMENTO ADD MEDICOSOL INT "
    Banco.Execute sql
    
    sql = " ALTER TABLE AMBULATORIALPROCEDIMENTO ADD MEDICOSOL INT "
    Banco.Execute sql
    
    
    sql = " ALTER TABLE TMP_PRESCRICAOELETRONICA_REPETIR ADD NECESSARIO INT NULL "
    Banco.Execute sql
    
    sql = " ALTER TABLE MENU ADD [HIERARQUIANOVA] varchar(10)"
    Banco.Execute sql

    sql = " ALTER TABLE MENU ADD [NOMESUBALTERADO] varchar(50)"
    Banco.Execute sql
    
    sql = " insert into menu(MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,"
    sql = sql & " ATIVADO,MODULO,HIERARQUIA,NOMESUBAUX,NIVELVISIBILIDADE,HIERARQUIANOVA,NOMESUBALTERADO) "
    sql = sql & " SELECT MAX(MENU)+1 AS MENU ,'Prescrição Eletrônica','mnuMed_Pel','IMPORTAÇÃO','mnuMed_Pel',"
    sql = sql & " 0,1,'0602000000','mnuMed_Pel',1,'0202000000','mnuCli_Pel'"
    sql = sql & " FROM MENU"
    Banco.Execute sql
    
    If Layout = 52 Then
            sql = " UPDATE MENU SET ATIVADO = 1 WHERE NOMESUBNOVO = 'mnuMed_Pel' "
            Banco.Execute sql
    End If
    
    sql = " ALTER TABLE TMPAMBULATORIAL ADD RESPOSTAELEGIBILIDADE CHAR(3) NULL "
    Banco.Execute sql
    
    sql = " ALTER TABLE TMPINTERNACAO ADD RESPOSTAELEGIBILIDADE CHAR(3) NULL "
    Banco.Execute sql
    
    sql = "ALTER TABLE PRODUTOCONVENIO ADD DESPESA INT "
    Banco.Execute sql
    
    sql = "CREATE TABLE [USUARIOPERMISSAOFORM]( " & Chr(13)
    sql = sql & "[USUARIO] [int] NOT NULL, " & Chr(13)
    sql = sql & "[MENU] [int] NOT NULL, " & Chr(13)
    sql = sql & "[INCLUIR] [int] NULL, " & Chr(13)
    sql = sql & "[ALTERAR] [int] NULL, " & Chr(13)
    sql = sql & "[EXCLUIR] [int] NULL, " & Chr(13)
    sql = sql & "[ATUALIZACAO] [varchar](50) NULL, " & Chr(13)
    sql = sql & "CONSTRAINT [PK_USUARIOPERMISSAOFORM] PRIMARY KEY NONCLUSTERED " & Chr(13)
    sql = sql & "( " & Chr(13)
    sql = sql & "    [USUARIO] ASC, " & Chr(13)
    sql = sql & "    [Menu] Asc " & Chr(13)
    sql = sql & ")) ON [PRIMARY]"
    Banco.Execute sql
    
    sql = "ALTER TABLE [dbo].[USUARIOPERMISSAOFORM]  ADD  CONSTRAINT [FK_USUARIOPERMISSAOFORM_MENU] FOREIGN KEY([MENU]) REFERENCES [dbo].[MENU] ([MENU])"
    Banco.Execute sql
    
    sql = "ALTER TABLE MENU ADD PERMISSAOFORM INT"
    Banco.Execute sql
    
    If Layout = 48 Then
        sql = " ALTER TABLE INTERNOPROCEDIMENTO ADD PROCEDIMENTONOMETUSS VARCHAR(250) NULL "
        Banco.Execute sql
        
        sql = " ALTER TABLE INTERNOPROCEDIMENTO ADD PROCEDIMENTOTUSS INT NULL "
        Banco.Execute sql
        
        sql = ""
        sql = sql & "ALTER VIEW [dbo].[v_relatorio_boletiminternacao] " & vbCrLf
        sql = sql & "AS " & vbCrLf
        sql = sql & "  SELECT " & vbCrLf
        sql = sql & "  I.REGISTRO                                                        AS ""REGISTRO"", " & vbCrLf
        sql = sql & "  F.FICHA                                                           AS ""PRONTUARIO"", " & vbCrLf
        sql = sql & "  F.NOME                                                            AS ""PACIENTE"", " & vbCrLf
        sql = sql & "  ISNULL(CONVERT(VARCHAR, F.NASCIMENTO, 101), '    /    /        ') AS ""NASCIMENTO"", " & vbCrLf
        sql = sql & "  ISNULL(F.RG, '')                                                  AS ""RG"", " & vbCrLf
        sql = sql & "  ISNULL(F.CPF, '')                                                 AS ""CPF"", " & vbCrLf
        sql = sql & "  F.TELEFONE                                                        AS ""TELEFONE"", " & vbCrLf
        sql = sql & "  F.ENDERECO                                                        AS ""ENDERECO"", " & vbCrLf
        sql = sql & "  I.DATAINTERNACAO                                                  AS ""INTERNACAO"", " & vbCrLf
        sql = sql & "  ''                                                                AS ""QUARTO"", " & vbCrLf
        sql = sql & "  R.RESPONSAVEL                                                     AS ""ID_RESP"", " & vbCrLf
        sql = sql & "  ISNULL(R.NOME, '')                                                AS ""NOME_RESP"", " & vbCrLf
        sql = sql & "  ''                                                                AS ""ENDERECO_RESP"", " & vbCrLf
        sql = sql & "  R.PARENTESCO                                                      AS ""PARENTESCO"", " & vbCrLf
        sql = sql & "  R.PROFISSAO                                                       AS ""PROFISSAO_RESP"", " & vbCrLf
        sql = sql & "  R.RG                                                              AS ""RG_RESP"", " & vbCrLf
        sql = sql & "  R.CPF                                                             AS ""CPF_RESP"", " & vbCrLf
        sql = sql & "  R.CEP                                                             AS ""CEP_RESP"", " & vbCrLf
        sql = sql & "  R.TELEFONE                                                        AS ""TELEFONE_RESP"", " & vbCrLf
        sql = sql & "  C.DESCRICAO                                                       AS ""ID_CONVENIO"", " & vbCrLf
        sql = sql & "  ''                                                                AS ""ID_PLANO"", " & vbCrLf
        sql = sql & "  ISNULL(I.CARTEIRINHA, '')                                         AS ""MATRICULA"", " & vbCrLf
        sql = sql & "  ISNULL(CONVERT(VARCHAR, I.VALIDADE, 101), '    /    /        ')   AS ""VALIDADE"", " & vbCrLf
        sql = sql & "  M.MEDICO                                                          AS ""CD_MEDICO"", " & vbCrLf
        sql = sql & "  M.CRM                                                             AS ""CRM"", " & vbCrLf
        sql = sql & "  M.NOME                                                            AS ""NOME_MEDICO"", " & vbCrLf
        sql = sql & "  E.DESCRICAO                                                       AS ""ESPECIALIDADE"" " & vbCrLf
        sql = sql & "  FROM   FICHAS F " & vbCrLf
        sql = sql & "         INNER JOIN INTERNO I " & vbCrLf
        sql = sql & "                 ON F.FICHA = I.FICHA " & vbCrLf
        sql = sql & "         LEFT JOIN RESPONSAVEL R " & vbCrLf
        sql = sql & "                ON I.RESPONSAVEL = R.RESPONSAVEL " & vbCrLf
        sql = sql & "         LEFT JOIN CONVENIOS C " & vbCrLf
        sql = sql & "                ON I.CONVENIO = C.CONVENIO " & vbCrLf
        sql = sql & "         LEFT JOIN PLANOCONVENIO P " & vbCrLf
        sql = sql & "                ON P.CONVENIO = I.CONVENIO " & vbCrLf
        sql = sql & "                   AND P.PLANO = I.PLANOCONVENIOTIPO " & vbCrLf
        sql = sql & "         LEFT JOIN MEDICOS M " & vbCrLf
        sql = sql & "                ON I.MEDICO = M.MEDICO " & vbCrLf
        sql = sql & "         LEFT JOIN ESPMEDICA E " & vbCrLf
        sql = sql & "                ON E.ESPECMEDICA = M.ESPECMEDICA"
        Banco.Execute sql
        
    End If
    
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes082013()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '082013'"
   Banco.Execute sql
 
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD TESTE_TUBOSECO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_NASCIDOS ADD TESTE_LOTE_ANO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE TESTEPEZINHOTIPO ADD QTDEREGISTROLOTE INT"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[USUARIOPERMISSAOFORM]( " & Chr(13)
   sql = sql & " [USUARIO] [int] NOT NULL, " & Chr(13)
   sql = sql & " [MENU] [int] NOT NULL, " & Chr(13)
   sql = sql & " [INCLUIR] [int] NULL, " & Chr(13)
   sql = sql & " [ALTERAR] [int] NULL, " & Chr(13)
   sql = sql & " [EXCLUIR] [int] NULL, " & Chr(13)
   sql = sql & " [ATUALIZACAO] [varchar](50) NULL, " & Chr(13)
   sql = sql & " CONSTRAINT [PK_USUARIOPERMISSAOFORM] PRIMARY KEY NONCLUSTERED " & Chr(13)
   sql = sql & " ( " & Chr(13)
   sql = sql & "   [USUARIO] ASC, " & Chr(13)
   sql = sql & "   [menu] Asc " & Chr(13)
   sql = sql & " ))"
   Banco.Execute sql
   
   sql = "ALTER TABLE [dbo].[USUARIOPERMISSAOFORM]  ADD  CONSTRAINT [FK_USUARIOPERMISSAOFORM_MENU] FOREIGN KEY([MENU]) REFERENCES [dbo].[MENU] ([MENU])"
   Banco.Execute sql
   
   sql = "ALTER TABLE MENU ADD PERMISSAOFORM INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_DADOS_OBSTETRICO ADD DADOS_OB_AFU_NUM MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_DADOS_OBSTETRICO ADD VDRL3TRIMESTRE MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_DADOS_OBSTETRICO ADD GLICEMIA1TRIMESTRE MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO_DADOS_OBSTETRICO ADD GLICEMIA3TRIMESTRE MONEY"
   Banco.Execute sql
   
   sql = " ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_AMB ADD IDIMPRESSAO VARCHAR(100) NULL "
   Banco.Execute sql
   
   sql = " ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT ADD IDIMPRESSAO VARCHAR(100) NULL "
   Banco.Execute sql
   
   sql = " ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_EXT ADD IDIMPRESSAO VARCHAR(100) NULL "
   Banco.Execute sql
   
   
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes092013()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '092013'"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD DIETATIPOHIPOCALORICA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD DIETATIPOHIPOPROTEICA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD DIETATIPOHIPERHIPER INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD DIETATIPOPASTOSA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD DIETATIPOBRANDA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD DIETAQUANTIDADE INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD CONSTRAINT DF_Interno_DietaQuantidade  DEFAULT ((1)) FOR DIETAQUANTIDADE "
   Banco.Execute sql
   
   sql = "ALTER TABLE INTERNO ADD DIETADATAALTERACAO DATETIME"
   Banco.Execute sql
       
   sql = "ALTER TABLE MOVIM_INT_TMP_SUS ADD CENTROCUSTO_INTERNADO INT"
   Banco.Execute sql
   
   sql = "ALTER PROCEDURE SP_RECUPERA_MOVIMENTOCONTASUS_INTERNO " & Chr(13)
   sql = sql & " @REGISTRO   AS INT, " & Chr(13)
   sql = sql & " @NUMEROAIH  AS CHAR(15), " & Chr(13)
   sql = sql & "    @IP      AS VARCHAR(100) " & Chr(13)
   sql = sql & " AS " & Chr(13)
   sql = sql & " DELETE FROM MOVIM_INT_TMP_SUS " & Chr(13)
   sql = sql & " WHERE IP = @IP " & Chr(13)
   sql = sql & "    INSERT INTO MOVIM_INT_TMP_SUS ( " & Chr(13)
   sql = sql & "       REGISTRO, SEQUENCIA, DATACONSUMO, TIPOLANCAMENTO, SEQUENCIAFILMEPROCEDIMENTO, " & Chr(13)
   sql = sql & "       TIPOPAGAMENTO, PROCEDIMENTONOME, PROCEDIMENTO, QUANTIDADE, UNIDADEFATURAMENTO, NOTA, " & Chr(13)
   sql = sql & "       DATA, LOCALCONSUMO, HORA, MEDICO, ATUACAO, TIPOGUIA, ATIVPROF, TIPO, TIPOATO, CENTROCUSTO, " & Chr(13)
   sql = sql & "       ATUALIZACAO, PACOTE, VIAACESSO, FORNECEDOR, TIPOPROCEDIMENTOSUS, MUDANCAPROCEDIMENTO, " & Chr(13)
   sql = sql & "       SUS_PROCEDIMENTOCIRURGIAMULTIPLA, VALORUNITARIO, TIPOPROFISSIONAL, NOMEATUACAO, HONORARIOVIDEO, " & Chr(13)
   sql = sql & "       GUIAINTERNA, MEDICOEXECUTANTE, GUIAAUTORIZACAO, NUMEROAIH, SEQUENCIALANCAMENTOMUDANCAPROCEDIMENTO, " & Chr(13)
   sql = sql & "       PROCEDIMENTOREALIZADO, ESPECIALIDADE, ESPECIALIDADENOME, PORCENTAGEMHOSPITAL1, PORCENTAGEMHOSPITAL2, " & Chr(13)
   sql = sql & "       PORCENTAGEMHOSPITAL3, PORCENTAGEMHOSPITAL4, PORCENTAGEMHOSPITAL5, SEQUENCIACIRURGIA, GUIATISS, " & Chr(13)
   sql = sql & "       CNESLANCAMENTO, CBOMEDICO, APURARVALOR,   SEQUENCIAEQUIPE, NOTA_LOTE, NOTA_NUMEROSERIE, " & Chr(13)
   sql = sql & "       NOTA_CGCFABRICANTE, NOTA_REGISTROANVISA, SUS_UTIMOTIVOSAIDA, PERMITELANCARSEMFATURAR, PLANTAO, CENTROCUSTO_INTERNADO, IP) " & Chr(13)
   sql = sql & "     SELECT A.REGISTRO, A.SEQUENCIA, A.DATACONSUMO, A.TIPOLANCAMENTO, A.SEQUENCIAFILMEPROCEDIMENTO, " & Chr(13)
   sql = sql & "       ISNULL(A.TIPOPAGAMENTO,0) AS TIPOPAGAMENTO, " & Chr(13)
   sql = sql & "       (CASE WHEN A.TIPOLANCAMENTO = 1 AND J.DESCRICAO IS NOT NULL THEN J.DESCRICAO ELSE A.PROCEDIMENTONOME END ) AS PROCEDIMENTONOME, " & Chr(13)
   sql = sql & "       (CASE WHEN A.TIPOLANCAMENTO = 1 AND J.DESCRICAO IS NOT NULL THEN J.PRODUTO   ELSE A.PROCEDIMENTO END ) AS PROCEDIMENTO, " & Chr(13)
   sql = sql & "       A.QUANTIDADE/ (CASE WHEN A.TIPOLANCAMENTO=1 THEN ISNULL(B.CONVLANC,1) ELSE 1 END) AS QUANTIDADE, " & Chr(13)
   sql = sql & "       (CASE WHEN A.TIPOLANCAMENTO = 1 THEN B.UNIDADEFATURAMENTO ELSE A.UNIDADEFATURAMENTO END) AS UNIDADEFATURAMENTO , " & Chr(13)
   sql = sql & "       A.NOTA, A.DATA, A.LOCALCONSUMO, A.HORA, A.MEDICO, A.ATUACAO, A.TIPOGUIA, A.ATIVPROF, A.TIPO, A.TIPOATO, " & Chr(13)
   sql = sql & "       A.CENTROCUSTO, A.ATUALIZACAO, A.PACOTE, A.VIAACESSO, FORNECEDOR, " & Chr(13)
   sql = sql & "       ISNULL(A.TIPOPROCEDIMENTOSUS,0) AS TIPOPROCEDIMENTOSUS,A.MUDANCAPROCEDIMENTO, " & Chr(13)
   sql = sql & "       ISNULL(A.SUS_PROCEDIMENTOCIRURGIAMULTIPLA,0) AS SUS_PROCEDIMENTOCIRURGIAMULTIPLA, " & Chr(13)
   sql = sql & "         (CASE WHEN A.TIPOLANCAMENTO=2 THEN ISNULL(A.VALOREXAME,0) + ISNULL(A.VALORHOSPITAL,0) " & Chr(13)
   sql = sql & "               WHEN A.TIPOLANCAMENTO=3 THEN ISNULL(A.VALORPROFISSIONAL,0) " & Chr(13)
   sql = sql & "               WHEN A.TIPOLANCAMENTO=1 AND A.UNIDADEFATURAMENTO=B.UNIDADELANCAMENTO THEN ISNULL(A.VALORUNITARIO,0)*ISNULL(B.CONVLANC,1) " & Chr(13)
   sql = sql & "          ELSE A.VALORUNITARIO END) AS VALORUNITARIO " & Chr(13)
   sql = sql & "       ,ISNULL(G.TIPOPROFISSIONAL,0) AS TIPOPROFISSIONAL,ISNULL(G.DESCRICAO,'') AS NOMEATUACAO, A.HONORARIOVIDEO, " & Chr(13)
   sql = sql & "       A.GUIAINTERNA, A.MEDICOEXECUTANTE, A.GUIAAUTORIZACAO, D.NUMEROAIH, D.SEQUENCIALANCAMENTOMUDANCAPROCEDIMENTO, " & Chr(13)
   sql = sql & "       E.PROCEDIMENTOREALIZADO,  ISNULL(D.ESPECIALIDADE,0) AS ESPECIALIDADE,ISNULL(F.DESCRICAO,0) AS ESPECIALIDADENOME, " & Chr(13)
   sql = sql & "       ISNULL(H.PORCENTAGEMHOSPITAL1,100) AS PORCENTAGEMHOSPITAL1, " & Chr(13)
   sql = sql & "       ISNULL(H.PORCENTAGEMHOSPITAL2,100) AS PORCENTAGEMHOSPITAL2, " & Chr(13)
   sql = sql & "       ISNULL(H.PORCENTAGEMHOSPITAL3,100) AS PORCENTAGEMHOSPITAL3, " & Chr(13)
   sql = sql & "       ISNULL(H.PORCENTAGEMHOSPITAL4,100) AS PORCENTAGEMHOSPITAL4, " & Chr(13)
   sql = sql & "       ISNULL(H.PORCENTAGEMHOSPITAL5, 100) As PORCENTAGEMHOSPITAL5 " & Chr(13)
   sql = sql & "       , ISNULL(A.SEQUENCIA_CIRURGIA,0) AS SEQUENCIA_CIRURGIA, ISNULL(L.GUIA,0) AS GUIATISS, " & Chr(13)
   sql = sql & "       ISNULL(CNESLANCAMENTO,0) AS CNESLANCAMENTO, ISNULL(CBOMEDICO,0) AS CBOMEDICO, " & Chr(13)
   sql = sql & "       ISNULL(APURARVALOR,0) AS APURARVALOR, ISNULL(A.SEQUENCIAEQUIPE,0) AS SEQUENCIAEQUIPE, " & Chr(13)
   sql = sql & "       A.NOTA_LOTE, A.NOTA_NUMEROSERIE, A.NOTA_CGCFABRICANTE, A.NOTA_REGISTROANVISA, " & Chr(13)
   sql = sql & "       ISNULL(A.SUS_UTIMOTIVOSAIDA,0) AS SUS_UTIMOTIVOSAIDA, " & Chr(13)
   sql = sql & "       ISNULL(I.PERMITELANCARSEMFATURAR,0) AS PERMITELANCARSEMFATURAR, A.PLANTAO, A.CENTROCUSTO_INTERNADO, @IP " & Chr(13)
   sql = sql & "     FROM MOVIM_INT A WITH (NOLOCK) " & Chr(13)
   sql = sql & "                   LEFT JOIN PRODUTO      B WITH (NOLOCK) ON A.PROCEDIMENTO = B.PRODUTO " & Chr(13)
   sql = sql & "                                                  AND A.TIPOLANCAMENTO = 1 " & Chr(13)
   sql = sql & "                   LEFT JOIN FARMACIAFAMILIA  C WITH (NOLOCK) ON B.FAMILIA = C.FARMACIAFAMILIA " & Chr(13)
   sql = sql & "                   INNER JOIN DADOSAIH        D WITH (NOLOCK) ON A.REGISTRO=D.REGISTRO " & Chr(13)
   sql = sql & "                                                     AND LTRIM(D.NUMEROAIH) = @NUMEROAIH " & Chr(13)
   sql = sql & "                                                     AND LTRIM(A.NUMEROAIH) = LTRIM(D.NUMEROAIH) " & Chr(13)
   sql = sql & "                   LEFT JOIN COMPLEMENTAR     E WITH (NOLOCK) ON A.REGISTRO=E.REGISTRO " & Chr(13)
   sql = sql & "                                                     AND LTRIM(D.NUMEROAIH)=LTRIM(E.NUMEROAIH) " & Chr(13)
   sql = sql & "                   LEFT JOIN ESPMEDICASUS     F WITH (NOLOCK) ON D.ESPECIALIDADE=F.ESPECMEDICASUS AND F.TIPO=1 " & Chr(13)
   sql = sql & "                   LEFT JOIN SUSINTERNOS      H WITH (NOLOCK) ON E.PROCEDIMENTOREALIZADO=H.CODIGOSUSINTERNO " & Chr(13)
   sql = sql & "                   LEFT JOIN SUS_PROCEDIMENTO I WITH (NOLOCK) ON A.PROCEDIMENTO=I.CODIGOSUS_PROCEDIMENTO " & Chr(13)
   sql = sql & "                                                     AND A.TIPOLANCAMENTO IN(2,3) " & Chr(13)
   sql = sql & "                   LEFT JOIN ATUACOES     G WITH (NOLOCK) ON A.ATUACAO=G.ATUACAO " & Chr(13)
   sql = sql & "                   LEFT JOIN CENTROCUSTO     M WITH (NOLOCK) ON A.CENTROCUSTO=M.CENTROCUSTO " & Chr(13)
   sql = sql & "                   LEFT JOIN PRODUTO      J WITH (NOLOCK) ON B.CODIGOCOMERCIALPADRAO = J.PRODUTO " & Chr(13)
   sql = sql & "                   LEFT JOIN DADOSGUIASECUNDARIO L WITH (NOLOCK) ON A.REGISTRO = L.REGISTRO " & Chr(13)
   sql = sql & "                                                           AND A.GUIAINTERNA = L.GUIAINTERNA " & Chr(13)
   sql = sql & "                                                           AND TIPOREGISTRO = 2 " & Chr(13)
   sql = sql & "     WHERE A.REGISTRO = @REGISTRO " & Chr(13)
   sql = sql & "     AND ISNULL(A.QUANTIDADE,0) > 0 " & Chr(13)
   sql = sql & "     AND A.TIPOLANCAMENTO IN(2,3,4) " & Chr(13)
   sql = sql & "     ORDER BY A.PROCEDIMENTO,A.DATACONSUMO DESC,A.SEQUENCIA DESC " & Chr(13)
   
   Banco.Execute sql
   
    sql = ""
    sql = sql & "IF Object_id('TR_LOG_MOVIMENTOPRODUTODIVERSOS') IS NOT NULL " & vbCrLf
    sql = sql & "  DROP TRIGGER tr_log_movimentoprodutodiversos"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "CREATE TRIGGER tr_log_movimentoprodutodiversos " & vbCrLf
    sql = sql & "ON movimentoprodutodiversos " & vbCrLf
    sql = sql & "WITH encryption " & vbCrLf
    sql = sql & "FOR INSERT, UPDATE, DELETE " & vbCrLf
    sql = sql & "AS " & vbCrLf
    sql = sql & "    DECLARE @TIPOOPERACAO AS INT " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    -- 0 - INCLUSÃO " & vbCrLf
    sql = sql & "    -- 1 - ALTERAÇÃO " & vbCrLf
    sql = sql & "    -- 2 - EXCLUSÃO " & vbCrLf
    sql = sql & "    SET @TIPOOPERACAO = 2 " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    IF EXISTS(SELECT produto " & vbCrLf
    sql = sql & "              FROM   inserted) " & vbCrLf
    sql = sql & "      BEGIN " & vbCrLf
    sql = sql & "          IF NOT EXISTS(SELECT produto " & vbCrLf
    sql = sql & "                        FROM   deleted) " & vbCrLf
    sql = sql & "            SET @TIPOOPERACAO=0 " & vbCrLf
    sql = sql & "          ELSE " & vbCrLf
    sql = sql & "            SET @TIPOOPERACAO=1 " & vbCrLf
    sql = sql & "      END " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    IF @TIPOOPERACAO = 2  " & vbCrLf
    sql = sql & "--EXCLUSÃO " & vbCrLf
    sql = sql & "      BEGIN " & vbCrLf
    sql = sql & "          INSERT INTO savelog..movimentoprodutodiversos " & vbCrLf
    sql = sql & "                      (data, " & vbCrLf
    sql = sql & "                       produto, " & vbCrLf
    sql = sql & "                       nome, " & vbCrLf
    sql = sql & "                       tipomovimento, " & vbCrLf
    sql = sql & "                       quantidade, " & vbCrLf
    sql = sql & "                       local, " & vbCrLf
    sql = sql & "                       ip, " & vbCrLf
    sql = sql & "                       atualizacao, " & vbCrLf
    sql = sql & "                       centrocusto, " & vbCrLf
    sql = sql & "                       frascointeiro, " & vbCrLf
    sql = sql & "                       sembaixa, " & vbCrLf
    sql = sql & "                       chave, " & vbCrLf
    sql = sql & "                       lote, " & vbCrLf
    sql = sql & "                       validadelote, " & vbCrLf
    sql = sql & "                       documento, " & vbCrLf
    sql = sql & "                       fatorestoque, " & vbCrLf
    sql = sql & "                       fornecedor, " & vbCrLf
    sql = sql & "                       fornecedornome, " & vbCrLf
    sql = sql & "                       tipooperacao, " & vbCrLf
    sql = sql & "                       dataoperacao, " & vbCrLf
    sql = sql & "                       valor) " & vbCrLf
    sql = sql & "          SELECT data, " & vbCrLf
    sql = sql & "                 produto, " & vbCrLf
    sql = sql & "                 nome, " & vbCrLf
    sql = sql & "                 tipomovimento, " & vbCrLf
    sql = sql & "                 quantidade, " & vbCrLf
    sql = sql & "                 local, " & vbCrLf
    sql = sql & "                 ip, " & vbCrLf
    sql = sql & "                 atualizacao, " & vbCrLf
    sql = sql & "                 centrocusto, " & vbCrLf
    sql = sql & "                 frascointeiro, " & vbCrLf
    sql = sql & "                 sembaixa, " & vbCrLf
    sql = sql & "                 chave, " & vbCrLf
    sql = sql & "                 lote, " & vbCrLf
    sql = sql & "                 validadelote, " & vbCrLf
    sql = sql & "                 documento, " & vbCrLf
    sql = sql & "                 fatorestoque, " & vbCrLf
    sql = sql & "                 fornecedor, " & vbCrLf
    sql = sql & "                 fornecedornome, " & vbCrLf
    sql = sql & "                 @TIPOOPERACAO, " & vbCrLf
    sql = sql & "                 Getdate(), " & vbCrLf
    sql = sql & "                 valor " & vbCrLf
    sql = sql & "          FROM   deleted " & vbCrLf
    sql = sql & "      END " & vbCrLf
    sql = sql & "    ELSE " & vbCrLf
    sql = sql & "      BEGIN " & vbCrLf
    sql = sql & "          INSERT INTO savelog..movimentoprodutodiversos " & vbCrLf
    sql = sql & "                      (data, " & vbCrLf
    sql = sql & "                       produto, " & vbCrLf
    sql = sql & "                       nome, " & vbCrLf
    sql = sql & "                       tipomovimento, " & vbCrLf
    sql = sql & "                       quantidade, " & vbCrLf
    sql = sql & "                       local, " & vbCrLf
    sql = sql & "                       ip, " & vbCrLf
    sql = sql & "                       atualizacao, " & vbCrLf
    sql = sql & "                       centrocusto, " & vbCrLf
    sql = sql & "                       frascointeiro, " & vbCrLf
    sql = sql & "                       sembaixa, " & vbCrLf
    sql = sql & "                       chave, " & vbCrLf
    sql = sql & "                       lote, " & vbCrLf
    sql = sql & "                       validadelote, " & vbCrLf
    sql = sql & "                       documento, " & vbCrLf
    sql = sql & "                       fatorestoque, " & vbCrLf
    sql = sql & "                       fornecedor, " & vbCrLf
    sql = sql & "                       fornecedornome, " & vbCrLf
    sql = sql & "                       tipooperacao, " & vbCrLf
    sql = sql & "                       dataoperacao, " & vbCrLf
    sql = sql & "                       valor) " & vbCrLf
    sql = sql & "          SELECT data, " & vbCrLf
    sql = sql & "                 produto, " & vbCrLf
    sql = sql & "                 nome, " & vbCrLf
    sql = sql & "                 tipomovimento, " & vbCrLf
    sql = sql & "                 quantidade, " & vbCrLf
    sql = sql & "                 local, " & vbCrLf
    sql = sql & "                 ip, " & vbCrLf
    sql = sql & "                 atualizacao, " & vbCrLf
    sql = sql & "                 centrocusto, " & vbCrLf
    sql = sql & "                 frascointeiro, " & vbCrLf
    sql = sql & "                 sembaixa, " & vbCrLf
    sql = sql & "                 chave, " & vbCrLf
    sql = sql & "                 lote, " & vbCrLf
    sql = sql & "                 validadelote, " & vbCrLf
    sql = sql & "                 documento, " & vbCrLf
    sql = sql & "                 fatorestoque, " & vbCrLf
    sql = sql & "                 fornecedor, " & vbCrLf
    sql = sql & "                 fornecedornome, " & vbCrLf
    sql = sql & "                 @TIPOOPERACAO, " & vbCrLf
    sql = sql & "                 Getdate(), " & vbCrLf
    sql = sql & "                 valor " & vbCrLf
    sql = sql & "          FROM   inserted " & vbCrLf
    sql = sql & "      END "
    
    Banco.Execute sql

   sql = ""
    sql = sql & "IF Object_id('TR_PSICOTROPICO_MOVIMENTOPRODUTODIVERSOS') IS NOT NULL " & vbCrLf
    sql = sql & "  DROP TRIGGER tr_psicotropico_movimentoprodutodiversos"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "CREATE TRIGGER tr_psicotropico_movimentoprodutodiversos " & vbCrLf
    sql = sql & "ON movimentoprodutodiversos " & vbCrLf
    sql = sql & "WITH encryption " & vbCrLf
    sql = sql & "FOR INSERT, UPDATE, DELETE " & vbCrLf
    sql = sql & "AS " & vbCrLf
    sql = sql & "    --TIPO=0/INSERT/1=UPDATE/2=DELETE " & vbCrLf
    sql = sql & "    DECLARE @TIPO            AS INT, " & vbCrLf
    sql = sql & "            @QUANTIDADEATUAL AS MONEY, " & vbCrLf
    sql = sql & "            @QUANTIDADENOVA  AS MONEY " & vbCrLf
    sql = sql & "    DECLARE @SQL AS VARCHAR(255) " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    SET @TIPO=0 " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    IF EXISTS(SELECT * " & vbCrLf
    sql = sql & "              FROM   deleted) " & vbCrLf
    sql = sql & "      BEGIN " & vbCrLf
    sql = sql & "          SET @TIPO=2 " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "          IF EXISTS(SELECT * " & vbCrLf
    sql = sql & "                    FROM   inserted) " & vbCrLf
    sql = sql & "            SET @TIPO=1 " & vbCrLf
    sql = sql & "      END " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    SELECT @TIPO " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    SELECT @TIPO " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    IF @TIPO = 0 " & vbCrLf
    sql = sql & "      BEGIN " & vbCrLf
    sql = sql & "          INSERT INTO psicotropicomovimento " & vbCrLf
    sql = sql & "                      (produto, " & vbCrLf
    sql = sql & "                       produtonome, " & vbCrLf
    sql = sql & "                       data, " & vbCrLf
    sql = sql & "                       documento, " & vbCrLf
    sql = sql & "                       tipo, " & vbCrLf
    sql = sql & "                       codigo, " & vbCrLf
    sql = sql & "                       nome, " & vbCrLf
    sql = sql & "                       quantidade, " & vbCrLf
    sql = sql & "                       local, " & vbCrLf
    sql = sql & "                       atualizacao, " & vbCrLf
    sql = sql & "                       lote, " & vbCrLf
    sql = sql & "                       validadelote) " & vbCrLf
    sql = sql & "          SELECT A.produto, " & vbCrLf
    sql = sql & "                 B.descricao, " & vbCrLf
    sql = sql & "                 A.data, " & vbCrLf
    sql = sql & "                 A.documento, " & vbCrLf
    sql = sql & "                 A.tipomovimento, " & vbCrLf
    sql = sql & "                 A.fornecedor, " & vbCrLf
    sql = sql & "                 A.fornecedornome, " & vbCrLf
    sql = sql & "                 ( A.quantidade * Isnull(B.fatorestoque, 1) ), " & vbCrLf
    sql = sql & "                 'MOVIMENTOPRODUTODIVERSOS', " & vbCrLf
    sql = sql & "                 Getdate(), " & vbCrLf
    sql = sql & "                 A.lote, " & vbCrLf
    sql = sql & "                 A.validadelote " & vbCrLf
    sql = sql & "          FROM   inserted A " & vbCrLf
    sql = sql & "                 INNER JOIN produto B " & vbCrLf
    sql = sql & "                         ON A.produto = B.produto " & vbCrLf
    sql = sql & "                 INNER JOIN farmaciafamilia C " & vbCrLf
    sql = sql & "                         ON C.farmaciafamilia = B.familia " & vbCrLf
    sql = sql & "                 CROSS JOIN parametro P " & vbCrLf
    sql = sql & "          WHERE  C.tipo = 3 " & vbCrLf
    sql = sql & "                 AND B.centrocusto = P.farmaciacc " & vbCrLf
    sql = sql & "       " & vbCrLf
    sql = sql & "-- PARA NÃO SAIR AS FAMILIAS DO ALMOXARIFADO " & vbCrLf
    sql = sql & "      END " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    IF @TIPO = 1 " & vbCrLf
    sql = sql & "      BEGIN " & vbCrLf
    sql = sql & "          SELECT @QUANTIDADEATUAL = quantidade " & vbCrLf
    sql = sql & "          FROM   deleted " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "          SELECT @QUANTIDADENOVA = quantidade " & vbCrLf
    sql = sql & "          FROM   inserted " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "          IF @QUANTIDADEATUAL <> @QUANTIDADENOVA " & vbCrLf
    sql = sql & "            BEGIN " & vbCrLf
    sql = sql & "                INSERT INTO psicotropicomovimento " & vbCrLf
    sql = sql & "                            (produto, " & vbCrLf
    sql = sql & "                             produtonome, " & vbCrLf
    sql = sql & "                             data, " & vbCrLf
    sql = sql & "                             documento, " & vbCrLf
    sql = sql & "                             tipo, " & vbCrLf
    sql = sql & "                             codigo, " & vbCrLf
    sql = sql & "                             nome, " & vbCrLf
    sql = sql & "                             quantidade, " & vbCrLf
    sql = sql & "                             local, " & vbCrLf
    sql = sql & "                             atualizacao, " & vbCrLf
    sql = sql & "                             lote, " & vbCrLf
    sql = sql & "                             validadelote) " & vbCrLf
    sql = sql & "                SELECT A.produto, " & vbCrLf
    sql = sql & "                       B.descricao, " & vbCrLf
    sql = sql & "                       A.data, " & vbCrLf
    sql = sql & "                       A.documento, " & vbCrLf
    sql = sql & "                       ( CASE " & vbCrLf
    sql = sql & "                           WHEN @QUANTIDADEATUAL < @QUANTIDADENOVA THEN 0 " & vbCrLf
    sql = sql & "                           ELSE 1 " & vbCrLf
    sql = sql & "                         END ), " & vbCrLf
    sql = sql & "                       A.fornecedor, " & vbCrLf
    sql = sql & "                       A.fornecedornome, " & vbCrLf
    sql = sql & "                       Abs(( @QUANTIDADEATUAL * Isnull(B.fatorestoque, 1) ) - ( " & vbCrLf
    sql = sql & "                           @QUANTIDADENOVA * Isnull(B.fatorestoque, 1) )), " & vbCrLf
    sql = sql & "                       'MOVIMENTOPRODUTODIVERSOS', " & vbCrLf
    sql = sql & "                       Getdate(), " & vbCrLf
    sql = sql & "                       A.lote, " & vbCrLf
    sql = sql & "                       A.validadelote " & vbCrLf
    sql = sql & "                FROM   inserted A " & vbCrLf
    sql = sql & "                       INNER JOIN produto B " & vbCrLf
    sql = sql & "                               ON A.produto = B.produto " & vbCrLf
    sql = sql & "                       INNER JOIN farmaciafamilia C " & vbCrLf
    sql = sql & "                               ON C.farmaciafamilia = B.familia " & vbCrLf
    sql = sql & "                       CROSS JOIN parametro P " & vbCrLf
    sql = sql & "                WHERE  C.tipo = 3 " & vbCrLf
    sql = sql & "                       AND B.centrocusto = P.farmaciacc " & vbCrLf
    sql = sql & "             " & vbCrLf
    sql = sql & "-- PARA NÃO SAIR AS FAMILIAS DO ALMOXARIFADO " & vbCrLf
    sql = sql & "            END " & vbCrLf
    sql = sql & "      END " & vbCrLf
    sql = sql & " " & vbCrLf
    sql = sql & "    IF @TIPO = 2 " & vbCrLf
    sql = sql & "      BEGIN " & vbCrLf
    sql = sql & "          INSERT INTO psicotropicomovimento " & vbCrLf
    sql = sql & "                      (produto, " & vbCrLf
    sql = sql & "                       produtonome, " & vbCrLf
    sql = sql & "                       data, " & vbCrLf
    sql = sql & "                       documento, " & vbCrLf
    sql = sql & "                       tipo, " & vbCrLf
    sql = sql & "                       codigo, " & vbCrLf
    sql = sql & "                       nome, " & vbCrLf
    sql = sql & "                       quantidade, " & vbCrLf
    sql = sql & "                       local, " & vbCrLf
    sql = sql & "                       atualizacao, " & vbCrLf
    sql = sql & "                       lote, " & vbCrLf
    sql = sql & "                       validadelote) " & vbCrLf
    sql = sql & "          SELECT A.produto, " & vbCrLf
    sql = sql & "                 B.descricao, " & vbCrLf
    sql = sql & "                 A.data, " & vbCrLf
    sql = sql & "                 A.documento, " & vbCrLf
    sql = sql & "                 ( CASE " & vbCrLf
    sql = sql & "                     WHEN A.tipomovimento = 1 THEN 0 " & vbCrLf
    sql = sql & "                     ELSE 1 " & vbCrLf
    sql = sql & "                   END ), " & vbCrLf
    sql = sql & "                 A.fornecedor, " & vbCrLf
    sql = sql & "                 A.fornecedornome, " & vbCrLf
    sql = sql & "                 ( A.quantidade * Isnull(B.fatorestoque, 1) ), " & vbCrLf
    sql = sql & "                 'MOVIMENTOPRODUTODIVERSOS', " & vbCrLf
    sql = sql & "                 Getdate(), " & vbCrLf
    sql = sql & "                 A.lote, " & vbCrLf
    sql = sql & "                 A.validadelote " & vbCrLf
    sql = sql & "          FROM   deleted A " & vbCrLf
    sql = sql & "                 INNER JOIN produto B " & vbCrLf
    sql = sql & "                         ON A.produto = B.produto " & vbCrLf
    sql = sql & "                 INNER JOIN farmaciafamilia C " & vbCrLf
    sql = sql & "                         ON C.farmaciafamilia = B.familia " & vbCrLf
    sql = sql & "                 CROSS JOIN parametro P " & vbCrLf
    sql = sql & "          WHERE  C.tipo = 3 " & vbCrLf
    sql = sql & "                 AND B.centrocusto = P.farmaciacc " & vbCrLf
    sql = sql & "       " & vbCrLf
    sql = sql & "-- PARA NÃO SAIR AS FAMILIAS DO ALMOXARIFADO " & vbCrLf
    sql = sql & "      END "
    
    Banco.Execute sql
   
   
   sql = "ALTER TABLE MOVIM_AMB_TMP_SUS ADD CENTROCUSTO_INTERNADO INT "
   Banco.Execute sql
   
   sql = "ALTER PROCEDURE SP_RECUPERA_MOVIMENTOCONTASUS_AMBULATORIAL " & Chr(13)
   sql = sql & " @REGISTRO   AS INT, " & Chr(13)
   sql = sql & " @IP      AS VARCHAR(100) " & Chr(13)
   sql = sql & " AS " & Chr(13)
   sql = sql & "    DELETE FROM MOVIM_AMB_TMP_SUS " & Chr(13)
   sql = sql & "    WHERE IP = @IP " & Chr(13)
   sql = sql & "    INSERT INTO MOVIM_AMB_TMP_SUS ( " & Chr(13)
   sql = sql & "       REGISTRO, SEQUENCIA, DATACONSUMO, TIPOLANCAMENTO, SEQUENCIAFILMEPROCEDIMENTO, " & Chr(13)
   sql = sql & "       TIPOPAGAMENTO, PROCEDIMENTONOME, PROCEDIMENTO, QUANTIDADE, UNIDADEFATURAMENTO, NOTA, " & Chr(13)
   sql = sql & "       DATA, LOCALCONSUMO, HORA, MEDICO, ATUACAO, TIPOGUIA, ATIVPROF, TIPO, TIPOATO, CENTROCUSTO, " & Chr(13)
   sql = sql & "       ATUALIZACAO, PACOTE, VIAACESSO, FORNECEDOR, VALORUNITARIO, TIPOPROFISSIONAL, NOMEATUACAO, HONORARIOVIDEO, " & Chr(13)
   sql = sql & "       GUIAINTERNA, MEDICOEXECUTANTE, GUIAAUTORIZACAO, GRUPOATENDIMENTO, SEQUENCIA_CIRURGIA, GUIATISS, " & Chr(13)
   sql = sql & "       CNESLANCAMENTO, CBOMEDICO, APURARVALOR,   SEQUENCIAEQUIPE, SUS_FATURAR, PERMITELANCARSEMFATURAR, PLANTAO, CENTROCUSTO_INTERNADO, IP) " & Chr(13)
   sql = sql & "     SELECT A.REGISTRO, A.SEQUENCIA, A.DATACONSUMO, A.TIPOLANCAMENTO, A.SEQUENCIAFILMEPROCEDIMENTO, " & Chr(13)
   sql = sql & "       ISNULL(A.TIPOPAGAMENTO,0) AS TIPOPAGAMENTO, " & Chr(13)
   sql = sql & "       (CASE WHEN A.TIPOLANCAMENTO=1 AND J.DESCRICAO IS NOT NULL THEN J.DESCRICAO      ELSE A.PROCEDIMENTONOME END ) AS PROCEDIMENTONOME, " & Chr(13)
   sql = sql & "       (CASE WHEN A.TIPOLANCAMENTO=1 AND J.DESCRICAO IS NOT NULL THEN J.PRODUTO      ELSE A.PROCEDIMENTO END ) AS PROCEDIMENTO, " & Chr(13)
   sql = sql & "       A.QUANTIDADE/ (CASE WHEN A.TIPOLANCAMENTO=1 THEN ISNULL(B.CONVLANC,1) ELSE 1 END) AS QUANTIDADE, " & Chr(13)
   sql = sql & "       (CASE WHEN A.TIPOLANCAMENTO=1 THEN B.UNIDADEFATURAMENTO ELSE A.UNIDADEFATURAMENTO END) AS UNIDADEFATURAMENTO , " & Chr(13)
   sql = sql & "            A.NOTA, A.DATA, A.LOCALCONSUMO, A.HORA, A.MEDICO, A.ATUACAO, A.TIPOGUIA, A.ATIVPROF, A.TIPO, A.TIPOATO, A.CENTROCUSTO, " & Chr(13)
   sql = sql & "       A.ATUALIZACAO, A.PACOTE, A.VIAACESSO, FORNECEDOR, " & Chr(13)
   sql = sql & "       (CASE WHEN A.TIPOLANCAMENTO=1 THEN ISNULL(A.VALORUNITARIO,0)*ISNULL(B.CONVLANC,1) ELSE A.VALORUNITARIO END ) AS VALORUNITARIO, " & Chr(13)
   sql = sql & "       ISNULL(G.TIPOPROFISSIONAL,0) AS TIPOPROFISSIONAL, ISNULL(G.DESCRICAO,'') AS NOMEATUACAO, A.HONORARIOVIDEO, " & Chr(13)
   sql = sql & "       A.GUIAINTERNA, A.MEDICOEXECUTANTE, A.GUIAAUTORIZACAO, GRUPOATENDIMENTO  , ISNULL(A.SEQUENCIA_CIRURGIA,0) AS SEQUENCIA_CIRURGIA, " & Chr(13)
   sql = sql & "       ISNULL(L.GUIA,0) AS GUIATISS, ISNULL(CNESLANCAMENTO,0) AS CNESLANCAMENTO, ISNULL(CBOMEDICO,0) AS CBOMEDICO, " & Chr(13)
   sql = sql & "       ISNULL(APURARVALOR,0) AS APURARVALOR, ISNULL(A.SEQUENCIAEQUIPE,0) AS SEQUENCIAEQUIPE, " & Chr(13)
   sql = sql & "       ISNULL(A.SUS_FATURAR,0) AS  SUS_FATURAR, ISNULL(I.PERMITELANCARSEMFATURAR,0) AS  PERMITELANCARSEMFATURAR, A.PLANTAO, A.CENTROCUSTO_INTERNADO, @IP " & Chr(13)
   sql = sql & "     FROM MOVIM_AMB A WITH (NOLOCK) " & Chr(13)
   sql = sql & "                   LEFT JOIN PRODUTO         B WITH (NOLOCK) ON A.PROCEDIMENTO = B.PRODUTO " & Chr(13)
   sql = sql & "                                                  AND A.TIPOLANCAMENTO = 1 " & Chr(13)
   sql = sql & "                   LEFT JOIN FARMACIAFAMILIA     C WITH (NOLOCK) ON B.FAMILIA = C.FARMACIAFAMILIA " & Chr(13)
   sql = sql & "                   LEFT JOIN SUS_PROCEDIMENTO    I WITH (NOLOCK) ON A.PROCEDIMENTO=I.CODIGOSUS_PROCEDIMENTO " & Chr(13)
   sql = sql & "                                                           AND A.TIPOLANCAMENTO IN(2,3) " & Chr(13)
   sql = sql & "                   LEFT JOIN ATUACOES        G WITH (NOLOCK) ON A.ATUACAO=G.ATUACAO " & Chr(13)
   sql = sql & "                   LEFT JOIN CENTROCUSTO        M WITH (NOLOCK) ON A.CENTROCUSTO=M.CENTROCUSTO " & Chr(13)
   sql = sql & "                   LEFT JOIN PRODUTO         J WITH (NOLOCK) ON B.CODIGOCOMERCIALPADRAO = J.PRODUTO " & Chr(13)
   sql = sql & "                   LEFT JOIN DADOSGUIASECUNDARIO L WITH (NOLOCK) ON A.REGISTRO = L.REGISTRO " & Chr(13)
   sql = sql & "                                                         AND A.GUIAINTERNA = L.GUIAINTERNA " & Chr(13)
   sql = sql & "                                                     AND TIPOREGISTRO = 1 " & Chr(13)
   sql = sql & "     WHERE A.REGISTRO = @REGISTRO " & Chr(13)
   sql = sql & "     AND ISNULL(A.QUANTIDADE,0) > 0 " & Chr(13)
   sql = sql & "     AND A.TIPOLANCAMENTO IN(2,3,4) " & Chr(13)
   sql = sql & "     ORDER BY A.PROCEDIMENTO,A.DATACONSUMO DESC,A.SEQUENCIA DESC " & Chr(13)
   
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_ANAMNESE ADD USUARIO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_ANAMNESE ADD TUTOR INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_ANAMNESE ADD DATA DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_ANAMNESE ADD HORA DATETIME"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD PRESCRICAO_PERIODO_INICIAL_MANHA CHAR(5)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD PRESCRICAO_PERIODO_FINAL_MANHA CHAR(5)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD PRESCRICAO_PERIODO_INICIAL_TARDE CHAR(5)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD PRESCRICAO_PERIODO_FINAL_TARDE CHAR(5)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD PRESCRICAO_PERIODO_INICIAL_NOITE CHAR(5)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD PRESCRICAO_PERIODO_FINAL_NOITE CHAR(5)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRESCRICAOELETRONICAPERIODO_AMB ADD PRESCRICAOREPETIDA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRESCRICAOELETRONICAPERIODO_INT ADD PRESCRICAOREPETIDA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRESCRICAOELETRONICAPERIODO_EXT ADD PRESCRICAOREPETIDA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRESCRICAOELETRONICAPERIODO_LOG ADD PRESCRICAOREPETIDA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD ENFERMAGEMCC INT"
   Banco.Execute sql
   
   'CI=16335 RAPHAEL 09/09/2013 15:46
   'CRIAÇÃO DOS CAMPOS, NECESSÁRIOS NA TELA DE REGISTRO
   sql = " ALTER TABLE INTERNOPROCEDIMENTO ADD MEDICOSOL INT "
   Banco.Execute sql
    
   sql = " ALTER TABLE AMBULATORIALPROCEDIMENTO ADD MEDICOSOL INT "
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPRELACAONASCIDOS ADD OBSTETRA VARCHAR(100) "
   Banco.Execute sql
   
   sql = " ALTER TABLE TMPINTERNACAO ALTER COLUMN IP VARCHAR(155) "
   Banco.Execute sql
   
   sql = "ALTER TABLE DET_SERV ADD PLANTAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD SEQUENCIAL_RPA INT"
   Banco.Execute sql
   
    sql = ""
    sql = sql & "IF OBJECT_ID('V_RELATORIO_ETIQUETAPRODUTO') IS NOT NULL " & vbCrLf
    sql = sql & "  DROP VIEW v_relatorio_etiquetaproduto"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "CREATE VIEW [DBO].[V_RELATORIO_ETIQUETAPRODUTO] " & vbCrLf
    sql = sql & "AS " & vbCrLf
    sql = sql & "  SELECT DISTINCT 'PRODUTO'                                               AS " & vbCrLf
    sql = sql & "                  TIPO, " & vbCrLf
    sql = sql & "                  A.PRODUTO                                               AS " & vbCrLf
    sql = sql & "                     PRODUTO, " & vbCrLf
    sql = sql & "                  A.DESCRICAO                                             AS " & vbCrLf
    sql = sql & "                     DESCRICAO, " & vbCrLf
    sql = sql & "                  ISNULL(B.LOTE, 'S/ LOTE')                               AS " & vbCrLf
    sql = sql & "                  LOTE, " & vbCrLf
    sql = sql & "                  ISNULL(CONVERT(VARCHAR, B.VALIDADE, 103), '01/01/1900') AS " & vbCrLf
    sql = sql & "                     VALIDADE, " & vbCrLf
    sql = sql & "                  ISNULL(ISNULL(C.SEQUENCIA, A.PRODUTO), 0)               AS " & vbCrLf
    sql = sql & "                     SEQUENCIA, " & vbCrLf
    sql = sql & "                  ISNULL(CAST(D.CRF AS CHAR), '')                         AS CRF " & vbCrLf
    sql = sql & "                  , " & vbCrLf
    sql = sql & "                  ISNULL(D.FARMACEUTICO, '') " & vbCrLf
    sql = sql & "                  AS " & vbCrLf
    sql = sql & "                     FARMACEUTICO " & vbCrLf
    sql = sql & "  FROM   PRODUTO A WITH (nolock) " & vbCrLf
    sql = sql & "         INNER JOIN PRODUTOCENTROCUSTOLOTE B WITH (nolock) " & vbCrLf
    sql = sql & "                 ON A.PRODUTO = B.PRODUTO " & vbCrLf
    sql = sql & "         INNER JOIN PRODUTO_LOTE_ETIQUETA C WITH (nolock) " & vbCrLf
    sql = sql & "                 ON C.PRODUTO = B.PRODUTO " & vbCrLf
    sql = sql & "                    AND C.LOTE = B.LOTE " & vbCrLf
    sql = sql & "                    AND C.VALIDADELOTE = B.VALIDADE " & vbCrLf
    sql = sql & "         CROSS JOIN PARAMETROETIQUETA D " & vbCrLf
    sql = sql & "  UNION ALL " & vbCrLf
    sql = sql & "  SELECT DISTINCT 'KIT'                           AS TIPO, " & vbCrLf
    sql = sql & "                  A.KITPROCEDIMENTO               AS PRODUTO, " & vbCrLf
    sql = sql & "                  A.DESCRICAO                     AS DESCRICAO, " & vbCrLf
    sql = sql & "                  'S/ LOTE'                       AS LOTE, " & vbCrLf
    sql = sql & "                  '01/01/1900'                    AS VALIDADE, " & vbCrLf
    sql = sql & "                  A.KITPROCEDIMENTO               AS SEQUENCIA, " & vbCrLf
    sql = sql & "                  ISNULL(CAST(D.CRF AS CHAR), '') AS CRF, " & vbCrLf
    sql = sql & "                  ISNULL(D.FARMACEUTICO, '')      AS FARMACEUTICO " & vbCrLf
    sql = sql & "  FROM   KITPROCEDIMENTO A WITH (nolock) " & vbCrLf
    sql = sql & "         CROSS JOIN PARAMETROETIQUETA D "
    Banco.Execute sql
    
    sql = " ALTER TABLE MEDICOHORARIO ADD [SEQUENCIA] [int] IDENTITY(1,1) NOT NULL "
    Banco.Execute sql
        
    sql = ""
    sql = sql & "ALTER TABLE FORNECEDORES " & vbCrLf
    sql = sql & "  ADD DATAFUNDACAO DATETIME"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "ALTER TABLE FORNECEDORES " & vbCrLf
    sql = sql & "  ADD TIPOPESSOA INT DEFAULT(0)"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "ALTER TABLE FORNECEDORES " & vbCrLf
    sql = sql & "  ADD PRAZOENTREGA INT DEFAULT(0)"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "ALTER TABLE FORNECEDORES " & vbCrLf
    sql = sql & "  ADD VALORMINIMOPEDIDO MONEY DEFAULT(0)"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "ALTER TABLE FORNECEDORES " & vbCrLf
    sql = sql & "  ADD CONDICAOPAGAMENTO VARCHAR(40) NULL"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "ALTER TABLE FORNECEDORES " & vbCrLf
    sql = sql & "  ADD RESPNEGOCIACAO VARCHAR(100) NULL"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "CREATE TABLE FORNPRODTIPO " & vbCrLf
    sql = sql & "  ( " & vbCrLf
    sql = sql & "     FORNECEDOR  INT NOT NULL, " & vbCrLf
    sql = sql & "     PRODUTOTIPO INT NOT NULL " & vbCrLf
    sql = sql & "  )"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "ALTER TABLE FORNPRODTIPO " & vbCrLf
    sql = sql & "  ADD CONSTRAINT fk_fornprodtipo_fornecedor FOREIGN KEY (FORNECEDOR) REFERENCES " & vbCrLf
    sql = sql & "  FORNECEDORES(FORNECEDOR)"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "ALTER TABLE FORNPRODTIPO " & vbCrLf
    sql = sql & "  ADD CONSTRAINT fk_fornprodtipo_produtotipo FOREIGN KEY (PRODUTOTIPO) " & vbCrLf
    sql = sql & "  REFERENCES PRODUTOTIPO(CODIGO)"
    Banco.Execute sql
    
    sql = ""
    sql = sql & "ALTER TABLE FORNPRODTIPO " & vbCrLf
    sql = sql & "  ADD CONSTRAINT pk_fornprodtipo PRIMARY KEY (FORNECEDOR, PRODUTOTIPO)"
    Banco.Execute sql
   
    sql = ""
    sql = sql & "ALTER TABLE COTACAO1 ADD SEQUENCIA_ERRADA INT NULL "
    Banco.Execute sql
       
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes102013()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '102013'"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_AMB ADD REFERENCIA VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_INT ADD REFERENCIA VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MOVIM_EXT ADD REFERENCIA VARCHAR(20)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD MEDICOLAUDOPADRAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE LAU_MODELO ADD NOME_ABREVIADO VARCHAR(5)"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNOPROCEDIMENTO ADD MEDICOSOL INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE EXTERNOPROCEDIMENTO ADD AUTORIZACAOPROCEDIMENTO BIGINT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD WEBSERVICEURL VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PARAMETRO ADD WEBSERVICEURI VARCHAR(500)"
   Banco.Execute sql
   
   sql = "ALTER TABLE REGIMENTO ADD PERMITEIMPRIMIR INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE REGIMENTO ADD CONSTRAINT DF_REGIMENTO_PERMITEIMPRIMIR DEFAULT 1 FOR PERMITEIMPRIMIR"
   Banco.Execute sql
   
   sql = "UPDATE REGIMENTO " & Chr(13)
   sql = sql & "SET PERMITEIMPRIMIR = 1 " & Chr(13)
   sql = sql & "WHERE PERMITEIMPRIMIR IS NULL "
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_INT ADD  CONSTRAINT FK_PREELETPROCEDIMENTOENFERMAGEM_INT_MOVIM_PRES_INT FOREIGN KEY(REGISTRO, SEQUENCIA) REFERENCES MOVIM_PRES_INT (REGISTRO, Sequencia)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_AMB DROP CONSTRAINT FK_PREELETPROCEDIMENTOENFERMAGEM_AMB_Movim_AMB"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_AMB ADD  CONSTRAINT FK_PREELETPROCEDIMENTOENFERMAGEM_AMB_MOVIM_PRES_AMB FOREIGN KEY(REGISTRO, SEQUENCIA) REFERENCES MOVIM_PRES_AMB (REGISTRO, Sequencia)"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_EXT DROP CONSTRAINT FK_PREELETPROCEDIMENTOENFERMAGEM_EXT_Movim_EXT"
   Banco.Execute sql
   
   sql = "ALTER TABLE PREELETPROCEDIMENTOENFERMAGEM_EXT ADD  CONSTRAINT FK_PREELETPROCEDIMENTOENFERMAGEM_EXT_MOVIM_PRES_EXT FOREIGN KEY(REGISTRO, SEQUENCIA) REFERENCES MOVIM_PRES_EXT (REGISTRO, Sequencia)"
   Banco.Execute sql
   
   sql = " ALTER TABLE CONVENIOS ADD TISS_RESUMOINTERNACAO_EXIBEREGISTRO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMP_CIHA ADD PACIENTEDIA INT"
   Banco.Execute sql
   
   sql = "UPDATE MENU SET PERMISSAOFORM = 1 WHERE NOMESUBNOVO = 'mnuEnf_Alp'"
   Banco.Execute sql
   
   'Menu para nova tela de contagem de Estoque
   If Layout = 48 Then
        sql = ""
        sql = sql & "INSERT INTO [dbo].[MENU] " & vbCrLf
        sql = sql & "            ([MENU], " & vbCrLf
        sql = sql & "             [NOMECAPTION], " & vbCrLf
        sql = sql & "             [NOMESUB], " & vbCrLf
        sql = sql & "             [ATUALIZACAO], " & vbCrLf
        sql = sql & "             [NOMESUBNOVO], " & vbCrLf
        sql = sql & "             [ATIVADO], " & vbCrLf
        sql = sql & "             [MODULO], " & vbCrLf
        sql = sql & "             [HIERARQUIA], " & vbCrLf
        sql = sql & "             [NOMESUBAUX], " & vbCrLf
        sql = sql & "             [NIVELVISIBILIDADE]) " & vbCrLf
        sql = sql & "SELECT Max(MENU) + 1 AS MENU, " & vbCrLf
        sql = sql & "       'Contagem de Estoque [Novo]', " & vbCrLf
        sql = sql & "       'mnuPrd_PCC_CEN', " & vbCrLf
        sql = sql & "       '', " & vbCrLf
        sql = sql & "       'mnuPrd_PCC_CEN', " & vbCrLf
        sql = sql & "       1, " & vbCrLf
        sql = sql & "       2, " & vbCrLf
        sql = sql & "       '0102200000', " & vbCrLf
        sql = sql & "       'mnuPrd_PCC_CEN', " & vbCrLf
        sql = sql & "       1             menu "
        Banco.Execute sql
   End If
   
   sql = ""
   sql = sql & "CREATE TABLE same_taxaocupacao " & vbCrLf
   sql = sql & "  ( " & vbCrLf
   sql = sql & "     unidade    INT NULL, " & vbCrLf
   sql = sql & "     descricao  VARCHAR(155) NULL, " & vbCrLf
   sql = sql & "     quantidade MONEY NULL, " & vbCrLf
   sql = sql & "     mes        VARCHAR(155) NULL, " & vbCrLf
   sql = sql & "     ip         VARCHAR(255) NULL " & vbCrLf
   sql = sql & "  ) "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "CREATE TABLE tmpcensomensal " & vbCrLf
   sql = sql & "  ( " & vbCrLf
   sql = sql & "     unidade                       VARCHAR(50) NULL, " & vbCrLf
   sql = sql & "     data                          DATETIME NULL, " & vbCrLf
   sql = sql & "     quantidadeleito               INT NULL, " & vbCrLf
   sql = sql & "     tipoconveniosusentrada        INT NULL, " & vbCrLf
   sql = sql & "     tipoconveniosussaida          INT NULL, " & vbCrLf
   sql = sql & "     tipoconveniosustotal          INT NULL, " & vbCrLf
   sql = sql & "     tipoconveniooutrosentrada     INT NULL, " & vbCrLf
   sql = sql & "     tipoconveniooutrossaida       INT NULL, " & vbCrLf
   sql = sql & "     tipoconveniooutrostotal       INT NULL, " & vbCrLf
   sql = sql & "     ip                            VARCHAR(155) NULL, " & vbCrLf
   sql = sql & "     tipoconvenioparticularentrada INT NULL, " & vbCrLf
   sql = sql & "     tipoconvenioparticularsaida   INT NULL, " & vbCrLf
   sql = sql & "     tipoconvenioparticulartotal   INT NULL, " & vbCrLf
   sql = sql & "     tipoconveniounimedtotal       INT NULL, " & vbCrLf
   sql = sql & "     tipoconveniounimedsaida       INT NULL, " & vbCrLf
   sql = sql & "     tipoconveniounimedentrada     INT NULL, " & vbCrLf
   sql = sql & "     taxaocupacao                  MONEY NULL, " & vbCrLf
   sql = sql & "     mediapermanencia              MONEY NULL " & vbCrLf
   sql = sql & "  ) "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "CREATE TABLE same_estatistica " & vbCrLf
   sql = sql & "  ( " & vbCrLf
   sql = sql & "     sequencia        INT IDENTITY(1, 1) NOT NULL, " & vbCrLf
   sql = sql & "     tipounidade      INT NULL, " & vbCrLf
   sql = sql & "     tipoconvenio     INT NULL, " & vbCrLf
   sql = sql & "     descricao        VARCHAR(50) NULL, " & vbCrLf
   sql = sql & "     mes1             MONEY NULL, " & vbCrLf
   sql = sql & "     mes2             MONEY NULL, " & vbCrLf
   sql = sql & "     mes3             MONEY NULL, " & vbCrLf
   sql = sql & "     mes4             MONEY NULL, " & vbCrLf
   sql = sql & "     mes5             MONEY NULL, " & vbCrLf
   sql = sql & "     mes6             MONEY NULL, " & vbCrLf
   sql = sql & "     mes7             MONEY NULL, " & vbCrLf
   sql = sql & "     mes8             MONEY NULL, " & vbCrLf
   sql = sql & "     mes9             MONEY NULL, " & vbCrLf
   sql = sql & "     mes10            MONEY NULL, " & vbCrLf
   sql = sql & "     mes11            MONEY NULL, " & vbCrLf
   sql = sql & "     mes12            MONEY NULL, " & vbCrLf
   sql = sql & "     ip               VARCHAR(100) NULL, " & vbCrLf
   sql = sql & "     total            MONEY NULL, " & vbCrLf
   sql = sql & "     descricaounidade VARCHAR(100) NULL, " & vbCrLf
   sql = sql & "     procedimentonome VARCHAR(250) NULL, " & vbCrLf
   sql = sql & "     cota             INT NULL, " & vbCrLf
   sql = sql & "     procedimento     INT NULL " & vbCrLf
   sql = sql & "  ) "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE same_estatistica " & vbCrLf
   sql = sql & "  ADD CONSTRAINT pk_same_estatistica PRIMARY KEY CLUSTERED (sequencia) "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "CREATE TABLE same_estatistica_ambulatorial " & vbCrLf
   sql = sql & "  ( " & vbCrLf
   sql = sql & "     procedimento     BIGINT NULL, " & vbCrLf
   sql = sql & "     procedimentonome VARCHAR(255) NULL, " & vbCrLf
   sql = sql & "     mes1             MONEY NULL, " & vbCrLf
   sql = sql & "     mes2             MONEY NULL, " & vbCrLf
   sql = sql & "     mes3             MONEY NULL, " & vbCrLf
   sql = sql & "     mes4             MONEY NULL, " & vbCrLf
   sql = sql & "     mes5             MONEY NULL, " & vbCrLf
   sql = sql & "     mes6             MONEY NULL, " & vbCrLf
   sql = sql & "     mes7             MONEY NULL, " & vbCrLf
   sql = sql & "     mes8             MONEY NULL, " & vbCrLf
   sql = sql & "     mes9             MONEY NULL, " & vbCrLf
   sql = sql & "     mes10            MONEY NULL, " & vbCrLf
   sql = sql & "     mes11            MONEY NULL, " & vbCrLf
   sql = sql & "     mes12            MONEY NULL, " & vbCrLf
   sql = sql & "     total            MONEY NULL, " & vbCrLf
   sql = sql & "     ip               VARCHAR(100) NULL " & vbCrLf
   sql = sql & "  ) "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "CREATE TABLE same_estatistica_internacao " & vbCrLf
   sql = sql & "  ( " & vbCrLf
   sql = sql & "     unidade      INT NULL, " & vbCrLf
   sql = sql & "     descricao    VARCHAR(255) NULL, " & vbCrLf
   sql = sql & "     mes1         MONEY NULL, " & vbCrLf
   sql = sql & "     pd1          MONEY NULL, " & vbCrLf
   sql = sql & "     mes2         MONEY NULL, " & vbCrLf
   sql = sql & "     pd2          MONEY NULL, " & vbCrLf
   sql = sql & "     mes3         MONEY NULL, " & vbCrLf
   sql = sql & "     pd3          MONEY NULL, " & vbCrLf
   sql = sql & "     mes4         MONEY NULL, " & vbCrLf
   sql = sql & "     pd4          MONEY NULL, " & vbCrLf
   sql = sql & "     mes5         MONEY NULL, " & vbCrLf
   sql = sql & "     pd5          MONEY NULL, " & vbCrLf
   sql = sql & "     mes6         MONEY NULL, " & vbCrLf
   sql = sql & "     pd6          MONEY NULL, " & vbCrLf
   sql = sql & "     mes7         MONEY NULL, " & vbCrLf
   sql = sql & "     pd7          MONEY NULL, " & vbCrLf
   sql = sql & "     mes8         MONEY NULL, " & vbCrLf
   sql = sql & "     pd8          MONEY NULL, " & vbCrLf
   sql = sql & "     mes9         MONEY NULL, " & vbCrLf
   sql = sql & "     pd9          MONEY NULL, " & vbCrLf
   sql = sql & "     mes10        MONEY NULL, " & vbCrLf
   sql = sql & "     pd10         MONEY NULL, " & vbCrLf
   sql = sql & "     mes11        MONEY NULL, " & vbCrLf
   sql = sql & "     pd11         MONEY NULL, " & vbCrLf
   sql = sql & "     mes12        MONEY NULL, " & vbCrLf
   sql = sql & "     pd12         MONEY NULL, " & vbCrLf
   sql = sql & "     total        MONEY NULL, " & vbCrLf
   sql = sql & "     totalpd      MONEY NULL, " & vbCrLf
   sql = sql & "     ip           VARCHAR(100) NULL, " & vbCrLf
   sql = sql & "     tipoconvenio INT NULL, " & vbCrLf
   sql = sql & "     convenio     INT NULL, " & vbCrLf
   sql = sql & "     tipo         INT NULL " & vbCrLf
   sql = sql & "  ) "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE tmpcensomensal " & vbCrLf
   sql = sql & "  ADD taxaocupacao MONEY NULL "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE leitounidade " & vbCrLf
   sql = sql & "  ADD leitoativos INT NULL "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE tmpcensomensal " & vbCrLf
   sql = sql & "  ADD mediapermanencia MONEY NULL "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE tmpinternacao " & vbCrLf
   sql = sql & "  ALTER COLUMN responsavel VARCHAR(255)"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPAUTORIZACAODEBITOBANCO ADD VENCIMENTO DATETIME"
   Banco.Execute sql
   
        sql = ""
        sql = sql & "CREATE VIEW V_RPT_CONTAGEM " & vbCrLf
        sql = sql & "AS " & vbCrLf
        sql = sql & "  SELECT " & vbCrLf
        sql = sql & "  --Centro de custo " & vbCrLf
        sql = sql & "  C.CENTROCUSTO                           AS ID_CENTROCUSTO, " & vbCrLf
        sql = sql & "  CC.DESCRICAO                            AS CENTROCUSTO, " & vbCrLf
        sql = sql & "  --Produto " & vbCrLf
        sql = sql & "  A.PRODUTO                               AS ID_PRODUTO, " & vbCrLf
        sql = sql & "  Cast(LEFT(Str(A.PRODUTO, 7), 2) AS INT) AS TIPOPRODUTO, " & vbCrLf
        sql = sql & "  B.DESCRICAO                             AS PRODUTO, " & vbCrLf
        sql = sql & "  CL.SALDO                                AS SALDO_ATUAL, " & vbCrLf
        sql = sql & "  Isnull(CL.LOTE, 'S/ LOTE')              AS LOTE, " & vbCrLf
        sql = sql & "  Isnull(CL.VALIDADE, '1900/01/01')       AS VALIDADE, " & vbCrLf
        sql = sql & "  B.UNIDADELANCAMENTO, " & vbCrLf
        sql = sql & "  B.FAMILIA, " & vbCrLf
        sql = sql & "  B.CUSTOMEDIO1                           AS CUSTO, " & vbCrLf
        sql = sql & "  --Contagem " & vbCrLf
        sql = sql & "  A.SALDO                                 AS SALDO_NOVO, " & vbCrLf
        sql = sql & "  A.CONTAGEMESTOQUE, " & vbCrLf
        sql = sql & "  A.NUMEROCONTAGEM, " & vbCrLf
        sql = sql & "  A.SITUACAO " & vbCrLf
        sql = sql & "  FROM   CONTAGEM_PRODUTOS A " & vbCrLf
        sql = sql & "         INNER JOIN PRODUTO B " & vbCrLf
        sql = sql & "                 ON B.PRODUTO = A.PRODUTO " & vbCrLf
        sql = sql & "         INNER JOIN PRODUTOCENTROCUSTO C " & vbCrLf
        sql = sql & "                 ON C.PRODUTO = A.PRODUTO " & vbCrLf
        sql = sql & "         INNER JOIN CENTROCUSTO CC " & vbCrLf
        sql = sql & "                 ON CC.CENTROCUSTO = C.CENTROCUSTO " & vbCrLf
        sql = sql & "         LEFT JOIN PRODUTOCENTROCUSTOLOTE CL " & vbCrLf
        sql = sql & "                ON CL.PRODUTO = C.PRODUTO " & vbCrLf
        sql = sql & "                   AND CL.CENTROCUSTO = C.CENTROCUSTO " & vbCrLf
        sql = sql & "                   AND Isnull(CL.LOTE, 'S/ LOTE') = Isnull(A.LOTE, 'S/ LOTE') " & vbCrLf
        sql = sql & "                   AND Isnull(CL.VALIDADE, '1900/01/01') = " & vbCrLf
        sql = sql & "                       Isnull(A.VALIDADE, '1900/01/01') "
    Banco.Execute sql
   
        sql = ""
        sql = sql & "CREATE VIEW V_RPT_FOOTER " & vbCrLf
        sql = sql & "AS " & vbCrLf
        sql = sql & "  SELECT 'M2TECNOLOGIA'                          AS EMPRESA, " & vbCrLf
        sql = sql & "         'WWW.M2TEC.COM.BR'                      AS SITE, " & vbCrLf
        sql = sql & "         'M2 TECNOLOGIA | WWW.M2TEC.COM.BR'      AS EMPRESA_SITE, " & vbCrLf
        sql = sql & "         '[SAVE - SISTEMA DE GESTÃO HOSPITALAR]' AS PRODUTO_HOSPITALAR "
    Banco.Execute sql
        
        sql = ""
        sql = sql & "CREATE VIEW V_RPT_HEADER " & vbCrLf
        sql = sql & "AS " & vbCrLf
        sql = sql & "  SELECT A.EMPRESA   AS NOME, " & vbCrLf
        sql = sql & "         A.RUA + ', ' + A.NUMERO + ', ' + A.BAIRRO + '-' + A.CIDADE " & vbCrLf
        sql = sql & "         + '/' + UF  AS ENDERECO, " & vbCrLf
        sql = sql & "         A.EMAIL     AS EMAIL, " & vbCrLf
        sql = sql & "         A.TELEFONE1 AS TELEFONE, " & vbCrLf
        sql = sql & "         A.LOGO      AS LOGO " & vbCrLf
        sql = sql & "  FROM   PARAMETRO A "
        Banco.Execute sql
   
    sql = ""
        sql = sql & "ALTER TABLE PARAMETRO " & vbCrLf
        sql = sql & "  ADD EXIBE_SALDO_LISTA_CONTAGEM BIT DEFAULT(0)"
        Banco.Execute sql
        
        If Layout = 48 Then
                sql = ""
                sql = sql & " " & vbCrLf
                sql = sql & "UPDATE PARAMETRO " & vbCrLf
                sql = sql & "SET    EXIBE_SALDO_LISTA_CONTAGEM = 1 "
                Banco.Execute sql
        End If
    
    sql = ""
    sql = sql & "ALTER TABLE CIR_PACIENTE_CIRURGIA " & vbCrLf
    sql = sql & "  ADD BOXATUALIZACAO VARCHAR(100) "
    Banco.Execute sql
   sql = "CREATE VIEW V_RPT_HEAD " & Chr(13)
   sql = sql & "AS " & Chr(13)
   sql = sql & "SELECT " & Chr(13)
   sql = sql & "A.EMPRESA AS NOME, " & Chr(13)
   sql = sql & "A.RUA +', '+ A.NUMERO +', '+ A.BAIRRO +'-'+ A.CIDADE +'/'+ UF AS ENDERECO, " & Chr(13)
   sql = sql & "A.EMAIL AS EMAIL, " & Chr(13)
   sql = sql & "A.TELEFONE1 AS TELEFONE, " & Chr(13)
   sql = sql & "A.LOGO AS LOGO " & Chr(13)
   sql = sql & "FROM PARAMETRO A"
   Banco.Execute sql
   
   sql = "CREATE VIEW V_RPT_TRAIL " & Chr(13)
   sql = sql & "AS " & Chr(13)
   sql = sql & "SELECT " & Chr(13)
   sql = sql & "'M2TECNOLOGIA' AS EMPRESA, " & Chr(13)
   sql = sql & "'[SAVE - SISTEMA DE GESTÃO HOSPITALAR]' AS PRODUTO, " & Chr(13)
   sql = sql & "'WWW.M2TEC.COM.BR' AS SITE"
   Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function
Public Function AtualizaMes112013()
On Error GoTo Erro
   
    'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
    'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
    sql = ""
    sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '112013'"
    Banco.Execute sql

   If Layout = 47 Or Layout = 44 Then
        
        sql = ""
        sql = sql & "INSERT INTO MENU " & vbCrLf
        sql = sql & "            ([MENU], " & vbCrLf
        sql = sql & "             [NOMECAPTION], " & vbCrLf
        sql = sql & "             [NOMESUB], " & vbCrLf
        sql = sql & "             [ATUALIZACAO], " & vbCrLf
        sql = sql & "             [NOMESUBNOVO], " & vbCrLf
        sql = sql & "             [ATIVADO], " & vbCrLf
        sql = sql & "             [MODULO], " & vbCrLf
        sql = sql & "             [HIERARQUIA], " & vbCrLf
        sql = sql & "             [NOMESUBAUX], " & vbCrLf
        sql = sql & "             [NIVELVISIBILIDADE]) " & vbCrLf
        sql = sql & "SELECT Max(MENU) + 1, " & vbCrLf
        sql = sql & "       'Brasindice/ABCFarma', " & vbCrLf
        sql = sql & "       'mnuPrd_Tbl_Atp_BraABC', " & vbCrLf
        sql = sql & "       '2013-04-26 SA', " & vbCrLf
        sql = sql & "       'mnuPrd_Tbl_Atp_BraABC', " & vbCrLf
        sql = sql & "       1, " & vbCrLf
        sql = sql & "       2, " & vbCrLf
        sql = sql & "       '0111020100', " & vbCrLf
        sql = sql & "       'mnuPrd_Tbl_Atp_BraABC', " & vbCrLf
        sql = sql & "       1 " & vbCrLf
        sql = sql & "FROM   MENU"
        Banco.Execute sql
    
    sql = "ALTER TABLE TMP_TISS_HONORARIOINDIVIDUAL ADD DATAFATURAMENTO DATETIME"
    Banco.Execute sql
    
    sql = "ALTER TABLE CONVENIOS ADD TISS_HONORARIOINDIVIDUALDATAFATURAMENTO INT"
    Banco.Execute sql

        sql = ""
        sql = sql & "INSERT INTO MENU " & vbCrLf
        sql = sql & "            ([MENU], " & vbCrLf
        sql = sql & "             [NOMECAPTION], " & vbCrLf
        sql = sql & "             [NOMESUB], " & vbCrLf
        sql = sql & "             [ATUALIZACAO], " & vbCrLf
        sql = sql & "             [NOMESUBNOVO], " & vbCrLf
        sql = sql & "             [ATIVADO], " & vbCrLf
        sql = sql & "             [MODULO], " & vbCrLf
        sql = sql & "             [HIERARQUIA], " & vbCrLf
        sql = sql & "             [NOMESUBAUX], " & vbCrLf
        sql = sql & "             [NIVELVISIBILIDADE]) " & vbCrLf
        sql = sql & "SELECT Max(MENU) + 1, " & vbCrLf
        sql = sql & "       'SIMPRO', " & vbCrLf
        sql = sql & "       'mnuPrd_Tbl_Atp_Spo', " & vbCrLf
        sql = sql & "       '2013-04-26 SA', " & vbCrLf
        sql = sql & "       'mnuPrd_Tbl_Atp_Spo', " & vbCrLf
        sql = sql & "       1, " & vbCrLf
        sql = sql & "       2, " & vbCrLf
        sql = sql & "       '0111020200', " & vbCrLf
        sql = sql & "       'mnuPrd_Tbl_Atp_Spo', " & vbCrLf
        sql = sql & "       1 " & vbCrLf
        sql = sql & "FROM   MENU "
        Banco.Execute sql
   End If

        sql = ""
        sql = sql & "ALTER TABLE CIRURGIACADASTRO " & vbCrLf
        sql = sql & "  ADD PESQUISA_PADRAO INT "
    Banco.Execute sql
    
   sql = ""
   sql = sql & " INSERT INTO MENU("
   sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
   sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE)"
   sql = sql & " SELECT MAX(MENU)+1,'Plantão', 'MnuPar_Cad_Pla', ' ', 'MnuPar_Cad_Pla',"
   sql = sql & "        0, 1, '0101100000', 'MnuPar_Cad_Pla', 1"
   sql = sql & " FROM MENU "
   Banco.Execute sql

        sql = ""
        sql = sql & "ALTER TABLE MOVIM_PRES_INT " & vbCrLf
        sql = sql & "  ADD HORA_CONFERENCIA CHAR(5) NULL"
        Banco.Execute sql

        sql = ""
        sql = sql & "ALTER TABLE MOVIM_PRES_AMB " & vbCrLf
        sql = sql & "  ADD HORA_CONFERENCIA CHAR(5) NULL"
        Banco.Execute sql

        sql = ""
        sql = sql & "ALTER TABLE MOVIM_PRES_EXT " & vbCrLf
        sql = sql & "  ADD HORA_CONFERENCIA CHAR(5) NULL "
        Banco.Execute sql
   
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes122013()
On Error GoTo Erro
   
    'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
    'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '122013'"
   Banco.Execute sql
        
   sql = ""
   sql = sql & "ALTER TABLE MOVIM_PRES_INT " & vbCrLf
   sql = sql & "  ADD HORA_CONFERENCIA CHAR(5) NULL"
   Banco.Execute sql

   sql = ""
   sql = sql & "ALTER TABLE MOVIM_PRES_AMB " & vbCrLf
   sql = sql & "  ADD HORA_CONFERENCIA CHAR(5) NULL"
   Banco.Execute sql

   sql = ""
   sql = sql & "ALTER TABLE MOVIM_PRES_EXT " & vbCrLf
   sql = sql & "  ADD HORA_CONFERENCIA CHAR(5) NULL "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE CONT_PLANOCONTA" & vbCrLf
   sql = sql & "ADD ANALITICA BIT NULL" & vbCrLf
   Banco.Execute sql
        
   sql = "ALTER TABLE FARMACIAFAMILIA " & Chr(13)
   sql = sql & "ADD ANTIMICROBIANO INT"
   Banco.Execute sql
   
        sql = ""
        sql = sql & " SELECT ISNULL(OBJECT_ID('EMAIL_SERVIDOR'),0) AS EXISTE "
        Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
        
        If tbl!Existe = 0 Then
            sql = ""
            sql = sql & "CREATE TABLE EMAIL_SERVIDOR " & vbCrLf
            sql = sql & "  ( " & vbCrLf
            sql = sql & "     ID            INT IDENTITY(1, 1) NOT NULL, " & vbCrLf
            sql = sql & "     NOME          VARCHAR(100) NOT NULL, " & vbCrLf
            sql = sql & "     ENDERECO_SMTP VARCHAR(100), " & vbCrLf
            sql = sql & "     PORTA_SMTP    CHAR(5), " & vbCrLf
            sql = sql & "     ENDERECO_POP  VARCHAR(100), " & vbCrLf
            sql = sql & "     PORTA_POP     CHAR(5), " & vbCrLf
            sql = sql & "     AUTENTICACAO  BIT, " & vbCrLf
            sql = sql & "     TSL           BIT, " & vbCrLf
            sql = sql & "     SSL           BIT, " & vbCrLf
            sql = sql & "     TIMEOUT       INT " & vbCrLf
            sql = sql & "  )"
            Banco.Execute sql
            
            sql = ""
            sql = sql & " " & vbCrLf
            sql = sql & "ALTER TABLE EMAIL_SERVIDOR " & vbCrLf
            sql = sql & "  ADD CONSTRAINT PK_EMAIL_SERVIDOR PRIMARY KEY (ID)"
            Banco.Execute sql
            
            sql = ""
            sql = sql & " " & vbCrLf
            sql = sql & "ALTER TABLE EMAIL_SERVIDOR " & vbCrLf
            sql = sql & "  ADD CONSTRAINT UN_EMAIL_SERVIDOR_DOMINIO UNIQUE (NOME)"
            Banco.Execute sql
        End If
        
        sql = ""
        sql = sql & " SELECT ISNULL(OBJECT_ID('EMAIL_USUARIO'),0) AS EXISTE "
        Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
        
        If tbl!Existe = 0 Then
            sql = ""
            sql = sql & " " & vbCrLf
            sql = sql & "CREATE TABLE EMAIL_USUARIO " & vbCrLf
            sql = sql & "  ( " & vbCrLf
            sql = sql & "     USUARIO INT NOT NULL, " & vbCrLf
            sql = sql & "     LOGIN   VARCHAR(100), " & vbCrLf
            sql = sql & "     SENHA   VARCHAR(100) " & vbCrLf
            sql = sql & "  )"
            Banco.Execute sql
            
            sql = ""
            sql = sql & " " & vbCrLf
            sql = sql & "ALTER TABLE EMAIL_USUARIO " & vbCrLf
            sql = sql & "  ADD CONSTRAINT PK_EMAIL_USUARIO_USUARIO PRIMARY KEY (USUARIO)"
            Banco.Execute sql
            
            sql = ""
            sql = sql & " " & vbCrLf
            sql = sql & "ALTER TABLE EMAIL_USUARIO " & vbCrLf
            sql = sql & "  ADD CONSTRAINT FK_EMAIL_USUARIO_USUARIO FOREIGN KEY (USUARIO) REFERENCES " & vbCrLf
            sql = sql & "  USUARIO(USUARIO)"
            Banco.Execute sql
            
            sql = ""
            sql = sql & " " & vbCrLf
            sql = sql & "ALTER TABLE EMAIL_USUARIO " & vbCrLf
            sql = sql & "  ADD CONSTRAINT UN_EMAIL_USUARIO_LOGIN UNIQUE (LOGIN) "
            Banco.Execute sql
          End If
        
        sql = ""
        sql = sql & "INSERT INTO MENU " & vbCrLf
        sql = sql & "            (MENU, " & vbCrLf
        sql = sql & "             NOMECAPTION, " & vbCrLf
        sql = sql & "             NOMESUB, " & vbCrLf
        sql = sql & "             ATUALIZACAO, " & vbCrLf
        sql = sql & "             NOMESUBNOVO, " & vbCrLf
        sql = sql & "             ATIVADO, " & vbCrLf
        sql = sql & "             MODULO, " & vbCrLf
        sql = sql & "             HIERARQUIA, " & vbCrLf
        sql = sql & "             NOMESUBAUX, " & vbCrLf
        sql = sql & "             NIVELVISIBILIDADE) " & vbCrLf
        sql = sql & "SELECT MAX(MENU) + 1, " & vbCrLf
        sql = sql & "       'Parâmetros Globais', " & vbCrLf
        sql = sql & "       'mnuPar_Par_Global', " & vbCrLf
        sql = sql & "       '', " & vbCrLf
        sql = sql & "       'mnuPar_Par_Global', " & vbCrLf
        sql = sql & "       1, " & vbCrLf
        sql = sql & "       1, " & vbCrLf
        sql = sql & "       '0102010000', " & vbCrLf
        sql = sql & "       'mnuPar_Par_Global', " & vbCrLf
        sql = sql & "       1 " & vbCrLf
        sql = sql & "FROM   MENU"
        Banco.Execute sql
        
        sql = ""
        sql = sql & " " & vbCrLf
        sql = sql & "INSERT INTO MENU " & vbCrLf
        sql = sql & "            (MENU, " & vbCrLf
        sql = sql & "             NOMECAPTION, " & vbCrLf
        sql = sql & "             NOMESUB, " & vbCrLf
        sql = sql & "             ATUALIZACAO, " & vbCrLf
        sql = sql & "             NOMESUBNOVO, " & vbCrLf
        sql = sql & "             ATIVADO, " & vbCrLf
        sql = sql & "             MODULO, " & vbCrLf
        sql = sql & "             HIERARQUIA, " & vbCrLf
        sql = sql & "             NOMESUBAUX, " & vbCrLf
        sql = sql & "             NIVELVISIBILIDADE) " & vbCrLf
        sql = sql & "SELECT MAX(MENU) + 1, " & vbCrLf
        sql = sql & "       'Servidor de Email', " & vbCrLf
        sql = sql & "       'mnuPar_Par_Email', " & vbCrLf
        sql = sql & "       '', " & vbCrLf
        sql = sql & "       'mnuPar_Par_Email', " & vbCrLf
        sql = sql & "       1, " & vbCrLf
        sql = sql & "       1, " & vbCrLf
        sql = sql & "       '0102020000', " & vbCrLf
        sql = sql & "       'mnuPar_Par_Email', " & vbCrLf
        sql = sql & "       1 " & vbCrLf
        sql = sql & "FROM   MENU"
        Banco.Execute sql

        sql = ""
        sql = sql & " SELECT ISNULL(OBJECT_ID('FORNECEDOR_FAMILIA'),0) AS EXISTE "
        Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
        
        If tbl!Existe = 0 Then
                 sql = ""
                 sql = sql & "CREATE TABLE FORNECEDOR_FAMILIA " & vbCrLf
                 sql = sql & "  ( " & vbCrLf
                 sql = sql & "     FORNECEDOR  INT NOT NULL, " & vbCrLf
                 sql = sql & "     CENTROCUSTO SMALLINT NOT NULL, " & vbCrLf
                 sql = sql & "     FAMILIA     INT NOT NULL " & vbCrLf
                 sql = sql & "  )"
                 Banco.Execute sql

                 sql = ""
                 sql = sql & " " & vbCrLf
                 sql = sql & "ALTER TABLE FORNECEDOR_FAMILIA " & vbCrLf
                 sql = sql & "  ADD CONSTRAINT FK_FORNECEDOR_FAMILIAS_FORNECEDOR FOREIGN KEY (FORNECEDOR) " & vbCrLf
                 sql = sql & "  REFERENCES FORNECEDORES(FORNECEDOR)"
                 Banco.Execute sql

                 sql = ""
                 sql = sql & " " & vbCrLf
                 sql = sql & "ALTER TABLE FORNECEDOR_FAMILIA " & vbCrLf
                 sql = sql & "  ADD CONSTRAINT FK_FORNECEDOR_FAMILIA_CENTROCUSTO FOREIGN KEY (CENTROCUSTO) " & vbCrLf
                 sql = sql & "  REFERENCES CENTROCUSTO(CENTROCUSTO) "
                 Banco.Execute sql
         End If
        
        'Colocado para executar primeiro a de atualiza saldo porque a trigger de calculo de preco medio considera que esta ja foi executada
        sql = ""
        sql = "sp_settriggerorder 'TR_ATUALIZASALDOPRODUTODIVERSOS', 'First', 'insert'"
        Banco.Execute sql
        
        sql = ""
        sql = "sp_settriggerorder 'TR_ATUALIZASALDOPRODUTODIVERSOS', 'First', 'delete'"
        Banco.Execute sql
        
        sql = ""
        sql = " ALTER TABLE SAVELOG..PSICOTROPICOMOVIMENTOLOG " & vbCrLf
        sql = sql & " ADD CENTROCUSTO INT NULL" & vbCrLf
        Banco.Execute sql
        
        sql = ""
        sql = "ALTER TABLE INTERNO " & vbCrLf
        sql = sql & "  ADD USUARIO_FECHAMENTO INT NULL"
        Banco.Execute sql

        sql = ""
        sql = "ALTER TABLE EXTERNO " & vbCrLf
        sql = sql & "  ADD USUARIO_FECHAMENTO INT NULL"
        Banco.Execute sql
        
        sql = ""
        sql = "ALTER TABLE AMBULATORIAL " & vbCrLf
        sql = sql & "  ADD USUARIO_FECHAMENTO INT NULL"
        Banco.Execute sql
        
        sql = "ALTER TABLE CONVENIOS ADD AVISARPROCEDIMENTOSZERADOS INT"
        Banco.Execute sql
        
        sql = "ALTER TABLE CONVENIOS ADD NAOPERMITIRLANCARPROCEDIMENTOZERADO INT"
        Banco.Execute sql
        
        sql = ""
        sql = "ALTER TABLE LAU_MOVIM_AMB " & vbCrLf
        sql = sql & "  ADD IMPRESSOGUIA INT NULL"
        Banco.Execute sql
        
        sql = ""
        sql = "ALTER TABLE LAU_MOVIM_EXT " & vbCrLf
        sql = sql & "  ADD IMPRESSOGUIA INT NULL"
        Banco.Execute sql
        
        sql = ""
        sql = "ALTER TABLE LAU_MOVIM_INT " & vbCrLf
        sql = sql & "  ADD IMPRESSOGUIA INT NULL"
        Banco.Execute sql
   Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes012014()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '012014'"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIODATA_PERIODO ADD TIPOINTEGRACAO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE CONVENIODATA_PERIODO ADD INTEGRACAOTERCEIRO INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPINTEGRAFATURAMENTO ADD VALORFATURADO MONEY"
   Banco.Execute sql
   
   sql = "ALTER TABLE TMPINTEGRAFATURAMENTO ADD SEQUENCIARELACIONADA INT"
   Banco.Execute sql
   
   sql = "ALTER TABLE TER_LANCAMENTO ADD CONVENIO INT"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[TER_INTEGRACAOTERCEIRO]( " & Chr(13)
   sql = sql & "[CONVENIO] [int] NOT NULL, " & Chr(13)
   sql = sql & "[FATURA] [varchar](10) NOT NULL, " & Chr(13)
   sql = sql & "[TERCEIRO] [int] NOT NULL, " & Chr(13)
   sql = sql & "[VALORPAGO] [money] NULL, " & Chr(13)
   sql = sql & "[VALORFATURADO] [money] NULL, " & Chr(13)
   sql = sql & "[OBSERVACAO] [varchar](255) NULL, " & Chr(13)
   sql = sql & "[ATUALIZACAO] [nchar](155) NULL, " & Chr(13)
   sql = sql & "[SEQUENCIARELACIONADA] [int] NULL, " & Chr(13)
   sql = sql & "[SEQUENCIA] [int] NULL, " & Chr(13)
   sql = sql & "[TIPOREGISTRO] [int] NULL, " & Chr(13)
   sql = sql & "[NOME] [varchar](155) NULL " & Chr(13)
   sql = sql & ") ON [PRIMARY] " & Chr(13)
   Banco.Execute sql
   
   sql = "ALTER TABLE MOVIM_PRES_AMB ADD INSEREAVISOREPETIDO INT"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[tmpRECIBOPAGAMENTO]( " & Chr(13)
   sql = sql & "[PAGADOR] [varchar](255) NULL, " & Chr(13)
   sql = sql & "[PACIENTE] [varchar](255) NULL, " & Chr(13)
   sql = sql & "[VALOR] [varchar](255) NULL, " & Chr(13)
   sql = sql & "[VALOREXTENSO] [varchar](255) NULL, " & Chr(13)
   sql = sql & "[FORMAPAGAMENTO] [varchar](255) NULL, " & Chr(13)
   sql = sql & "[REFERENTE] [varchar](255) NULL, " & Chr(13)
   sql = sql & "[DESCRICAO] [varchar](255) NULL " & Chr(13)
   sql = sql & ") ON [PRIMARY] " & Chr(13)
   Banco.Execute sql
   
   Dim tbl As rdoResultset
   sql = "SELECT TABLE_NAME,COLUMN_NAME" & Chr(13)
   sql = sql & "FROM INFORMATION_SCHEMA.Columns WITH(NOLOCK)" & Chr(13)
   sql = sql & "WHERE TABLE_NAME IN ('COTACAO1','COTACAOITEM1')" & Chr(13)
   sql = sql & "AND DATA_TYPE = 'varchar'" & Chr(13)
   sql = sql & "AND CHARACTER_MAXIMUM_LENGTH <> 1024" & Chr(13)
   sql = sql & "AND COLUMN_NAME IN ('OBSERVACAO')" & Chr(13)
   Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
   While Not tbl.EOF
         sql = "ALTER TABLE " & tbl!TABLE_NAME & " ALTER COLUMN " & tbl!COLUMN_NAME & " varchar(1024)"
         Banco.Execute sql
         tbl.MoveNext
   Wend
      
   sql = "ALTER TABLE PARAMETRO ADD CAMINHO_BACKUP_SERVIDOR VARCHAR(250)"
   Banco.Execute sql

   sql = "SELECT TABLE_NAME,COLUMN_NAME" & Chr(13)
   sql = sql & "FROM INFORMATION_SCHEMA.Columns WITH(NOLOCK)" & Chr(13)
   sql = sql & "WHERE TABLE_NAME IN ('INTERNO','EXTERNO','AMBULATORIAL')" & Chr(13)
   sql = sql & "AND DATA_TYPE = 'varchar'" & Chr(13)
   sql = sql & "AND CHARACTER_MAXIMUM_LENGTH <> 1023" & Chr(13)
   sql = sql & "AND COLUMN_NAME IN ('OBSERVACAO','OBS','OBSERVACAO_ADMINISTRATIVA','OBSERVACAO_PROFISSIONAL','OBSERVACAO_ENFERMAGEM','OBSERVACAO_CONVENIO')" & Chr(13)
   Set tbl = Banco.OpenResultset(sql, rdOpenStatic)
   While Not tbl.EOF
         sql = "ALTER TABLE " & tbl!TABLE_NAME & " ALTER COLUMN " & tbl!COLUMN_NAME & " varchar(1023)"
         Banco.Execute sql
         tbl.MoveNext
   Wend
   
   sql = ""
   sql = sql & "ALTER TABLE COTACAO1"
   sql = sql & " ADD TITULO_COTACAO VARCHAR(300),"
   sql = sql & " FORMA_PAGAMENTO INT"
   Banco.Execute sql
   
   sql = ""
   sql = sql & "CREATE TABLE [dbo].[FORMAS_PAGAMENTO](" & Chr(13)
   sql = sql & "[id] [int] NOT NULL," & Chr(13)
   sql = sql & "[descricao] [varchar](200) NOT NULL," & Chr(13)
   sql = sql & "CONSTRAINT [PK_FORMAS_PAGAMENTO] PRIMARY KEY CLUSTERED" & Chr(13)
   sql = sql & "(" & Chr(13)
   sql = sql & "[id] Asc" & Chr(13)
   sql = sql & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]" & Chr(13)
   sql = sql & ") ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & "CREATE TABLE [dbo].[COTACAOITEM_ENTREGAPROG](" & Chr(13)
   sql = sql & "[id] [int] IDENTITY(1,1) NOT NULL," & Chr(13)
   sql = sql & "[Produto] [int] NOT NULL," & Chr(13)
   sql = sql & "[Cotacao] [int] NOT NULL," & Chr(13)
   sql = sql & "[Quantidade] [decimal](18, 2) NOT NULL," & Chr(13)
   sql = sql & "[Data] [datetime] NOT NULL," & Chr(13)
   sql = sql & "CONSTRAINT [PK_COTACAOITEM_ENTREGAPROGID] PRIMARY KEY NONCLUSTERED" & Chr(13)
   sql = sql & "(" & Chr(13)
   sql = sql & "[id] Asc" & Chr(13)
   sql = sql & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]" & Chr(13)
   sql = sql & ") ON [PRIMARY]"
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE dbo.CID ADD " & Chr(13)
   sql = sql & "SITUACAO bit NOT NULL CONSTRAINT DF_CID_SITUACAO DEFAULT 1" & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE SAVELOG..MOVIM_INT" & Chr(13)
   sql = sql & "ADD CODUSUARIO VARCHAR(50)" & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE SAVELOG..MOVIM_INT" & Chr(13)
   sql = sql & "ADD CODUSUARIO VARCHAR(50)" & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE SAVELOG..MOVIM_EXT" & Chr(13)
   sql = sql & "ADD CODUSUARIO VARCHAR(50)" & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE SAVELOG..MOVIM_AMB" & Chr(13)
   sql = sql & "ADD CODUSUARIO VARCHAR(50)" & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE SAVELOG..LOG_MOVIMENTACAO" & Chr(13)
   sql = sql & "ADD CODUSUARIO VARCHAR(50)" & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE LOG_MOVIMENTACAO" & Chr(13)
   sql = sql & "ADD CODUSUARIO VARCHAR(50)" & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & " ALTER TRIGGER [dbo].[TR_LOG_MOVIM_INT] " & Chr(13)
   sql = sql & " ON [dbo].[Movim_INT] " & Chr(13)
   sql = sql & " FOR INSERT,UPDATE,DELETE " & Chr(13)
   sql = sql & " AS " & Chr(13)
   sql = sql & "  " & Chr(13)
   sql = sql & "  DECLARE @TIPOOPERACAO AS INT " & Chr(13)
   sql = sql & "  DECLARE @SEMLOG AS INT " & Chr(13)
   sql = sql & "  DECLARE @GRAVAR AS INT " & Chr(13)
   sql = sql & "  DECLARE @REGISTRO AS INT " & Chr(13)
   sql = sql & "  " & Chr(13)
   sql = sql & "  -- 0 - INCLUSÃO " & Chr(13)
   sql = sql & "  -- 1 - ALTERAÇÃO " & Chr(13)
   sql = sql & "  -- 2 - EXCLUSÃO " & Chr(13)
   sql = sql & "   " & Chr(13)
   sql = sql & "  " & Chr(13)
   sql = sql & "  SET @TIPOOPERACAO = 2 " & Chr(13)
   sql = sql & "  SET @GRAVAR=1 " & Chr(13)
   sql = sql & "   " & Chr(13)
   sql = sql & "  /* " & Chr(13)
   sql = sql & "     SE NÃO  EXISTIR NA DELETED ENTÃO É INCLUSÃO " & Chr(13)
   sql = sql & "     SE EXISTIR NA DELETED E NÃO EXISTIR NA INSERTED ENTÃO É EXCLUSÃO " & Chr(13)
   sql = sql & "     SE EXISTIR NA DELETED E EXISTIR NA INSERTED ENTÃO É ALTERAÇÃO, PORÉM NEM SEMPRE É ATUALIZADO O LOG " & Chr(13)
   sql = sql & "     POIS QUANDO É EXCLUSÃO É ALTERADO SÓ O CAMPO ATUALIZACAO PARA QUANDO FOR EXCLUIR PEGAR A COLUNA ATAULIAZACAO " & Chr(13)
   sql = sql & "     DA TABELA POR ISTO TEM O CAMPO SEMLOG, ESTE PROCESSO DE EXCLUSÃO SEMPRE QUE FOR EXCLUIR ANTES É COLOCADO " & Chr(13)
   sql = sql & "     UPDATE NA ATUALIZACAO " & Chr(13)
   sql = sql & "  */ " & Chr(13)
   sql = sql & "  IF EXISTS(SELECT REGISTRO FROM DELETED)  " & Chr(13)
   sql = sql & "  BEGIN " & Chr(13)
   sql = sql & "     SELECT @REGISTRO=REGISTRO,@SEMLOG=SEMLOG FROM INSERTED " & Chr(13)
   sql = sql & "     IF @REGISTRO IS NOT NULL " & Chr(13)
   sql = sql & "     BEGIN " & Chr(13)
   sql = sql & "        IF @SEMLOG=1  " & Chr(13)
   sql = sql & "           SET @GRAVAR=0 " & Chr(13)
   sql = sql & "        SET @TIPOOPERACAO=1 " & Chr(13)
   sql = sql & "     END " & Chr(13)
   sql = sql & "     ELSE " & Chr(13)
   sql = sql & "        SET @TIPOOPERACAO=2 " & Chr(13)
   sql = sql & "  END  " & Chr(13)
   sql = sql & "  ELSE " & Chr(13)
   sql = sql & "     SET @TIPOOPERACAO=0 " & Chr(13)
   sql = sql & "  " & Chr(13)
   sql = sql & "  IF @GRAVAR=1  " & Chr(13)
   sql = sql & "  BEGIN " & Chr(13)
   sql = sql & "  IF @TIPOOPERACAO = 2    --EXCLUSÃO " & Chr(13)
   sql = sql & "  BEGIN " & Chr(13)
   sql = sql & "     INSERT INTO SAVELOG..MOVIM_INT( " & Chr(13)
   sql = sql & "        TIPOOPERACAO,REGISTRO,DATA,USUARIO,TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE, " & Chr(13)
   sql = sql & "        VALOR,CENTROCUSTO,DATACONSUMO,SEMBAIXA, LOTE, VALIDADELOTE, DATAOPERACAO, " & Chr(13)
   sql = sql & "        FATORESTOQUE, CONVLANC, QUANTIDADEDISPENSADA, DATADIGITACAO,NOTA,TIPOPAGAMENTO,LOCALCONSUMO, " & Chr(13)
   sql = sql & "        EXCLUIDOAUDITORIA,CODUSUARIO) " & Chr(13)
   sql = sql & "     SELECT   @TIPOOPERACAO,REGISTRO,(CASE WHEN ISDATE(LEFT(ATUALIZACAO,19))=1 THEN LEFT(ATUALIZACAO,19) ELSE GETDATE() END),RTRIM(LTRIM(SUBSTRING(ATUALIZACAO,26,255))), " & Chr(13)
   sql = sql & "        TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALORUNITARIO, " & Chr(13)
   sql = sql & "        CENTROCUSTO,DATACONSUMO,SEMBAIXA, LOTE, VALIDADELOTE, GETDATE(), " & Chr(13)
   sql = sql & "        FATORESTOQUE, CONVLANC, QUANTIDADEDISPENSADA, DATADIGITACAO,NOTA,TIPOPAGAMENTO,LOCALCONSUMO, " & Chr(13)
   sql = sql & "        EXCLUIDOAUDITORIA,USUARIO " & Chr(13)
   sql = sql & "     FROM  DELETED " & Chr(13)
   sql = sql & "  END " & Chr(13)
   sql = sql & "  ELSE --ALTERAÇÃO E EXCLUSÃO " & Chr(13)
   sql = sql & "     INSERT INTO SAVELOG..MOVIM_INT( " & Chr(13)
   sql = sql & "        TIPOOPERACAO,REGISTRO,DATA,USUARIO,TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALOR, " & Chr(13)
   sql = sql & "        CENTROCUSTO,DATACONSUMO,SEMBAIXA,LOTE, VALIDADELOTE, DATAOPERACAO, " & Chr(13)
   sql = sql & "        FATORESTOQUE, CONVLANC, QUANTIDADEDISPENSADA, DATADIGITACAO,NOTA,TIPOPAGAMENTO,LOCALCONSUMO, " & Chr(13)
   sql = sql & "        EXCLUIDOAUDITORIA,CODUSUARIO) " & Chr(13)
   sql = sql & "     SELECT   @TIPOOPERACAO,REGISTRO,(CASE WHEN ISDATE(LEFT(ATUALIZACAO,19))=1 THEN LEFT(ATUALIZACAO,19) ELSE GETDATE() END),RTRIM(LTRIM(SUBSTRING(ATUALIZACAO,26,255))), " & Chr(13)
   sql = sql & "        TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALORUNITARIO, " & Chr(13)
   sql = sql & "        CENTROCUSTO,DATACONSUMO,SEMBAIXA, LOTE, VALIDADELOTE, GETDATE(), " & Chr(13)
   sql = sql & "        FATORESTOQUE, CONVLANC, QUANTIDADEDISPENSADA, DATADIGITACAO,NOTA,TIPOPAGAMENTO,LOCALCONSUMO, " & Chr(13)
   sql = sql & "        EXCLUIDOAUDITORIA,CODUSUARIO " & Chr(13)
   sql = sql & "     FROM  INSERTED " & Chr(13)
   sql = sql & "  END " & Chr(13)
   Banco.Execute sql

   sql = ""
   sql = sql & " ALTER TRIGGER [dbo].[TR_LOG_MOVIM_AMB] " & Chr(13)
   sql = sql & " ON [dbo].[Movim_AMB] " & Chr(13)
   sql = sql & " FOR INSERT,UPDATE,DELETE " & Chr(13)
   sql = sql & " AS " & Chr(13)
   sql = sql & " " & Chr(13)
   sql = sql & "  DECLARE @TIPOOPERACAO AS INT " & Chr(13)
   sql = sql & "  DECLARE @SEMLOG AS INT " & Chr(13)
   sql = sql & "  DECLARE @GRAVAR AS INT " & Chr(13)
   sql = sql & "  DECLARE @REGISTRO AS INT " & Chr(13)
   sql = sql & "  " & Chr(13)
   sql = sql & "  -- 0 - INCLUSÃO " & Chr(13)
   sql = sql & "  -- 1 - ALTERAÇÃO " & Chr(13)
   sql = sql & "  -- 2 - EXCLUSÃO " & Chr(13)
   sql = sql & "   " & Chr(13)
   sql = sql & "  " & Chr(13)
   sql = sql & "  SET @TIPOOPERACAO = 2 " & Chr(13)
   sql = sql & "  SET @GRAVAR=1 " & Chr(13)
   sql = sql & "   " & Chr(13)
   sql = sql & "  /* " & Chr(13)
   sql = sql & "     SE NÃO  EXISTIR NA DELETED ENTÃO É INCLUSÃO " & Chr(13)
   sql = sql & "     SE EXISTIR NA DELETED E NÃO EXISTIR NA INSERTED ENTÃO É EXCLUSÃO " & Chr(13)
   sql = sql & "     SE EXISTIR NA DELETED E EXISTIR NA INSERTED ENTÃO É ALTERAÇÃO, PORÉM NEM SEMPRE É ATUALIZADO O LOG " & Chr(13)
   sql = sql & "     POIS QUANDO É EXCLUSÃO É ALTERADO SÓ O CAMPO ATUALIZACAO PARA QUANDO FOR EXCLUIR PEGAR A COLUNA ATAULIAZACAO " & Chr(13)
   sql = sql & "     DA TABELA POR ISTO TEM O CAMPO SEMLOG, ESTE PROCESSO DE EXCLUSÃO SEMPRE QUE FOR EXCLUIR ANTES É COLOCADO " & Chr(13)
   sql = sql & "     UPDATE NA ATUALIZACAO " & Chr(13)
   sql = sql & "  */ " & Chr(13)
   sql = sql & "  IF EXISTS(SELECT REGISTRO FROM DELETED)  " & Chr(13)
   sql = sql & "  BEGIN " & Chr(13)
   sql = sql & "     SELECT @REGISTRO=REGISTRO,@SEMLOG=SEMLOG FROM INSERTED " & Chr(13)
   sql = sql & "     IF @REGISTRO IS NOT NULL " & Chr(13)
   sql = sql & "     BEGIN " & Chr(13)
   sql = sql & "        IF @SEMLOG=1  " & Chr(13)
   sql = sql & "           SET @GRAVAR=0 " & Chr(13)
   sql = sql & "        SET @TIPOOPERACAO=1 " & Chr(13)
   sql = sql & "     END " & Chr(13)
   sql = sql & "     ELSE " & Chr(13)
   sql = sql & "        SET @TIPOOPERACAO=2 " & Chr(13)
   sql = sql & "  END  " & Chr(13)
   sql = sql & "  ELSE " & Chr(13)
   sql = sql & "     SET @TIPOOPERACAO=0 " & Chr(13)
   sql = sql & "  " & Chr(13)
   sql = sql & "  IF @GRAVAR=1  " & Chr(13)
   sql = sql & "  BEGIN " & Chr(13)
   sql = sql & "  IF @TIPOOPERACAO = 2    --EXCLUSÃO " & Chr(13)
   sql = sql & "  BEGIN " & Chr(13)
   sql = sql & "     INSERT INTO SAVELOG..MOVIM_AMB( " & Chr(13)
   sql = sql & "        TIPOOPERACAO,REGISTRO,DATA,USUARIO,TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALOR, " & Chr(13)
   sql = sql & "        CENTROCUSTO,DATACONSUMO,SEMBAIXA, LOTE, VALIDADELOTE,NOTA,TIPOPAGAMENTO,EXCLUIDOAUDITORIA,CODUSUARIO) " & Chr(13)
   sql = sql & "     SELECT   @TIPOOPERACAO,REGISTRO,LEFT(ATUALIZACAO,19),RTRIM(LTRIM(SUBSTRING(ATUALIZACAO,26,255))), " & Chr(13)
   sql = sql & "        TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALORUNITARIO, " & Chr(13)
   sql = sql & "        CENTROCUSTO,DATACONSUMO,SEMBAIXA, LOTE, VALIDADELOTE,NOTA,TIPOPAGAMENTO,EXCLUIDOAUDITORIA,USUARIO " & Chr(13)
   sql = sql & "     FROM  DELETED " & Chr(13)
   sql = sql & "  END " & Chr(13)
   sql = sql & "  ELSE --ALTERAÇÃO E EXCLUSÃO " & Chr(13)
   sql = sql & "     INSERT INTO SAVELOG..MOVIM_AMB( " & Chr(13)
   sql = sql & "        TIPOOPERACAO,REGISTRO,DATA,USUARIO,TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALOR, " & Chr(13)
   sql = sql & "        CENTROCUSTO,DATACONSUMO,SEMBAIXA,LOTE, VALIDADELOTE,NOTA,TIPOPAGAMENTO,EXCLUIDOAUDITORIA,CODUSUARIO) " & Chr(13)
   sql = sql & "     SELECT   @TIPOOPERACAO,REGISTRO,LEFT(ATUALIZACAO,19),RTRIM(LTRIM(SUBSTRING(ATUALIZACAO,26,255))), " & Chr(13)
   sql = sql & "        TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALORUNITARIO, " & Chr(13)
   sql = sql & "        CENTROCUSTO,DATACONSUMO,SEMBAIXA,LOTE, VALIDADELOTE,NOTA,TIPOPAGAMENTO,EXCLUIDOAUDITORIA,USUARIO " & Chr(13)
   sql = sql & "     FROM  INSERTED " & Chr(13)
   sql = sql & "  END " & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & " ALTER TRIGGER [dbo].[TR_LOG_MOVIM_EXT]" & Chr(13)
   sql = sql & " ON [dbo].[Movim_EXT]" & Chr(13)
   sql = sql & " FOR INSERT,UPDATE,DELETE" & Chr(13)
   sql = sql & " AS" & Chr(13)
   sql = sql & "  DECLARE @TIPOOPERACAO AS INT" & Chr(13)
   sql = sql & "  DECLARE @SEMLOG AS INT" & Chr(13)
   sql = sql & "  DECLARE @GRAVAR AS INT" & Chr(13)
   sql = sql & "  DECLARE @REGISTRO AS INT" & Chr(13)
   sql = sql & "  -- 0 - INCLUSÃO" & Chr(13)
   sql = sql & "  -- 1 - ALTERAÇÃO" & Chr(13)
   sql = sql & "  -- 2 - EXCLUSÃO" & Chr(13)
   sql = sql & "  SET @TIPOOPERACAO = 2" & Chr(13)
   sql = sql & "  SET @GRAVAR=1" & Chr(13)
   sql = sql & "  /*" & Chr(13)
   sql = sql & "     SE NÃO  EXISTIR NA DELETED ENTÃO É INCLUSÃO" & Chr(13)
   sql = sql & "     SE EXISTIR NA DELETED E NÃO EXISTIR NA INSERTED ENTÃO É EXCLUSÃO" & Chr(13)
   sql = sql & "     SE EXISTIR NA DELETED E EXISTIR NA INSERTED ENTÃO É ALTERAÇÃO, PORÉM NEM SEMPRE É ATUALIZADO O LOG" & Chr(13)
   sql = sql & "     POIS QUANDO É EXCLUSÃO É ALTERADO SÓ O CAMPO ATUALIZACAO PARA QUANDO FOR EXCLUIR PEGAR A COLUNA ATAULIAZACAO" & Chr(13)
   sql = sql & "     DA TABELA POR ISTO TEM O CAMPO SEMLOG, ESTE PROCESSO DE EXCLUSÃO SEMPRE QUE FOR EXCLUIR ANTES É COLOCADO" & Chr(13)
   sql = sql & "     UPDATE NA ATUALIZACAO" & Chr(13)
   sql = sql & "  */" & Chr(13)
   sql = sql & "  IF EXISTS(SELECT REGISTRO FROM DELETED) " & Chr(13)
   sql = sql & "  BEGIN" & Chr(13)
   sql = sql & "     SELECT @REGISTRO=REGISTRO,@SEMLOG=SEMLOG FROM INSERTED" & Chr(13)
   sql = sql & "     IF @REGISTRO IS NOT NULL" & Chr(13)
   sql = sql & "     BEGIN" & Chr(13)
   sql = sql & "        IF @SEMLOG=1 " & Chr(13)
   sql = sql & "           SET @GRAVAR=0" & Chr(13)
   sql = sql & "        SET @TIPOOPERACAO=1" & Chr(13)
   sql = sql & "     END" & Chr(13)
   sql = sql & "     ELSE" & Chr(13)
   sql = sql & "        SET @TIPOOPERACAO=2" & Chr(13)
   sql = sql & "  END " & Chr(13)
   sql = sql & "  ELSE" & Chr(13)
   sql = sql & "     SET @TIPOOPERACAO=0" & Chr(13)
   sql = sql & "  IF @GRAVAR=1 " & Chr(13)
   sql = sql & "  BEGIN" & Chr(13)
   sql = sql & "   IF @TIPOOPERACAO = 2    --EXCLUSÃO" & Chr(13)
   sql = sql & "  BEGIN" & Chr(13)
   sql = sql & "     INSERT INTO SAVELOG..MOVIM_EXT(" & Chr(13)
   sql = sql & "        TIPOOPERACAO,REGISTRO,DATA,USUARIO,TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE," & Chr(13)
   sql = sql & "        VALOR,CENTROCUSTO,DATACONSUMO,SEMBAIXA, SEQUENCIA, SEQUENCIALANCAMENTO, DATAOPERACAO, CODUSUARIO)" & Chr(13)
   sql = sql & "     SELECT   @TIPOOPERACAO,REGISTRO,LEFT(ATUALIZACAO,19),RTRIM(LTRIM(SUBSTRING(ATUALIZACAO,26,255)))," & Chr(13)
   sql = sql & "        TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALORUNITARIO,CENTROCUSTO," & Chr(13)
   sql = sql & "        DATACONSUMO,SEMBAIXA, SEQUENCIA, SEQUENCIALANCAMENTO, GETDATE(), USUARIO" & Chr(13)
   sql = sql & "     FROM  DELETED" & Chr(13)
   sql = sql & "  END" & Chr(13)
   sql = sql & "  ELSE --ALTERAÇÃO E EXCLUSÃO" & Chr(13)
   sql = sql & "     INSERT INTO SAVELOG..MOVIM_EXT(" & Chr(13)
   sql = sql & "        TIPOOPERACAO,REGISTRO,DATA,USUARIO,TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALOR," & Chr(13)
   sql = sql & "        CENTROCUSTO,DATACONSUMO,SEMBAIXA, SEQUENCIA, SEQUENCIALANCAMENTO, DATAOPERACAO, CODUSUARIO)" & Chr(13)
   sql = sql & "     SELECT   @TIPOOPERACAO,REGISTRO,LEFT(ATUALIZACAO,19),RTRIM(LTRIM(SUBSTRING(ATUALIZACAO,26,255)))," & Chr(13)
   sql = sql & "        TIPOLANCAMENTO,PROCEDIMENTO,PROCEDIMENTONOME,QUANTIDADE,VALORUNITARIO,CENTROCUSTO," & Chr(13)
   sql = sql & "        DATACONSUMO,SEMBAIXA, SEQUENCIA, SEQUENCIALANCAMENTO, GETDATE(), USUARIO" & Chr(13)
   sql = sql & "     FROM  INSERTED" & Chr(13)
   sql = sql & "  END" & Chr(13)
   Banco.Execute sql
   
   sql = ""
   sql = sql & "DROP PROCEDURE SP_LOG_MOVIMENTACAO"
   sql = sql & "CREATE PROCEDURE SP_LOG_MOVIMENTACAO" & Chr(13)
   sql = sql & "  @REGISTRO      AS INT," & Chr(13)
   sql = sql & "  @DATAINICIAL   AS DATETIME," & Chr(13)
   sql = sql & "  @DATAFINAL     AS DATETIME," & Chr(13)
   sql = sql & "  @TIPOREGISTRO  AS INT," & Chr(13)
   sql = sql & "  @IP            AS VARCHAR(100)," & Chr(13)
   sql = sql & "  @TIPOLANCAMENTO AS INT," & Chr(13)
   sql = sql & "  @PROCEDIMENTO  AS INT" & Chr(13)
   sql = sql & "WITH ENCRYPTION" & Chr(13)
   sql = sql & "AS" & Chr(13)
   sql = sql & "  -- ESTA SP FAZ O LOG DAS PASSAGEM" & Chr(13)
   sql = sql & "  --1 = INTERNO" & Chr(13)
   sql = sql & "  --2 = AMBULATORIAL" & Chr(13)
   sql = sql & "  --3 = EXTERNO" & Chr(13)
   sql = sql & "  DELETE FROM LOG_MOVIMENTACAO" & Chr(13)
   sql = sql & "  WHERE IP = @IP " & Chr(13)
   sql = sql & "  --INTERNO" & Chr(13)
   sql = sql & "  IF @TIPOREGISTRO = 1" & Chr(13)
   sql = sql & "  BEGIN" & Chr(13)
   sql = sql & "     INSERT INTO LOG_MOVIMENTACAO (" & Chr(13)
   sql = sql & "        TIPOLANCAMENTO, TIPOOPERACAO, REGISTRO, DATAOPERACAO, USUARIO, PROCEDIMENTO, PROCEDIMENTONOME, " & Chr(13)
   sql = sql & "        QUANTIDADE, QUANTIDADEDISPENSADA, VALOR, CENTROCUSTO, DESCRICAOCENTROCUSTO, DATACONSUMO, " & Chr(13)
   sql = sql & "        SEMBAIXA, IP, CODUSUARIO)" & Chr(13)
   sql = sql & "     SELECT (CASE WHEN A.TIPOLANCAMENTO = 1 THEN 'MAT/MED'" & Chr(13)
   sql = sql & "               WHEN A.TIPOLANCAMENTO = 2 THEN 'DIAGNOSTICO'" & Chr(13)
   sql = sql & "               WHEN A.TIPOLANCAMENTO = 3 THEN 'HONORÁRIO'" & Chr(13)
   sql = sql & "                                   ELSE 'INSUMO' END) AS TIPOLANCAMENTO," & Chr(13)
   sql = sql & "           (CASE WHEN A.TIPOOPERACAO = 0    THEN 'INCLUSÃO'" & Chr(13)
   sql = sql & "                WHEN A.TIPOOPERACAO = 2    THEN 'EXCLUSÃO'     " & Chr(13)
   sql = sql & "                WHEN A.TIPOOPERACAO = 1 AND " & Chr(13)
   sql = sql & "                    A.TIPOLANCAMENTO <> 1 THEN 'ALTERAÇÃO'" & Chr(13)
   sql = sql & "                                    ELSE" & Chr(13)
   sql = sql & "                 (CASE WHEN ISNULL(A.QUANTIDADE,0) = 0 AND ISNULL(A.QUANTIDADEDISPENSADA,0) = 0 THEN 'EXCLUSÃO'" & Chr(13)
   sql = sql & "                                                                             ELSE 'ALTERAÇÃO' " & Chr(13)
   sql = sql & "                  END)" & Chr(13)
   sql = sql & "           END) AS TIPOOPERACAO, A.REGISTRO, A.DATAOPERACAO, A.USUARIO, A.PROCEDIMENTO, A.PROCEDIMENTONOME, " & Chr(13)
   sql = sql & "           A.QUANTIDADE, A.QUANTIDADEDISPENSADA, A.VALOR, A.CENTROCUSTO, B.DESCRICAO AS DESCRICAOCENTROCUSTO, " & Chr(13)
   sql = sql & "           A.DATACONSUMO, A.SEMBAIXA, @IP, CODUSUARIO" & Chr(13)
   sql = sql & "     FROM SAVELOG..MOVIM_INT A LEFT JOIN CENTROCUSTO B ON A.CENTROCUSTO = B.CENTROCUSTO" & Chr(13)
   sql = sql & "     --WHERE A.REGISTRO = 1044" & Chr(13)
   sql = sql & "     WHERE A.REGISTRO = @REGISTRO" & Chr(13)
   sql = sql & "     AND   A.DATAOPERACAO BETWEEN @DATAINICIAL AND @DATAFINAL" & Chr(13)
   sql = sql & "     AND   A.TIPOLANCAMENTO = (CASE WHEN @TIPOLANCAMENTO = 0 THEN A.TIPOLANCAMENTO ELSE @TIPOLANCAMENTO END)" & Chr(13)
   sql = sql & "     AND   A.PROCEDIMENTO = (CASE WHEN @PROCEDIMENTO = 0 THEN A.PROCEDIMENTO ELSE @PROCEDIMENTO END)" & Chr(13)
   sql = sql & "  END" & Chr(13)
   sql = sql & "  --AMBULATORIAL" & Chr(13)
   sql = sql & "  IF @TIPOREGISTRO = 2" & Chr(13)
   sql = sql & "  BEGIN" & Chr(13)
   sql = sql & "     INSERT INTO LOG_MOVIMENTACAO (" & Chr(13)
   sql = sql & "        TIPOLANCAMENTO, TIPOOPERACAO, REGISTRO, DATAOPERACAO, USUARIO, PROCEDIMENTO, PROCEDIMENTONOME, " & Chr(13)
   sql = sql & "        QUANTIDADE, QUANTIDADEDISPENSADA, VALOR, CENTROCUSTO, DESCRICAOCENTROCUSTO, DATACONSUMO, " & Chr(13)
   sql = sql & "        SEMBAIXA, IP, CODUSUARIO)" & Chr(13)
   sql = sql & "     SELECT (CASE WHEN A.TIPOLANCAMENTO = 1 THEN 'MAT/MED'" & Chr(13)
   sql = sql & "               WHEN A.TIPOLANCAMENTO = 2 THEN 'DIAGNOSTICO'" & Chr(13)
   sql = sql & "               WHEN A.TIPOLANCAMENTO = 3 THEN 'HONORÁRIO'" & Chr(13)
   sql = sql & "                                   ELSE 'INSUMO' END) AS TIPOLANCAMENTO," & Chr(13)
   sql = sql & "           (CASE WHEN A.TIPOOPERACAO = 0    THEN 'INCLUSÃO'" & Chr(13)
   sql = sql & "                WHEN A.TIPOOPERACAO = 2    THEN 'EXCLUSÃO'     " & Chr(13)
   sql = sql & "                WHEN A.TIPOOPERACAO = 1 AND " & Chr(13)
   sql = sql & "                    A.TIPOLANCAMENTO <> 1 THEN 'ALTERAÇÃO'" & Chr(13)
   sql = sql & "                                    ELSE" & Chr(13)
   sql = sql & "                 (CASE WHEN ISNULL(A.QUANTIDADE,0) = 0 AND ISNULL(A.QUANTIDADEDISPENSADA,0) = 0 THEN 'EXCLUSÃO'" & Chr(13)
   sql = sql & "                                                                             ELSE 'ALTERAÇÃO' " & Chr(13)
   sql = sql & "                  END)" & Chr(13)
   sql = sql & "           END) AS TIPOOPERACAO, A.REGISTRO, A.DATAOPERACAO, A.USUARIO, A.PROCEDIMENTO, A.PROCEDIMENTONOME, " & Chr(13)
   sql = sql & "           A.QUANTIDADE, A.QUANTIDADEDISPENSADA, A.VALOR, A.CENTROCUSTO, B.DESCRICAO AS DESCRICAOCENTROCUSTO, " & Chr(13)
   sql = sql & "           A.DATACONSUMO, A.SEMBAIXA, @IP, CODUSUARIO" & Chr(13)
   sql = sql & "     FROM SAVELOG..MOVIM_AMB A LEFT JOIN CENTROCUSTO B ON A.CENTROCUSTO = B.CENTROCUSTO" & Chr(13)
   sql = sql & "     --WHERE A.REGISTRO = 1044" & Chr(13)
   sql = sql & "     WHERE A.REGISTRO = @REGISTRO" & Chr(13)
   sql = sql & "     AND   A.DATAOPERACAO BETWEEN @DATAINICIAL AND @DATAFINAL" & Chr(13)
   sql = sql & "     AND   A.TIPOLANCAMENTO = (CASE WHEN @TIPOLANCAMENTO = 0 THEN A.TIPOLANCAMENTO ELSE @TIPOLANCAMENTO END)" & Chr(13)
   sql = sql & "     AND   A.PROCEDIMENTO = (CASE WHEN @PROCEDIMENTO = 0 THEN A.PROCEDIMENTO ELSE @PROCEDIMENTO END)" & Chr(13)
   sql = sql & "  END" & Chr(13)
   sql = sql & "  --EXTERNO" & Chr(13)
   sql = sql & "  IF @TIPOREGISTRO = 3" & Chr(13)
   sql = sql & "  BEGIN" & Chr(13)
   sql = sql & "     INSERT INTO LOG_MOVIMENTACAO (" & Chr(13)
   sql = sql & "        TIPOLANCAMENTO, TIPOOPERACAO, REGISTRO, DATAOPERACAO, USUARIO, PROCEDIMENTO, PROCEDIMENTONOME, " & Chr(13)
   sql = sql & "        QUANTIDADE, QUANTIDADEDISPENSADA, VALOR, CENTROCUSTO, DESCRICAOCENTROCUSTO, DATACONSUMO, " & Chr(13)
   sql = sql & "        SEMBAIXA, IP, CODUSUARIO)" & Chr(13)
   sql = sql & "     SELECT (CASE WHEN A.TIPOLANCAMENTO = 1 THEN 'MAT/MED'" & Chr(13)
   sql = sql & "               WHEN A.TIPOLANCAMENTO = 2 THEN 'DIAGNOSTICO'" & Chr(13)
   sql = sql & "               WHEN A.TIPOLANCAMENTO = 3 THEN 'HONORÁRIO'" & Chr(13)
   sql = sql & "                                   ELSE 'INSUMO' END) AS TIPOLANCAMENTO," & Chr(13)
   sql = sql & "           (CASE WHEN A.TIPOOPERACAO = 0    THEN 'INCLUSÃO'" & Chr(13)
   sql = sql & "                WHEN A.TIPOOPERACAO = 2    THEN 'EXCLUSÃO'     " & Chr(13)
   sql = sql & "                WHEN A.TIPOOPERACAO = 1 AND " & Chr(13)
   sql = sql & "                    A.TIPOLANCAMENTO <> 1 THEN 'ALTERAÇÃO'" & Chr(13)
   sql = sql & "                                    ELSE" & Chr(13)
   sql = sql & "                 (CASE WHEN ISNULL(A.QUANTIDADE,0) = 0 AND ISNULL(A.QUANTIDADEDISPENSADA,0) = 0 THEN 'EXCLUSÃO'" & Chr(13)
   sql = sql & "                                                                             ELSE 'ALTERAÇÃO' " & Chr(13)
   sql = sql & "                  END)" & Chr(13)
   sql = sql & "           END) AS TIPOOPERACAO, A.REGISTRO, A.DATAOPERACAO, A.USUARIO, A.PROCEDIMENTO, A.PROCEDIMENTONOME, " & Chr(13)
   sql = sql & "           A.QUANTIDADE, A.QUANTIDADEDISPENSADA, A.VALOR, A.CENTROCUSTO, B.DESCRICAO AS DESCRICAOCENTROCUSTO, " & Chr(13)
   sql = sql & "           A.DATACONSUMO, A.SEMBAIXA, @IP, CODUSUARIO " & Chr(13)
   sql = sql & "     FROM SAVELOG..MOVIM_EXT A LEFT JOIN CENTROCUSTO B ON A.CENTROCUSTO = B.CENTROCUSTO" & Chr(13)
   sql = sql & "     --WHERE A.REGISTRO = 1044" & Chr(13)
   sql = sql & "     WHERE A.REGISTRO = @REGISTRO" & Chr(13)
   sql = sql & "     AND   A.DATAOPERACAO BETWEEN @DATAINICIAL AND @DATAFINAL" & Chr(13)
   sql = sql & "     AND   A.TIPOLANCAMENTO = (CASE WHEN @TIPOLANCAMENTO = 0 THEN A.TIPOLANCAMENTO ELSE @TIPOLANCAMENTO END)" & Chr(13)
   sql = sql & "     AND   A.PROCEDIMENTO = (CASE WHEN @PROCEDIMENTO = 0 THEN A.PROCEDIMENTO ELSE @PROCEDIMENTO END)" & Chr(13)
   sql = sql & "  END" & Chr(13)
   Banco.Execute sql
   
   
   sql = ""
   sql = sql & "ALTER TABLE [dbo].[KITPROCEDIMENTO] ADD SITUACAO BIT NOT NULL CONSTRAINT [kitprocedimento_situacao] DEFAULT (1)" & Chr(13)
   Banco.Execute sql
         
   sql = ""
   sql = sql & "INSERT INTO [MENU] " & vbCrLf
   sql = sql & "SELECT Max(MENU) + 1,N'Cadastros de Prescrição',N'mnuMed_Cad',N' ',N'mnuMed_Cad' " & vbCrLf
   sql = sql & "       ,1,1, " & vbCrLf
   sql = sql & "       N'0609000000',N'mnuMed_Cad',1,NULL,NULL,NULL,NULL,NULL " & vbCrLf
   sql = sql & "FROM   MENU " & vbCrLf
   sql = sql & "WHERE  NOT EXISTS (SELECT * " & vbCrLf
   sql = sql & "                   FROM   MENU " & vbCrLf
   sql = sql & "                   WHERE  NOMESUBNOVO = 'mnuMed_Cad')"
   Banco.Execute sql
   
   sql = ""
   sql = sql & "INSERT INTO [MENU] " & vbCrLf
   sql = sql & "SELECT Max(MENU) + 1,N'Prescrição Médica Eletrônica', " & vbCrLf
   sql = sql & "       N'mnuEPrescricaoMedicaEletronica', " & vbCrLf
   sql = sql & "       N'2001/11/06 16:20:17 36864 ',N'mnuMed_Pme',1,1,N'0602000000', " & vbCrLf
   sql = sql & "       N'mnuMed_Pme',1,NULL,NULL,NULL,NULL,NULL " & vbCrLf
   sql = sql & "FROM   MENU " & vbCrLf
   sql = sql & "WHERE  NOT EXISTS (SELECT * " & vbCrLf
   sql = sql & "                   FROM   MENU " & vbCrLf
   sql = sql & "                   WHERE  NOMESUBNOVO = 'mnuMed_Pme')"
   Banco.Execute sql
   
   sql = ""
   sql = sql & "INSERT INTO [MENU] " & vbCrLf
   sql = sql & "SELECT Max(MENU) + 1,N'Permissão Prescrição',N'mnuPar_Seg_Pre',N' ', " & vbCrLf
   sql = sql & "       N'mnuPar_Seg_Pre',1 " & vbCrLf
   sql = sql & "       ,1,N'0103040000',N'mnuPar_Seg_Pre',1,NULL,NULL,NULL,NULL,NULL " & vbCrLf
   sql = sql & "FROM   MENU " & vbCrLf
   sql = sql & "WHERE  NOT EXISTS (SELECT * " & vbCrLf
   sql = sql & "                   FROM   MENU " & vbCrLf
   sql = sql & "                   WHERE  NOMESUBNOVO = 'mnuPar_Seg_Pre') "
   Banco.Execute sql
   
   sql = ""
   sql = sql & "ALTER TABLE CONVENIODATA_PERIODO " & vbCrLf
   sql = sql & "  ADD FATURA_GERADA_MANUAL BIT DEFAULT(1) "
   Banco.Execute sql
   
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes022014()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '022014'"
   Banco.Execute sql
   
   sql = "CREATE TABLE [dbo].[MOTIVO_AUDITORIA]( " & Chr(13)
   sql = sql & " [SEQUENCIA] [bigint] IDENTITY(1,1) NOT NULL, " & Chr(13)
   sql = sql & "    [DESCRICAO] [varchar](300) NOT NULL, " & Chr(13)
   sql = sql & "    [ATIVO] [int] NOT NULL, " & Chr(13)
   sql = sql & "    [ATUALIZACAO] [varchar](150) NOT NULL, " & Chr(13)
   sql = sql & "  CONSTRAINT [PK_MOTIVO_AUDITORIA] PRIMARY KEY CLUSTERED " & Chr(13)
   sql = sql & " ( " & Chr(13)
   sql = sql & "    [sequencia] Asc " & Chr(13)
   sql = sql & " ))"
   Banco.Execute sql
   
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes032014()
On Error GoTo Erro
   
   'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
   'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
   sql = ""
   sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '032014'"
   Banco.Execute sql
   
   sql = "ALTER TABLE PRES_TMP_MEDICACAO ADD EMERGENCIA INT"
   Banco.Execute sql
   
Exit Function
Erro:
   Resume Next
End Function

Public Function AtualizaMes042014()
On Error GoTo Erro
   
    'SEMPRE COLOCAR ESTE CODIGO NAS FUNÇÕES
    'DA COLUNA ULTIMAATUALIZACAO COM O MES E O ANO DO MES AO QUAL A FUNCAO PERTENCE (Ex.: '122011' referente ao mês de dezembro de 2011)
    sql = ""
    sql = sql & " UPDATE PARAMETRO SET ULTIMAATUALIZACAO =  '042014'"
    Banco.Execute sql

    'Se for o layout Dracena:
    'Haverá a possibilidade do usuário exportar informações de contas a pagar/receber
    'para o sistema Totvs
    If Layout = 53 Then
        sql = ""
        sql = sql & " INSERT INTO MENU("
        sql = sql & "     MENU,NOMECAPTION,NOMESUB,ATUALIZACAO,NOMESUBNOVO,ATIVADO,MODULO,HIERARQUIA,"
        sql = sql & "     NOMESUBAUX,NIVELVISIBILIDADE, NOMESUBALTERADO, HIERARQUIANOVA)"
        sql = sql & " SELECT MAX(MENU)+1,'Exportação Totvs', 'mnuFin_ETT', ' ', 'mnuFin_ETT',"
        sql = sql & "        1, 1, '0918000000', 'mnuFin_ETT', 1, 'mnuFin_Fin_ETT', '0918000000'"
        sql = sql & " FROM MENU "
        Banco.Execute sql
        
        sql = ""
        sql = sql & " ALTER TABLE FORNECEDORES ADD CODIGOTOTVS INTEGER"
        Banco.Execute sql
    End If
   
    Exit Function
Erro:
   Resume Next
End Function
