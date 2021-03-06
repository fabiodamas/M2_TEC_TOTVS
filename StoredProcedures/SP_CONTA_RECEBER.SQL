	CREATE PROCEDURE [dbo].[SP_CONTA_RECEBER]
 ( @dataInicial AS VARCHAR(10) ,
   @dataFinal AS VARCHAR(10)  
 )
AS
BEGIN

SET nocount ON; 

DECLARE 
	-- Vari�veis adicionais
	@mensagem								varchar(1000),
	@mensagem_detalhe				varchar(1000),

	-- Vari�veis do Cabe�alho da Conta a Receber
	@referencia								varchar(10), 
	@data_transacao  						datetime, 
	@valor_total_lote						decimal(15, 2), 
	@tipo_cobranca							varchar(10), 
	@estabelecimento_externo				varchar(5), 
	@estabelecimento						varchar(5), 

	-- Vari�veis do Cabe�alho da Conta a Receber
	@codigo_cliente							integer, 
	@codigo_especie_documento				varchar(3),
	@codigo_serie_documento 				varchar(5),
	@codigo_documento 						varchar(18),
	@codigo_indicador_economico 			varchar(8),
	@codigo_parcela 						integer,
	@data_emissao_documento 				datetime,
	@data_vencimento_titulo 				datetime,
	@data_prevista_liquidacao 				datetime,
	@data_desconto 							datetime,
	@valor_documento 						money,
	@valor_desconto 						money,
	@porcentual_desconto 					integer,
	@valor_percentual_multa_atraso			varchar(10),
	@dias_carencia_multa 					integer,
	@dias_carencia_juros 					integer,
	@valor_liquido_titulo					money,
	@codigo_portador_externo 				varchar(10),
	@modalidade_portador_externo 			varchar(2),
	@codigo_representante				    integer,
	@codigo_primeira_instrucao_bancaria		varchar(5),
	@codigo_segunda_instrucao_bancaria		varchar(4),
	@codigo_portador						varchar(5),
	@codigo_carteira_bancaria				varchar(3),
	@codigo_motivo_movimento				varchar(8),
	@codigo_historico_padrao				varchar(8),
	@codigo_banco							varchar(8),
	@codigo_agencia							varchar(10),
	@codigo_conta_corrente					varchar(20),
	@digito_conta_corrente					varchar(2),
	@numero_cartao_Credito					varchar(20);				

	-- Inicializando vari�veis
	set @mensagem							=  null
	set @referencia							= null
	set @data_transacao  					= null
	set @valor_total_lote					= null
	set @tipo_cobranca						= null
	set @estabelecimento_externo			= null
	set @estabelecimento					= null
	set @codigo_cliente						= null
	set @codigo_especie_documento			= null
	set @codigo_serie_documento 			= null
	set @codigo_documento 					= null
	set @codigo_indicador_economico 		= null
	set @codigo_parcela 					= null
	set @data_emissao_documento 			= null
	set @data_vencimento_titulo 			= null
	set @data_prevista_liquidacao 			= null
	set @data_desconto 						= null
	set @valor_desconto 					= null
	set @porcentual_desconto 				= null
	set @valor_percentual_multa_atraso		= null
	set @dias_carencia_multa 				= null
	set @dias_carencia_juros 				= null
	set @valor_liquido_titulo				= null
	set @codigo_portador_externo 			= null
	set @modalidade_portador_externo 		= null
	set @codigo_representante				= null
	set @codigo_primeira_instrucao_bancaria	= null
	set @codigo_segunda_instrucao_bancaria	= null
	set @codigo_portador					= null
	set @codigo_carteira_bancaria			= null
	set @codigo_motivo_movimento			= null
	set @codigo_historico_padrao			= null
	set @codigo_banco						= null
	set @codigo_agencia						= null
	set @codigo_conta_corrente				= null
	set @digito_conta_corrente				= null
	set @numero_cartao_Credito				= null			
        
	DECLARE lancamento_cursor CURSOR FOR  
	
	 -- Obtendo os dados da tabela mestre de lan�amentos de contas a pagar
	SELECT 
		--	10, Refer�ncia
		lancamento,
		
		-- 8, 	Data de Transa��o                                            				
		datalancamento,
		
		-- 13, 	Valor Total do Lote                                         
		valortotal,
				
		-- 10,	Tipo de Cobran�a:: Normal/Especial
		space (10) as tipo_cobranca,
		
		-- 5, 	Estabelecimento Externo
		space (10) as estabelecimento_externo,
		
		-- 5, 	Estabelecimento 
		space (10) as estabelecimento		
	FROM
		filancamentorecebimento  
	WHERE
		filancamentorecebimento.lancamento IN (
													SELECT 
														lancamento
													FROM
														filancamentorecebimentovencimento 
													WHERE
														-- lancamento=179045
														DATAVENCIMENTO BETWEEN @dataInicial AND @dataFinal

												)
	ORDER BY datalancamento desc ;

	OPEN lancamento_cursor 

	-- Valores da tabela mestre necess�rios para exporta��o
	FETCH next FROM lancamento_cursor INTO @referencia, @data_transacao, @valor_total_lote, @tipo_cobranca, @estabelecimento_externo, @estabelecimento

	WHILE @@FETCH_STATUS = 0 
		BEGIN 
			-- Montando e exbindo a linha do mestre	
			SELECT 
			-- 3, Tipo de Registro, Obrigat�rio
			'100'  +     

			-- 3, Brancos							 
			SPACE(3)   +  

			-- 10, Refer�ncia de Implanta��o, Obrigat�rio 							
			RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@referencia),10) +

			-- 8, Data de Transa��o, Obrigat�rio
			CASE WHEN @data_transacao is null THEN 
				SPACE(8)
			ELSE
				REPLACE(convert(varchar, @data_transacao, 103),'/','')
			END + -- 

			-- 1, Espa�o em Branco
			SPACE(1)    +

			-- 13, Valor Total do Lote
			CASE WHEN 
					@valor_total_lote is null THEN SPACE(13)
			ELSE 
					RIGHT(REPLICATE('0',13) + REPLACE(REPLACE(CAST(@valor_total_lote AS VARCHAR(13)),'.',''),',','')   ,13)
			END +

			-- 3, Brancos
			SPACE(3)    +

			-- 10, Tipo de Cobran�a:: Normal/Especial
			CASE WHEN @tipo_cobranca is null THEN 
				SPACE(10)
			ELSE
				RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@tipo_cobranca),10) 
			END +  --										
			
			-- 5, Estabelecimento Externo
			CASE WHEN @estabelecimento_externo is null THEN 
				SPACE(5)
			ELSE
				RIGHT(SPACE(5)+CONVERT(VARCHAR(5),@estabelecimento_externo),5) 
			END +  --	

			-- 5, Estabelecimento			 
			CASE WHEN @estabelecimento is null THEN 
				SPACE(5)
			ELSE
				RIGHT(SPACE(5)+CONVERT(VARCHAR(5),@estabelecimento),5) 
			END 


			-- Declando o cursor interno
			-- Este cursor ser� repetido para cada item mestre de contas a pagar
			DECLARE lancamento_vencimento_cursor CURSOR FOR              

			-- Selecionamento todos os detalhes do mestre
			-- Relacionando atrav�s do campo lancamento
 
			SELECT
				-- 9, C�digo do Cliente, Obrigat�rio													
				filancamentorecebimento.codigo as codigo,

				-- 3, Esp�cie Documento, Obrigat�rio 
				SPACE(3) as especie,

				-- 3, S�rie , Obrigat�rio 
				SPACE(3) as serie,


				-- 10, C�digo do Documento (*), Obrigat�rio 
				filancamentorecebimento.documento as documento,

				-- 08, Indicador Econ�mico, Obrigat�rio 
				SPACE(8) as indicador_economico,

				-- 2, C�digo da Parcela, Obrigat�rio
				-- SPACE(2) as codigo_parcela,
				-- 2, C�digo da Parcela, Obrigat�rio  
			  ROW_NUMBER() OVER(ORDER BY filancamentorecebimentovencimento.datavencimento)  as codigo_parcela, 

				-- 8, Data de Emiss�o do Documento, Obrigat�rio 
				filancamentorecebimento.dataemissao as data_emissao,

				-- 8, Data de Vencimento do Documento, Obrigat�rio
				filancamentorecebimentovencimento.datavencimento as data_vencimento,

				-- 8, Data Prevista de Pagamento 
				SPACE(8) as data_prevista,

				-- 8, Data do Desconto
				SPACE(8) as data_desconto,
 
				-- 11, Valor do Documento, Obrigat�rio
				filancamentorecebimentovencimento.valorvencimento as valor_vencimento,

				-- 11, Valor do Desconto 
				SPACE(11) as valor_desconto,

				-- 6, Valor Percentual do Desconto 
				SPACE(6) as valor_percentual,

				-- 4, Valor Percentual Multa Atraso 
				SPACE(4) as valor_percentual_multa_atraso,

				-- 3, Dias Car�ncia Multa
				SPACE(3) as dias_carencia_multa,

				-- 3, Dias Car�ncia Juros
				SPACE(3) as dias_carencia_juros,

				-- 11, Valor do Documento, Obrigat�rio
				SPACE(11) as valor_liquido_titulo,

				 -- 8, C�digo do Portador Externo 
				SPACE(8) as codigo_portador_externo,

				-- 8, Modalidade Portador Exteno 
				SPACE(8) as modalidade_portador_externo,

				-- 6, C�digo da Representante, Obrigat�rio
				SPACE(6) as codigo_representante,
				
				-- 5, C�digo da Primeira Instru��o Banc�ria
				SPACE(5) as codigo_primeira_instrucao_bancaria,

				-- 4, C�digo da Segunda Instru��o Banc�ria
				SPACE(4) as codigo_segunda_instrucao_bancaria,

				-- 5, C�digo do Portador
				SPACE(5) as codigo_portador,

				-- 3, C�digo da Carteira Banc�ria 
				SPACE(3) as codigo_carteira_bancaria,
				
				-- 8, C�digo do Motivo do Movimento
				SPACE(8) as codigo_motivo_movimento,
				
				-- 8, C�digo do Hist�rico Padr�o 
				SPACE(8) as codigo_historico_padrao,

				-- 8, C�digo do Banco
				SPACE(8) as codigo_banco,
 			
				-- 10, C�digo da Ag�ncia
				SPACE(10) as codigo_agencia,

				-- 20, C�digo da Conta Corrente
				SPACE(20) as codigo_conta_corrente,
	 		
				-- 2, D�gito da Conta Corrente  
				SPACE(2) as digito_conta_corrente,

				-- 20, N�mero do Cart�o de Cr�dito
				SPACE(20) as numero_cartao_Credito

				
					FROM   filancamentorecebimento 
								 INNER JOIN filancamentorecebimentovencimento 
												 ON filancamentorecebimento.lancamento = 
														filancamentorecebimentovencimento.lancamento 
								 LEFT OUTER JOIN filancamentorecebimentorecebido 
															ON filancamentorecebimentovencimento.lancamento = 
					filancamentorecebimentorecebido.lancamento 
					AND filancamentorecebimentovencimento.datavencimento = 
					filancamentorecebimentorecebido.datavencimento 
					INNER JOIN v_recuperatiporecebimento 
					ON filancamentorecebimento.tipopaciente = 
					v_recuperatiporecebimento.tipopaciente 
					AND filancamentorecebimento.tipo = v_recuperatiporecebimento.tipo 
					AND filancamentorecebimento.codigo = v_recuperatiporecebimento.codigo 
					AND filancamentorecebimento.lancamento = 
					v_recuperatiporecebimento.lancamento 
					WHERE  filancamentorecebimento.lancamento=@referencia 
								 AND ( filancamentorecebimento.transferencia = 0 
												OR filancamentorecebimento.transferencia IS NULL ) 
								 AND ( ( filancamentorecebimentorecebido.datarecebimento IS NOT NULL ) 
												OR filancamentorecebimentorecebido.datarecebimento IS NULL ) 
								 AND filancamentorecebimentorecebido.datarecebimento IS NULL 
					GROUP  BY filancamentorecebimento.lancamento,
										filancamentorecebimento.documento, 
										filancamentorecebimento.tipopaciente, 
										filancamentorecebimento.valortotal, 
										filancamentorecebimentovencimento.datavencimento, 
										filancamentorecebimentovencimento.valorvencimento, 
										filancamentorecebimento.integracontabilidade, 
										filancamentorecebimento.tipo, 
										filancamentorecebimentorecebido.datarecebimento, 
										filancamentorecebimento.codigo, 
										v_recuperatiporecebimento.nome, 
										filancamentorecebimento.observacao ,
							filancamentorecebimento.dataemissao

					HAVING filancamentorecebimentovencimento.valorvencimento - ( 
												Isnull(Sum(filancamentorecebimentorecebido.valorrecebimento), 0) 
																																			 + Isnull(Sum( 
												filancamentorecebimentorecebido.desconto), 0) ) > 0 
					ORDER  BY codigo_parcela  


				  OPEN lancamento_vencimento_cursor 

				  -- Valores do detalhe necess�rios para exporta��o
				  FETCH next FROM lancamento_vencimento_cursor INTO @codigo_cliente, @codigo_especie_documento, @codigo_serie_documento, @codigo_documento, @codigo_indicador_economico, @codigo_parcela, @data_emissao_documento, @data_vencimento_titulo, @data_prevista_liquidacao, @data_desconto, @valor_documento, @valor_desconto, @porcentual_desconto, @valor_percentual_multa_atraso, @dias_carencia_multa, @dias_carencia_juros, @valor_liquido_titulo , @codigo_portador_externo, @modalidade_portador_externo, @codigo_representante, @codigo_primeira_instrucao_bancaria, @codigo_segunda_instrucao_bancaria, @codigo_portador,
				@codigo_carteira_bancaria, @codigo_motivo_movimento, @codigo_historico_padrao, @codigo_banco, @codigo_agencia, @codigo_conta_corrente, @digito_conta_corrente, @numero_cartao_Credito

				  WHILE @@FETCH_STATUS = 0 
					BEGIN 

			 

						-- Montando e exbindo o detalhe do contas a pagar
						SELECT  '200'  +  -- Tipo de Registro
																
																-- Espa�o em branco
																SPACE(1) +

																-- 9, C�digo do Cliente, Obrigat�rio													
																CASE WHEN @codigo_cliente is null THEN 
																	REPLICATE('0', 9)
																ELSE
																	RIGHT(REPLICATE('0', 9)+CONVERT(VARCHAR(9),@codigo_cliente),9) 
																END +  -- 

																-- 3, Esp�cie Documento, Obrigat�rio 
																	' PA'
																+ -- 

																-- 3, S�rie , Obrigat�rio 
																CASE WHEN @codigo_serie_documento is null THEN 
																	SPACE(3)
																ELSE
																	RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@codigo_serie_documento),3) 
																END  + -- 

																-- 10, C�digo do Documento (*), Obrigat�rio 
																CASE WHEN @codigo_documento is null THEN 
																	SPACE(10)
																ELSE
																	RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@codigo_documento),10) 
																END + 

																-- 08, Indicador Econ�mico, Obrigat�rio 
																CASE WHEN @codigo_indicador_economico is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@codigo_indicador_economico),8) 
																END + 

																-- 2, C�digo da Parcela, Obrigat�rio
																CASE WHEN @codigo_parcela is null THEN 
																	REPLICATE('0', 2)
																ELSE
																	RIGHT(REPLICATE('0', 2)+CONVERT(VARCHAR(2),@codigo_parcela),2) 
																END +  -- 

																-- 8, Data de Emiss�o do Documento, Obrigat�rio 
																CASE WHEN @data_emissao_documento is null THEN 
																	SPACE(8)
																ELSE
																	REPLACE(convert(varchar, @data_emissao_documento, 103),'/','')
																END + -- 

																-- 8, Data de Vencimento do Documento, Obrigat�rio
																CASE WHEN @data_vencimento_titulo is null THEN 
																	SPACE(8)
																ELSE
																	REPLACE(convert(varchar, @data_vencimento_titulo, 103),'/','')
																END + -- 

																-- 8, Data Prevista de Pagamento 
																CASE WHEN @data_prevista_liquidacao is null THEN 
																	SPACE(8)
																ELSE
																	REPLACE(convert(varchar, @data_prevista_liquidacao, 103),'/','')

																END + -- 

																-- 8, Data do Desconto
																CASE WHEN @data_desconto is null THEN 
																	SPACE(8)
																ELSE
																	REPLACE(convert(varchar, @data_desconto, 103),'/','')
																END + -- 

																-- Espa�o em branco
																SPACE(3) +												 

																-- 11, Valor do Documento, Obrigat�rio
																CASE WHEN @valor_documento is null THEN 
																	REPLICATE('0', 11)
																ELSE 
																	RIGHT(REPLICATE('0',11) + REPLACE(REPLACE(CAST(@valor_documento AS VARCHAR(11)),'.',''),',','')   ,11)
																END +

																-- Espa�o em branco
																SPACE(3) +

																-- 11, Valor do Desconto 
																CASE WHEN @valor_desconto is null THEN 
																	SPACE(11)
																ELSE 
																	RIGHT(REPLICATE('0',11) + REPLACE(REPLACE(CAST(@valor_desconto AS VARCHAR(11)),'.',''),',','')   ,11)
																END +

																-- 6, Valor Percentual do Desconto 
																CASE WHEN @porcentual_desconto is null THEN 
																	SPACE(6)
																ELSE
																	RIGHT(SPACE(6)+CONVERT(VARCHAR(6),@porcentual_desconto),6) 
																END + 	

																-- Espa�o em branco
																SPACE(2) +

																-- 4, Valor Percentual Multa Atraso 
																CASE WHEN @valor_percentual_multa_atraso is null THEN 
																	SPACE(4)
																ELSE
																	RIGHT(SPACE(4)+CONVERT(VARCHAR(4),@valor_percentual_multa_atraso),4) 
																END + 

																-- 3, Dias Car�ncia Multa
																CASE WHEN @dias_carencia_multa is null THEN 
																	SPACE(3)
																ELSE
																	RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@dias_carencia_multa),2) 
																END + 

															-- 3, Dias Car�ncia Juros
																CASE WHEN @dias_carencia_juros is null THEN 
																	SPACE(3)
																ELSE
																	RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@dias_carencia_juros),2) 
																END + 

																-- Espa�o em branco
																SPACE(3) +

																-- 11, Valor do Documento, Obrigat�rio
																CASE WHEN @valor_liquido_titulo is null THEN 
																	SPACE(11)
																ELSE 
																	RIGHT(REPLICATE('0',11) + REPLACE(REPLACE(CAST(@valor_liquido_titulo AS VARCHAR(11)),'.',''),',','')   ,11)
																END +

																 -- 8, C�digo do Portador Externo 
																CASE WHEN @codigo_portador_externo is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(@codigo_portador_externo)+CONVERT(VARCHAR(8),@codigo_portador_externo),8) 
																END + 

																-- 2, Modalidade Portador Exteno 
																CASE WHEN @modalidade_portador_externo is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@modalidade_portador_externo),8) 
																END +

																-- 6, C�digo da Parcela, Obrigat�rio
																CASE WHEN @codigo_representante is null THEN 
																	SPACE(6)
																ELSE
																	RIGHT(SPACE(6)+CONVERT(VARCHAR(6),@codigo_representante),6) 
																END +

																-- 5, C�digo da Primeira Instru��o Banc�ria
																CASE WHEN @codigo_primeira_instrucao_bancaria is null THEN 
																	SPACE(5)
																ELSE
																	RIGHT(SPACE(5)+CONVERT(VARCHAR(5),@codigo_primeira_instrucao_bancaria),5) 
																END +

																-- 4, C�digo da Segunda Instru��o Banc�ria
																CASE WHEN @codigo_segunda_instrucao_bancaria is null THEN 
																	SPACE(4)
																ELSE
																	RIGHT(SPACE(4)+CONVERT(VARCHAR(4),@codigo_segunda_instrucao_bancaria),5) 
																END +

																-- 5, C�digo do Portador
																CASE WHEN @codigo_portador is null THEN 
																	SPACE(5)
																ELSE
																	RIGHT(SPACE(5)+CONVERT(VARCHAR(5),@codigo_portador),5) 
																END +

																-- Espa�o em branco
																SPACE(3) +

																-- 3, C�digo da Carteira Banc�ria 
																CASE WHEN @codigo_carteira_bancaria is null THEN 
																	SPACE(3)
																ELSE
																	RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@codigo_carteira_bancaria),3) 
																END +
																
																-- 8, C�digo do Motivo do Movimento
																CASE WHEN @codigo_motivo_movimento is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@codigo_motivo_movimento),8) 
																END +
																
																-- 8, C�digo do Hist�rico Padr�o 
																CASE WHEN @codigo_historico_padrao is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@codigo_historico_padrao),8) 
																END +

																-- 8, C�digo do Banco
																CASE WHEN @codigo_banco is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@codigo_banco),8) 
																END +
																
																-- 10, C�digo da Ag�ncia
																CASE WHEN @codigo_agencia is null THEN 
																	SPACE(10)
																ELSE
																	RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@codigo_agencia),10) 
																END +

																-- 20, C�digo da Conta Corrente
																CASE WHEN @codigo_conta_corrente is null THEN 
																	SPACE(20)
																ELSE
																	RIGHT(SPACE(20)+CONVERT(VARCHAR(20),@codigo_conta_corrente),20) 
																END +		
																
																-- 2, D�gito da Conta Corrente  
																CASE WHEN @digito_conta_corrente is null THEN 
																	SPACE(2)
																ELSE
																	RIGHT(SPACE(2)+CONVERT(VARCHAR(2),@digito_conta_corrente),2) 
																END +

																-- 20, N�mero do Cart�o de Cr�dito
																CASE WHEN @numero_cartao_Credito is null THEN 
																	SPACE(20)
																ELSE
																	RIGHT(SPACE(20)+CONVERT(VARCHAR(20),@numero_cartao_Credito),20) 
																END
							
				  FETCH next FROM lancamento_vencimento_cursor INTO @codigo_cliente, @codigo_especie_documento, @codigo_serie_documento, @codigo_documento, @codigo_indicador_economico, @codigo_parcela, @data_emissao_documento, @data_vencimento_titulo, @data_prevista_liquidacao, @data_desconto, @valor_documento, @valor_desconto, @porcentual_desconto, @valor_percentual_multa_atraso, @dias_carencia_multa, @dias_carencia_juros, @valor_liquido_titulo , @codigo_portador_externo, @modalidade_portador_externo, @codigo_representante, @codigo_primeira_instrucao_bancaria, @codigo_segunda_instrucao_bancaria, @codigo_portador,
				@codigo_carteira_bancaria, @codigo_motivo_movimento, @codigo_historico_padrao, @codigo_banco, @codigo_agencia, @codigo_conta_corrente, @digito_conta_corrente, @numero_cartao_Credito


					END 

				  CLOSE lancamento_vencimento_cursor 

				  DEALLOCATE lancamento_vencimento_cursor 

				  -- Pr�ximo lan�amento
					FETCH next FROM lancamento_cursor INTO @referencia, @data_transacao, @valor_total_lote, @tipo_cobranca, @estabelecimento_externo, @estabelecimento
			  END 

CLOSE lancamento_cursor; 

DEALLOCATE lancamento_cursor; 
END
