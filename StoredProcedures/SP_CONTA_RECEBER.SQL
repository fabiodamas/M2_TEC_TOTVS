	CREATE PROCEDURE [dbo].[SP_CONTA_RECEBER]
 ( @dataInicial AS VARCHAR(10) ,
   @dataFinal AS VARCHAR(10)  
 )
AS
BEGIN

SET nocount ON; 

DECLARE 
	-- Variáveis adicionais
	@mensagem								varchar(1000),
	@mensagem_detalhe				varchar(1000),

	-- Variáveis do Cabeçalho da Conta a Receber
	@referencia								varchar(10), 
	@data_transacao  						datetime, 
	@valor_total_lote						decimal(15, 2), 
	@tipo_cobranca							varchar(10), 
	@estabelecimento_externo				varchar(5), 
	@estabelecimento						varchar(5), 

	-- Variáveis do Cabeçalho da Conta a Receber
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

	-- Inicializando variáveis
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
	
	 -- Obtendo os dados da tabela mestre de lançamentos de contas a pagar
	SELECT 
		--	10, Referência
		lancamento,
		
		-- 8, 	Data de Transação                                            				
		datalancamento,
		
		-- 13, 	Valor Total do Lote                                         
		valortotal,
				
		-- 10,	Tipo de Cobrança:: Normal/Especial
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

	-- Valores da tabela mestre necessários para exportação
	FETCH next FROM lancamento_cursor INTO @referencia, @data_transacao, @valor_total_lote, @tipo_cobranca, @estabelecimento_externo, @estabelecimento

	WHILE @@FETCH_STATUS = 0 
		BEGIN 
			-- Montando e exbindo a linha do mestre	
			SELECT 
			-- 3, Tipo de Registro, Obrigatório
			'100'  +     

			-- 3, Brancos							 
			SPACE(3)   +  

			-- 10, Referência de Implantação, Obrigatório 							
			RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@referencia),10) +

			-- 8, Data de Transação, Obrigatório
			CASE WHEN @data_transacao is null THEN 
				SPACE(8)
			ELSE
				REPLACE(convert(varchar, @data_transacao, 103),'/','')
			END + -- 

			-- 1, Espaço em Branco
			SPACE(1)    +

			-- 13, Valor Total do Lote
			CASE WHEN 
					@valor_total_lote is null THEN SPACE(13)
			ELSE 
					RIGHT(REPLICATE('0',13) + REPLACE(REPLACE(CAST(@valor_total_lote AS VARCHAR(13)),'.',''),',','')   ,13)
			END +

			-- 3, Brancos
			SPACE(3)    +

			-- 10, Tipo de Cobrança:: Normal/Especial
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
			-- Este cursor será repetido para cada item mestre de contas a pagar
			DECLARE lancamento_vencimento_cursor CURSOR FOR              

			-- Selecionamento todos os detalhes do mestre
			-- Relacionando através do campo lancamento
 
			SELECT
				-- 9, Código do Cliente, Obrigatório													
				filancamentorecebimento.codigo as codigo,

				-- 3, Espécie Documento, Obrigatório 
				SPACE(3) as especie,

				-- 3, Série , Obrigatório 
				SPACE(3) as serie,


				-- 10, Código do Documento (*), Obrigatório 
				filancamentorecebimento.documento as documento,

				-- 08, Indicador Econômico, Obrigatório 
				SPACE(8) as indicador_economico,

				-- 2, Código da Parcela, Obrigatório
				-- SPACE(2) as codigo_parcela,
				-- 2, Código da Parcela, Obrigatório  
			  ROW_NUMBER() OVER(ORDER BY filancamentorecebimentovencimento.datavencimento)  as codigo_parcela, 

				-- 8, Data de Emissão do Documento, Obrigatório 
				filancamentorecebimento.dataemissao as data_emissao,

				-- 8, Data de Vencimento do Documento, Obrigatório
				filancamentorecebimentovencimento.datavencimento as data_vencimento,

				-- 8, Data Prevista de Pagamento 
				SPACE(8) as data_prevista,

				-- 8, Data do Desconto
				SPACE(8) as data_desconto,
 
				-- 11, Valor do Documento, Obrigatório
				filancamentorecebimentovencimento.valorvencimento as valor_vencimento,

				-- 11, Valor do Desconto 
				SPACE(11) as valor_desconto,

				-- 6, Valor Percentual do Desconto 
				SPACE(6) as valor_percentual,

				-- 4, Valor Percentual Multa Atraso 
				SPACE(4) as valor_percentual_multa_atraso,

				-- 3, Dias Carência Multa
				SPACE(3) as dias_carencia_multa,

				-- 3, Dias Carência Juros
				SPACE(3) as dias_carencia_juros,

				-- 11, Valor do Documento, Obrigatório
				SPACE(11) as valor_liquido_titulo,

				 -- 8, Código do Portador Externo 
				SPACE(8) as codigo_portador_externo,

				-- 8, Modalidade Portador Exteno 
				SPACE(8) as modalidade_portador_externo,

				-- 6, Código da Representante, Obrigatório
				SPACE(6) as codigo_representante,
				
				-- 5, Código da Primeira Instrução Bancária
				SPACE(5) as codigo_primeira_instrucao_bancaria,

				-- 4, Código da Segunda Instrução Bancária
				SPACE(4) as codigo_segunda_instrucao_bancaria,

				-- 5, Código do Portador
				SPACE(5) as codigo_portador,

				-- 3, Código da Carteira Bancária 
				SPACE(3) as codigo_carteira_bancaria,
				
				-- 8, Código do Motivo do Movimento
				SPACE(8) as codigo_motivo_movimento,
				
				-- 8, Código do Histórico Padrão 
				SPACE(8) as codigo_historico_padrao,

				-- 8, Código do Banco
				SPACE(8) as codigo_banco,
 			
				-- 10, Código da Agência
				SPACE(10) as codigo_agencia,

				-- 20, Código da Conta Corrente
				SPACE(20) as codigo_conta_corrente,
	 		
				-- 2, Dígito da Conta Corrente  
				SPACE(2) as digito_conta_corrente,

				-- 20, Número do Cartão de Crédito
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

				  -- Valores do detalhe necessários para exportação
				  FETCH next FROM lancamento_vencimento_cursor INTO @codigo_cliente, @codigo_especie_documento, @codigo_serie_documento, @codigo_documento, @codigo_indicador_economico, @codigo_parcela, @data_emissao_documento, @data_vencimento_titulo, @data_prevista_liquidacao, @data_desconto, @valor_documento, @valor_desconto, @porcentual_desconto, @valor_percentual_multa_atraso, @dias_carencia_multa, @dias_carencia_juros, @valor_liquido_titulo , @codigo_portador_externo, @modalidade_portador_externo, @codigo_representante, @codigo_primeira_instrucao_bancaria, @codigo_segunda_instrucao_bancaria, @codigo_portador,
				@codigo_carteira_bancaria, @codigo_motivo_movimento, @codigo_historico_padrao, @codigo_banco, @codigo_agencia, @codigo_conta_corrente, @digito_conta_corrente, @numero_cartao_Credito

				  WHILE @@FETCH_STATUS = 0 
					BEGIN 

			 

						-- Montando e exbindo o detalhe do contas a pagar
						SELECT  '200'  +  -- Tipo de Registro
																
																-- Espaço em branco
																SPACE(1) +

																-- 9, Código do Cliente, Obrigatório													
																CASE WHEN @codigo_cliente is null THEN 
																	REPLICATE('0', 9)
																ELSE
																	RIGHT(REPLICATE('0', 9)+CONVERT(VARCHAR(9),@codigo_cliente),9) 
																END +  -- 

																-- 3, Espécie Documento, Obrigatório 
																	' PA'
																+ -- 

																-- 3, Série , Obrigatório 
																CASE WHEN @codigo_serie_documento is null THEN 
																	SPACE(3)
																ELSE
																	RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@codigo_serie_documento),3) 
																END  + -- 

																-- 10, Código do Documento (*), Obrigatório 
																CASE WHEN @codigo_documento is null THEN 
																	SPACE(10)
																ELSE
																	RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@codigo_documento),10) 
																END + 

																-- 08, Indicador Econômico, Obrigatório 
																CASE WHEN @codigo_indicador_economico is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@codigo_indicador_economico),8) 
																END + 

																-- 2, Código da Parcela, Obrigatório
																CASE WHEN @codigo_parcela is null THEN 
																	REPLICATE('0', 2)
																ELSE
																	RIGHT(REPLICATE('0', 2)+CONVERT(VARCHAR(2),@codigo_parcela),2) 
																END +  -- 

																-- 8, Data de Emissão do Documento, Obrigatório 
																CASE WHEN @data_emissao_documento is null THEN 
																	SPACE(8)
																ELSE
																	REPLACE(convert(varchar, @data_emissao_documento, 103),'/','')
																END + -- 

																-- 8, Data de Vencimento do Documento, Obrigatório
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

																-- Espaço em branco
																SPACE(3) +												 

																-- 11, Valor do Documento, Obrigatório
																CASE WHEN @valor_documento is null THEN 
																	REPLICATE('0', 11)
																ELSE 
																	RIGHT(REPLICATE('0',11) + REPLACE(REPLACE(CAST(@valor_documento AS VARCHAR(11)),'.',''),',','')   ,11)
																END +

																-- Espaço em branco
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

																-- Espaço em branco
																SPACE(2) +

																-- 4, Valor Percentual Multa Atraso 
																CASE WHEN @valor_percentual_multa_atraso is null THEN 
																	SPACE(4)
																ELSE
																	RIGHT(SPACE(4)+CONVERT(VARCHAR(4),@valor_percentual_multa_atraso),4) 
																END + 

																-- 3, Dias Carência Multa
																CASE WHEN @dias_carencia_multa is null THEN 
																	SPACE(3)
																ELSE
																	RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@dias_carencia_multa),2) 
																END + 

															-- 3, Dias Carência Juros
																CASE WHEN @dias_carencia_juros is null THEN 
																	SPACE(3)
																ELSE
																	RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@dias_carencia_juros),2) 
																END + 

																-- Espaço em branco
																SPACE(3) +

																-- 11, Valor do Documento, Obrigatório
																CASE WHEN @valor_liquido_titulo is null THEN 
																	SPACE(11)
																ELSE 
																	RIGHT(REPLICATE('0',11) + REPLACE(REPLACE(CAST(@valor_liquido_titulo AS VARCHAR(11)),'.',''),',','')   ,11)
																END +

																 -- 8, Código do Portador Externo 
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

																-- 6, Código da Parcela, Obrigatório
																CASE WHEN @codigo_representante is null THEN 
																	SPACE(6)
																ELSE
																	RIGHT(SPACE(6)+CONVERT(VARCHAR(6),@codigo_representante),6) 
																END +

																-- 5, Código da Primeira Instrução Bancária
																CASE WHEN @codigo_primeira_instrucao_bancaria is null THEN 
																	SPACE(5)
																ELSE
																	RIGHT(SPACE(5)+CONVERT(VARCHAR(5),@codigo_primeira_instrucao_bancaria),5) 
																END +

																-- 4, Código da Segunda Instrução Bancária
																CASE WHEN @codigo_segunda_instrucao_bancaria is null THEN 
																	SPACE(4)
																ELSE
																	RIGHT(SPACE(4)+CONVERT(VARCHAR(4),@codigo_segunda_instrucao_bancaria),5) 
																END +

																-- 5, Código do Portador
																CASE WHEN @codigo_portador is null THEN 
																	SPACE(5)
																ELSE
																	RIGHT(SPACE(5)+CONVERT(VARCHAR(5),@codigo_portador),5) 
																END +

																-- Espaço em branco
																SPACE(3) +

																-- 3, Código da Carteira Bancária 
																CASE WHEN @codigo_carteira_bancaria is null THEN 
																	SPACE(3)
																ELSE
																	RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@codigo_carteira_bancaria),3) 
																END +
																
																-- 8, Código do Motivo do Movimento
																CASE WHEN @codigo_motivo_movimento is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@codigo_motivo_movimento),8) 
																END +
																
																-- 8, Código do Histórico Padrão 
																CASE WHEN @codigo_historico_padrao is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@codigo_historico_padrao),8) 
																END +

																-- 8, Código do Banco
																CASE WHEN @codigo_banco is null THEN 
																	SPACE(8)
																ELSE
																	RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@codigo_banco),8) 
																END +
																
																-- 10, Código da Agência
																CASE WHEN @codigo_agencia is null THEN 
																	SPACE(10)
																ELSE
																	RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@codigo_agencia),10) 
																END +

																-- 20, Código da Conta Corrente
																CASE WHEN @codigo_conta_corrente is null THEN 
																	SPACE(20)
																ELSE
																	RIGHT(SPACE(20)+CONVERT(VARCHAR(20),@codigo_conta_corrente),20) 
																END +		
																
																-- 2, Dígito da Conta Corrente  
																CASE WHEN @digito_conta_corrente is null THEN 
																	SPACE(2)
																ELSE
																	RIGHT(SPACE(2)+CONVERT(VARCHAR(2),@digito_conta_corrente),2) 
																END +

																-- 20, Número do Cartão de Crédito
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

				  -- Próximo lançamento
					FETCH next FROM lancamento_cursor INTO @referencia, @data_transacao, @valor_total_lote, @tipo_cobranca, @estabelecimento_externo, @estabelecimento
			  END 

CLOSE lancamento_cursor; 

DEALLOCATE lancamento_cursor; 
END
