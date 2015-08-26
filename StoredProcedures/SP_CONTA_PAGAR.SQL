CREATE PROCEDURE [dbo].[SP_CONTA_PAGAR]
 ( @dataInicial AS VARCHAR(10) ,
   @dataFinal AS VARCHAR(10)  
 )
AS
BEGIN

SET nocount ON; 

DECLARE @lancamento											varchar(255), 
        @valortotal											decimal(15, 2), 
        @datalancamento  								datetime, 
        @mensagem												varchar(500),
        @lancamento_vencimento					varchar(255),
				@codigo_fornecedor							integer, 
				@especie 												varchar(3),
				@serie 													varchar(5),
				@documento 											varchar(18),
				@codigo_parcela 								integer,
				@data_emissao 									datetime,
				@data_vencimento 								datetime,
				@data_prevista 									datetime,
				@finalidade_externa 						varchar(10),
				@valor_vencimento 							money,
				@data_desconto 									datetime,
				@valor_desconto 								money,
				@valor_percentual 							integer,
				@numero_dias_atraso							integer,
				@valor_juros_dia_atraso 				money,
				@percentual_juros_dia_atraso   	varchar(10),
				@valor_percentual_multa_atraso 	varchar(10),
				@codigo_portador_externo 		 		varchar(10),
				@modalidade_portador_externo 		varchar(2),
				@codigo_apolice_seguro 					varchar(12),
				@codigo_seguradora 							varchar(12),
				@codigo_arrendador 							varchar(6),
				@codigo_contrato 								varchar(12),
				@codigo_portador								varchar(5),
				@codigo_indicador_economico 		varchar(8),
				@codigo_forma_pagamento 				varchar(3),
				@descricao_historico_padrao 		varchar(200),
				@valor_cotacao_moeda_corrente 	varchar(17),
				@ordem_investimento 						varchar(8);

				SET @valor_percentual_multa_atraso 	=  null
				SET @lancamento						 					= null  
				SET @valortotal						 					= null
				SET @datalancamento  				 				= null
				SET @mensagem						 						= null
				SET @lancamento_vencimento			 		= null
				SET @codigo_fornecedor				 			= null
				SET @especie 						 						= null
				SET @serie 							 						= null
				SET @documento 						 					= null
				SET @codigo_parcela 				 				= null
				SET @data_emissao 					 				= null
				SET @data_vencimento 				 				= null
				SET @data_prevista 					 				= null
				SET @finalidade_externa 			 			= null
				SET @valor_vencimento 				 			= null
				SET @data_desconto 					 				= null
				SET @valor_desconto 				 				= null
				SET @valor_percentual 				 			= null
				SET @numero_dias_atraso				 			= null
				SET @valor_juros_dia_atraso 		 		= null
				SET @percentual_juros_dia_atraso    = null
				SET @codigo_portador_externo 		 		= null
				SET @modalidade_portador_externo 	 	= null
				SET @codigo_apolice_seguro 			 		= null
				SET @codigo_seguradora 				 			= null
				SET @codigo_arrendador 				 			= null
				SET @codigo_contrato 				 				= null
				SET @codigo_portador				 				= null
				SET @codigo_indicador_economico 	 	= null
				SET @codigo_forma_pagamento 		 		= null
				SET @descricao_historico_padrao 	 	= null
				SET @valor_cotacao_moeda_corrente  	= null
				SET @ordem_investimento 			 			= null
        
DECLARE lancamento_cursor CURSOR FOR  
  
 -- Obtendo os dados da tabela mestre de lançamentos de contas a pagar
SELECT 			FILANCAMENTO.lancamento,
						FILANCAMENTO.valortotal,
						FILANCAMENTO.datalancamento
FROM
						FILANCAMENTO FILANCAMENTO
WHERE
						FILANCAMENTO.LANCAMENTO IN (
							SELECT
								LANCAMENTO
							FROM
								FILANCAMENTOVENCIMENTO
							WHERE
								DATAVENCIMENTO BETWEEN @dataInicial AND @dataFinal
						)														 
 ORDER BY datalancamento desc ;

OPEN lancamento_cursor 

-- Valores da tabela mestre necessários para exportação
FETCH next FROM lancamento_cursor INTO @lancamento, @valortotal, @datalancamento

WHILE @@FETCH_STATUS = 0 
  BEGIN 
	  -- Montando e exbindo a linha do mestre	
      SELECT
													-- 3, Tipo de Registro, Obrigatório
													'100'  +     

													-- 3, Brancos							 
													SPACE(3)   +  

													-- 10, Referência de Implantação, Obrigatório 							
													RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@lancamento),10) +

													-- 8, Data de Transação, Obrigatório
												  CASE WHEN @datalancamento is null THEN 
														SPACE(8)
													ELSE
														REPLACE(convert(varchar, @datalancamento, 103),'/','')
													END + -- 

													-- 14, Valor Total do Lote de Implantação, Obrigatório												  
													CASE WHEN 
															@valortotal is null THEN SPACE(14)
													ELSE 
															RIGHT(REPLICATE('0',14) + REPLACE(REPLACE(CAST(@valortotal AS VARCHAR(14)),'.',''),',','')   ,14)
													END +
													
													-- 5, Estabelecimento, Obrigatório
													-- O código de valor 04 corresponde ao Pronto Atendimento de Dracena no Totvs.
													'00004'    +

													-- 5, Estabelecimento Externo				 
													SPACE(5)

      AS mensagem

      -- Declando o cursor interno
      -- Este cursor será repetido para cada item mestre de contas a pagar
      DECLARE lancamento_vencimento_cursor CURSOR FOR              

		-- Selecionamento todos os detalhes do mestre
		-- Relacionando através do campo lancamento
 
SELECT
		-- 9, Código do Fornecedor, Obrigatório
		fornecedores.codigototvs codigo_fornecedor, 	
			
		-- 3, Espécie Documento , Obrigatório 
		space(3) as especie,

		-- 3, Série do Documento (*), Obrigatório
		filancamento.serie,                                 

		-- 10, Código do Documento (*), Obrigatório  
		filancamento.documento, 													   	

		-- 2, Código da Parcela, Obrigatório  
		ROW_NUMBER() OVER(ORDER BY filancamentovencimento.DATAVENCIMENTO)  as codigo_parcela,            				

		-- 8, Data de Emissão do Documento, Obrigatório     
		filancamento.dataemissao as data_emissao, 	

		-- 8, Data de Vencimento do Documento, Obrigatório  
		filancamentovencimento.datavencimento as data_vencimento,                  

		-- 10, Data Prevista de Pagamento
		null as data_prevista,

		-- 10, Finalidade Econômica Externa
		space(10) as finalidade_externa,

		-- 11, Valor do Documento, Obrigatório               
		filancamentovencimento.valorvencimento as valor_vencimento,                       

		-- 8, Data do Desconto 
		null as data_desconto,

		-- 11, Valor do Desconto 
		SPACE(11) as valor_desconto,

		-- 4, Valor Percentual do Desconto 
		SPACE(4)  as valor_percentual,

		-- 2, Número de Dias de Atraso 	
		SPACE(2)  as numero_dias_atraso,

		-- 11, Valor Juros por Dia Atraso 
		SPACE(11)  as valor_juros_dia_atraso,

		-- 4, Percentual de Juros por Dia de Atraso 
		SPACE(4)  as percentual_juros_dia_atraso,

		-- 4, Valor Percentual Multa Atraso 
		SPACE(4)  as valor_percentual_multa_atraso,

		-- 10, Código do Portador Externo 
		SPACE(10)  as codigo_portador_externo,

		-- 2, Modalidade Portador Exteno 
		SPACE(2)  as modalidade_portador_externo,

		-- 12, Codigo da Apólice de Seguro 
		SPACE(12)  as codigo_apolice_seguro, 

		-- 12, Código da Seguradora 
		SPACE(12)  as codigo_seguradora,

		-- 6, Código do Arrendador 
		SPACE(6)  as codigo_arrendador,

		-- 12, Código Contrato de Leasing 	
		SPACE(12)  as codigo_contrato,

		-- 5, Código do Portador 
		SPACE(5)  as codigo_portador,

		-- 8, Código do Indicador Econômico 
		SPACE(8)  as codigo_indicador_economico,

		-- 3, Código da Forma de Pagamento 
		SPACE(3)  as codigo_forma_pagamento,

		-- 200, Descrição do Histórico Padrão
		SPACE(200)  as descricao_historico_padrao,

		-- 17, Valor Cotação em Moeda Corrente 
		SPACE(17)  as valor_cotacao_moeda_corrente,

		-- 8, Indicar Ordem Investimento Integração com o EMS 2
		SPACE(8)  as ordem_investimento
	
FROM   filancamento  
			 INNER JOIN fornecedores 
							 ON filancamento.credor = Fornecedores.Fornecedor
       INNER JOIN filancamentovencimento filancamentovencimento 
               ON filancamentovencimento.lancamento = filancamento.lancamento 

WHERE
filancamento.lancamento=@lancamento
GROUP  BY 
					filancamento.dataemissao,
					filancamento.serie,
					fornecedores.codigototvs,
					filancamento.lancamento, 
          filancamento.documento, 
          filancamento.tipopagamento, 
          filancamento.valor_desconto, 
          filancamento.observacao, 
          filancamento.valortotal, 
          filancamento.pedido, 
          filancamento.integracontabilidade, 
          filancamento.lancamentodireto, 
          filancamento.valorimposto, 
          filancamento.valordevolvido, 
          filancamento.bancopagamento, 
          filancamento.pagamentoterceiro, 
          filancamento.importado_operadora, 
          filancamento.valormulta, 
          filancamento.valorjuros, 
          filancamentovencimento.datavencimento, 
          filancamentovencimento.valorvencimento 
ORDER  BY codigo_parcela 


      OPEN lancamento_vencimento_cursor 

      -- Valores do detalhe necessários para exportação
      FETCH next FROM lancamento_vencimento_cursor INTO @codigo_fornecedor, @especie,@serie, @documento, @codigo_parcela, @data_emissao,
				@data_vencimento, @data_prevista, @finalidade_externa, @valor_vencimento, @data_desconto, @valor_desconto,
				@valor_percentual, @numero_dias_atraso, @valor_juros_dia_atraso, @percentual_juros_dia_atraso, 
				@valor_percentual_multa_atraso,	@codigo_portador_externo, @modalidade_portador_externo, @codigo_apolice_seguro, 
				@codigo_seguradora, @codigo_arrendador, @codigo_contrato, @codigo_portador, @codigo_indicador_economico, 
				@codigo_forma_pagamento, @descricao_historico_padrao, @valor_cotacao_moeda_corrente , @ordem_investimento



      WHILE @@FETCH_STATUS = 0 
        BEGIN 

			-- Montando e exbindo o detalhe do contas a pagar
			SELECT '200'  +  -- Tipo de Registro
													-- + @lancamento_vencimento 
													
													-- 9, Código do Fornecedor, Obrigatório													
													CASE WHEN @codigo_fornecedor is null THEN 
													 	REPLICATE('0',9)
													ELSE
													 	RIGHT(REPLICATE('0',9)+CONVERT(VARCHAR(9),@codigo_fornecedor),9) 
													END +  -- 

													-- 3, Espécie Documento, Obrigatório 
														' PA'
													 + -- 

													--3, série do documento 
													CASE WHEN @serie is null THEN 
														SPACE(3)
													ELSE
														RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@serie),3) 
													END + 

													-- 10, Código do Documento (*), Obrigatório 
													CASE WHEN @documento is null THEN 
														SPACE(10)
													ELSE
														RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@documento),10) 
													END + 

													-- 2, Código da Parcela, Obrigatório
													CASE WHEN @codigo_parcela is null THEN 
														REPLICATE('0', 2)
													ELSE
														RIGHT(REPLICATE('0', 2)+CONVERT(VARCHAR(2),@codigo_parcela),2) 
													END +  -- 

													-- 8, Data de Emissão do Documento, Obrigatório 
													CASE WHEN @data_emissao is null THEN 
													 	SPACE(8)
													ELSE
														REPLACE(convert(varchar, @data_emissao, 103),'/','')
													END + -- 

													-- 8, Data de Vencimento do Documento, Obrigatório
													CASE WHEN @data_vencimento is null THEN 
														SPACE(8)
													ELSE
														REPLACE(convert(varchar, @data_vencimento, 103),'/','')
													END + -- 

													-- 8, Data Prevista de Pagamento 
													CASE WHEN @data_prevista is null THEN 
														SPACE(8)
													ELSE
														REPLACE(convert(varchar, @data_prevista, 103),'/','')
													END + -- 

													-- 10, Finalidade Econômica Externa 	
													CASE WHEN @finalidade_externa is null THEN 
														SPACE(10)
													ELSE
														RIGHT(SPACE(10)+CONVERT(VARCHAR(10),@finalidade_externa),10) 
													END + 													 

													-- 11, Valor do Documento, Obrigatório
													CASE WHEN @valor_vencimento is null THEN 
													 	REPLICATE('0',11)
													ELSE 
														RIGHT(REPLICATE('0',11) + REPLACE(REPLACE(CAST(@valor_vencimento AS VARCHAR(11)),'.',''),',','')   ,11)
													END +

													-- 8, Data do Desconto 
													CASE WHEN @data_desconto is null THEN 
														SPACE(8)
													ELSE
														REPLACE(convert(varchar, @data_desconto, 103),'/','')
													END + -- 

													-- 11, Valor do Desconto 
													CASE WHEN @valor_desconto is null THEN 
													 	SPACE(11)
													ELSE 
														RIGHT(REPLICATE('0',11) + REPLACE(REPLACE(CAST(@valor_desconto AS VARCHAR(11)),'.',''),',','')   ,11)
													END +

													-- 4, Valor Percentual do Desconto 
													CASE WHEN @valor_percentual is null THEN 
														SPACE(4)
													ELSE
														RIGHT(SPACE(4)+CONVERT(VARCHAR(4),@valor_percentual),4) 
													END + 													 

													-- 2, Número de Dias de Atraso 	
													CASE WHEN @numero_dias_atraso is null THEN 
														REPLICATE('0',2)
													ELSE
														RIGHT(REPLICATE('0',2)+CONVERT(VARCHAR(2),@numero_dias_atraso),2) 
													END + 
													 

													-- 11, Valor Juros por Dia Atraso 
													CASE WHEN @valor_juros_dia_atraso is null THEN 
													 	SPACE(11)
													ELSE 
														RIGHT(REPLICATE('0',11) + REPLACE(REPLACE(CAST(@valor_juros_dia_atraso AS VARCHAR(11)),'.',''),',','')   ,11)
													END +

													-- 4, Percentual de Juros por Dia de Atraso 
													CASE WHEN @percentual_juros_dia_atraso is null THEN 
														SPACE(4)
													ELSE
														RIGHT(SPACE(4)+CONVERT(VARCHAR(4),@percentual_juros_dia_atraso),4) 
													END + 													

													-- 4, Valor Percentual Multa Atraso 
													CASE WHEN @valor_percentual_multa_atraso is null THEN 
														SPACE(4)
													ELSE
														RIGHT(SPACE(4)+CONVERT(VARCHAR(4),@valor_percentual_multa_atraso),4) 
													END + 
													
													 -- 10, Código do Portador Externo 
													CASE WHEN @codigo_portador_externo is null THEN 
														SPACE(10)
													ELSE
														RIGHT(SPACE(@codigo_portador_externo)+CONVERT(VARCHAR(10),@codigo_portador_externo),10) 
													END + 													 

													-- 2, Modalidade Portador Exteno 
													CASE WHEN @modalidade_portador_externo is null THEN 
														SPACE(2)
													ELSE
														RIGHT(SPACE(2)+CONVERT(VARCHAR(2),@modalidade_portador_externo),2) 
													END + 													 

													-- 12, Codigo da Apólice de Seguro 
													CASE WHEN @codigo_apolice_seguro is null THEN 
														SPACE(12)
													ELSE
														RIGHT(SPACE(12)+CONVERT(VARCHAR(12),@codigo_apolice_seguro),12) 
													END + 													 

													-- 12, Código da Seguradora 
													CASE WHEN @codigo_seguradora is null THEN 
														SPACE(12)
													ELSE
														RIGHT(SPACE(12)+CONVERT(VARCHAR(12),@codigo_seguradora),12) 
													END + 													 

													-- 6, Código do Arrendador 
													CASE WHEN @codigo_arrendador is null THEN 
														SPACE(6)
													ELSE
														RIGHT(SPACE(6)+CONVERT(VARCHAR(6),@codigo_arrendador),6) 
													END + 													

													-- 12, Código Contrato de Leasing 	
													CASE WHEN @codigo_contrato is null THEN 
														SPACE(12)
													ELSE
														RIGHT(SPACE(12)+CONVERT(VARCHAR(12),@codigo_contrato),12) 
													END + 													 

													-- 5, Código do Portador 
													CASE WHEN @codigo_portador is null THEN 
														SPACE(5)
													ELSE
														RIGHT(SPACE(5)+CONVERT(VARCHAR(5),@codigo_portador),5) 
													END + 													

													-- 8, Código do Indicador Econômico 
													CASE WHEN @codigo_indicador_economico is null THEN 
														SPACE(8)
													ELSE
														RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@codigo_indicador_economico),8) 
													END + 
													

													-- 3, Código da Forma de Pagamento 
													CASE WHEN @codigo_forma_pagamento is null THEN 
														SPACE(3)
													ELSE
														RIGHT(SPACE(3)+CONVERT(VARCHAR(3),@codigo_forma_pagamento),3) 
													END + 
													 

													-- 200, Descrição do Histórico Padrão
													CASE WHEN @descricao_historico_padrao is null THEN 
														SPACE(200)
													ELSE
														RIGHT(SPACE(200)+CONVERT(VARCHAR(200),@descricao_historico_padrao),200) 
													END + 
													 

													-- 17, Valor Cotação em Moeda Corrente 
													CASE WHEN @valor_cotacao_moeda_corrente is null THEN 
														SPACE(17)
													ELSE
														RIGHT(SPACE(17)+CONVERT(VARCHAR(17),@valor_cotacao_moeda_corrente),17) 
													END + 
													

													-- 8, Indicar Ordem Investimento Integração com o EMS 2
													CASE WHEN @ordem_investimento is null THEN 
														SPACE(8)
													ELSE
														RIGHT(SPACE(8)+CONVERT(VARCHAR(8),@ordem_investimento),8) 
													END + 
													
													-- 5, Série do Documento (*), Obrigatório 
													CASE WHEN @serie is null THEN 
														SPACE(5)
													ELSE
														RIGHT(SPACE(5)+CONVERT(VARCHAR(5),@serie),5) 
													END + 

													-- 16, Código do Documento (*), Obrigatório
													CASE WHEN @documento is null THEN 
														SPACE(16)
													ELSE
														RIGHT(SPACE(16)+CONVERT(VARCHAR(16),@documento),16)													
													END 												 

            FETCH next FROM lancamento_vencimento_cursor INTO @codigo_fornecedor, @especie,@serie, @documento, @codigo_parcela, @data_emissao,
				@data_vencimento, @data_prevista, @finalidade_externa, @valor_vencimento, @data_desconto, @valor_desconto,
				@valor_percentual, @numero_dias_atraso, @valor_juros_dia_atraso, @percentual_juros_dia_atraso, 
				@valor_percentual_multa_atraso,	@codigo_portador_externo, @modalidade_portador_externo, @codigo_apolice_seguro, 
				@codigo_seguradora, @codigo_arrendador, @codigo_contrato, @codigo_portador, @codigo_indicador_economico, 
				@codigo_forma_pagamento, @descricao_historico_padrao, @valor_cotacao_moeda_corrente , @ordem_investimento

        END 

      CLOSE lancamento_vencimento_cursor 

      DEALLOCATE lancamento_vencimento_cursor 

      -- Próximo lançamento
      FETCH next FROM lancamento_cursor INTO @lancamento, @valortotal, @datalancamento
  END 

CLOSE lancamento_cursor; 

DEALLOCATE lancamento_cursor; 
END