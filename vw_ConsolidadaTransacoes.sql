-- VIEW para retorna todas as transações categorizadas para um período.
use GTX -- banco que criei para testes do Desafio Técnico - Desenvolvedor Pleno (VB6/VB.NET + SQL Server)
go

if  exists(select 1 from sysobjects where OBJECT_NAME(id) = 'vw_ConsolidadaTransacoes')
 drop VIEW vw_ConsolidadaTransacoes;

go
CREATE VIEW vw_ConsolidadaTransacoes AS
SELECT
		t.Id_Transacao,
		t.Data_Transacao,
		t.Descricao,
		convert(varchar(16),t.Numero_Cartao) Numero_Cartao,
		t.Status_Transacao,
		convert(money,t.Valor_Transacao) Valor_Transacao,
        dbo.fn_CategorizarValor(t.Valor_Transacao) AS Categoria

FROM
    Transacoes t;