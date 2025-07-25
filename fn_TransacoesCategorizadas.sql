-- Procedure para retorna todas as transações categorizadas para um período.
use GTX -- banco que criei para testes do Desafio Técnico - Desenvolvedor Pleno (VB6/VB.NET + SQL Server)
go

if  exists(select 1 from sysobjects where OBJECT_NAME(id) = 'fn_TransacoesCategorizadas')
 drop function fn_TransacoesCategorizadas;

go
CREATE FUNCTION dbo.fn_TransacoesCategorizadas(@DataInicio DATE, @DataFim DATE)
RETURNS TABLE
AS
RETURN
(
    SELECT
        t.Id_Transacao,
		t.Data_Transacao,
		t.Descricao,
		t.Numero_Cartao,
		t.Status_Transacao,
		t.Valor_Transacao,
        dbo.fn_CategorizarValor(t.Valor_Transacao) AS Categoria
    FROM
        Transacoes t
    WHERE
        t.Data_Transacao BETWEEN @DataInicio AND @DataFim
);