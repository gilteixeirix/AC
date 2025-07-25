-- Procedure para CalcularTransacoes Por Periodo
use GTX -- banco que criei para testes do Desafio Técnico - Desenvolvedor Pleno (VB6/VB.NET + SQL Server)
go

if  exists(select 1 from sysobjects where OBJECT_NAME(id) = 'sp_CalcularTransacoesPorPeriodo')
 drop procedure sp_CalcularTransacoesPorPeriodo;

go
CREATE PROCEDURE sp_CalcularTransacoesPorPeriodo
    @Data_Inicial DATETIME,
    @Data_Final DATETIME,
    @Status_Transacao VARCHAR(10)
AS
BEGIN
    SELECT 
        Numero_Cartao,
        SUM(Valor_Transacao) AS Valor_Total,
        COUNT(*) AS Quantidade_Transacoes,
        Status_Transacao
    FROM 
        Transacoes
    WHERE 
        Data_Transacao BETWEEN @Data_Inicial AND @Data_Final
        AND Status_Transacao = @Status_Transacao
    GROUP BY 
        Numero_Cartao,
		Status_Transacao;
END
