-- Procedure para CalcularTransacoes Por Periodo
use GTX -- banco que criei para testes do Desafio Técnico - Desenvolvedor Pleno (VB6/VB.NET + SQL Server)
go

if  exists(select 1 from sysobjects where OBJECT_NAME(id) = 'fn_CategorizarValor')
 drop function fn_CategorizarValor;

go
CREATE FUNCTION dbo.fn_CategorizarValor(@Valor DECIMAL)
RETURNS VARCHAR(10)
AS
BEGIN
    DECLARE @Categoria VARCHAR(10);

    IF @Valor > 2000
        SET @Categoria = 'Premium';
    ELSE IF @Valor >= 1000 AND @Valor <= 2000
        SET @Categoria = 'Alta';
    ELSE IF @Valor >= 500 AND @Valor < 1000
        SET @Categoria = 'Média';
    ELSE
        SET @Categoria = 'Baixa';

    RETURN @Categoria;
END;