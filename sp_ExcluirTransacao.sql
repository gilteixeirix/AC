-- Procedure para excluir uma transa��o
use GTX -- banco que criei para testes do Desafio T�cnico - Desenvolvedor Pleno (VB6/VB.NET + SQL Server)
go
if  exists(select 1 from sysobjects where OBJECT_NAME(id) = 'sp_ExcluirTransacao')
 drop procedure sp_ExcluirTransacao;

go

CREATE PROCEDURE sp_ExcluirTransacao
    @Id_Transacao INT
AS
BEGIN
    SET NOCOUNT ON;

    -- Verifica se a transa��o existe
    IF NOT EXISTS (SELECT 1 FROM Transacoes WHERE Id_Transacao = @Id_Transacao)
    BEGIN
        RAISERROR('Transa��o n�o encontrada.', 16, 1);
        RETURN;
	END
	ELSE
	BEGIN
		DELETE FROM Transacoes WHERE Id_Transacao = @Id_Transacao
	END

END

go