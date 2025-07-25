-- Procedure para alterar uma transação existente
use GTX -- banco que criei para testes do Desafio Técnico - Desenvolvedor Pleno (VB6/VB.NET + SQL Server)
go

if  exists(select 1 from sysobjects where OBJECT_NAME(id) = 'sp_AlterarTransacao')
 drop procedure sp_AlterarTransacao;

go
CREATE PROCEDURE sp_AlterarTransacao
    @Id_Transacao INT,
    @Numero_Cartao CHAR(16) ,
    @Valor_Transacao DECIMAL(10,2) ,
    @Data_Transacao DATETIME ,
    @Descricao VARCHAR(255) ,
    @Status_Transacao VARCHAR(10) 
AS
BEGIN
    SET NOCOUNT ON;

    -- Verifica se a transação existe
    IF NOT EXISTS (SELECT 1 FROM Transacoes WHERE Id_Transacao = @Id_Transacao)
    BEGIN
        RAISERROR('Transação não encontrada.', 16, 1);
        RETURN;
    END

   

    ---- Valida o status 
    --IF  (select Status_Transacao from Transacoes where Id_Transacao = @Id_Transacao) in('Aprovada')
    --BEGIN
    --    RAISERROR('Transações Aprovadas não podem ser Alteradas.', 16, 1);
    --    RETURN;
    --END

    -- Valida o valor 
    IF  @Valor_Transacao <= 0
    BEGIN
        RAISERROR('Valor da transação deve ser positivo.', 16, 1);
        RETURN;
    END

	 -- Atualiza os campos fornecidos
    UPDATE Transacoes
    SET
        Numero_Cartao = COALESCE(@Numero_Cartao, Numero_Cartao),
        Valor_Transacao = CASE WHEN @Valor_Transacao IS NOT NULL THEN @Valor_Transacao ELSE Valor_Transacao END,
        Data_Transacao = COALESCE(@Data_Transacao, Data_Transacao),
        Descricao = COALESCE(@Descricao, Descricao),
        Status_Transacao = @Status_Transacao
    WHERE Id_Transacao = @Id_Transacao;

END
GO
