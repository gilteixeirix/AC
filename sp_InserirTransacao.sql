-- Procedure para inserir uma nova transação
use GTX -- banco que criei para testes do Desafio Técnico - Desenvolvedor Pleno (VB6/VB.NET + SQL Server)
go

if  exists(select 1 from sysobjects where OBJECT_NAME(id) = 'sp_InserirTransacao')
 drop procedure sp_InserirTransacao;

go

CREATE PROCEDURE sp_InserirTransacao
    @Numero_Cartao CHAR(16),
    @Valor_Transacao DECIMAL(10,2),
    @Data_Transacao DATETIME ,
    @Descricao VARCHAR(255) ,
    @Status_Transacao VARCHAR(10) ,
	@identity int output
AS
BEGIN
    SET NOCOUNT ON;

    -- Verifica se o valor é positivo
    IF @Valor_Transacao <= 0
    BEGIN
        RAISERROR('Valor da transação deve ser positivo.', 16, 1);
        RETURN;
    END

    -- Verifica se o status é válido
    IF @Status_Transacao NOT IN ('Pendente','Aprovada','Cancelada')
    BEGIN
        RAISERROR('Status inválido.', 16, 1);
        RETURN;
    END

    -- Insere o registro
    INSERT INTO Transacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status_Transacao)
    VALUES (
        @Numero_Cartao,
        @Valor_Transacao,
        @Data_Transacao,
        @Descricao,
        @Status_Transacao
    );
	set @identity = @@identity
	return @identity
END
GO
