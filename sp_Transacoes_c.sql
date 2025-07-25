use GTX -- banco que criei para testes do Desafio Técnico - Desenvolvedor Pleno (VB6/VB.NET + SQL Server)
if not exists(select 1from sysobjects where OBJECT_NAME(id) = 'Transacoes')

CREATE TABLE Transacoes (
    Id_Transacao INT IDENTITY(1,1) PRIMARY KEY,
    Numero_Cartao CHAR(16) NOT NULL,
    Valor_Transacao DECIMAL(10,2) NOT NULL CHECK (Valor_Transacao > 0),
    Data_Transacao DATETIME DEFAULT GETDATE(),
    Descricao VARCHAR(255),
    Status_Transacao SMALLINT NOT NULL CHECK (Status_Transacao IN (1,2,3))
);
