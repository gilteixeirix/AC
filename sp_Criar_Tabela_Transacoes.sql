use GTX -- banco que criei para testes do Desafio Técnico - Desenvolvedor Pleno (VB6/VB.NET + SQL Server)
go

if not exists(select 1  from sysobjects where OBJECT_NAME(id) = 'Transacoes')

CREATE TABLE Transacoes (
    Id_Transacao INT IDENTITY(1,1) PRIMARY KEY,
    Numero_Cartao NCHAR(16) NOT NULL,
    Valor_Transacao DECIMAL(13,2) NOT NULL CHECK (Valor_Transacao > 0),
    Data_Transacao DATETIME  NOT NULL DEFAULT GETDATE(),
    Descricao VARCHAR(255)  NOT NULL,
    Status_Transacao VARCHAR(10) NOT NULL
);

go