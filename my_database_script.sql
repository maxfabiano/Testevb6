-- Create database if it doesn't exist
CREATE DATABASE IF NOT EXISTS my_database;
-- Change the delimiter to avoid conflict with internal semicolons in the procedure
DELIMITER $$

-- Drop the stored procedure if it exists
DROP PROCEDURE IF EXISTS my_database.sp_TotalTransacoesPorPeriodo$$

-- Create stored procedure
CREATE PROCEDURE my_database.sp_TotalTransacoesPorPeriodo(
    IN Data_Inicial DATE, 
    IN Data_Final DATE
)
BEGIN
    SELECT Numero_Cartao, 
           SUM(Valor_Transacao) AS Valor_Total, 
           COUNT(*) AS Quantidade_Transacoes
    FROM my_database.Transacoes
    WHERE Data_Transacao BETWEEN Data_Inicial AND Data_Final
    GROUP BY Numero_Cartao;
END$$

-- Revert to the default delimiter
DELIMITER ;


-- Use the database
USE my_database;

-- Create Clientes table
CREATE TABLE IF NOT EXISTS my_database.Clientes (
    Numero_Cartao INT AUTO_INCREMENT PRIMARY KEY,
    Nome_Cliente VARCHAR(100) NOT NULL,
    Email VARCHAR(100),
    Telefone VARCHAR(20)
);

-- Create Transacoes table
CREATE TABLE IF NOT EXISTS my_database.Transacoes (
    Id_Transacao INT AUTO_INCREMENT PRIMARY KEY,
    Numero_Cartao INT,
    Valor_Transacao DECIMAL(10,2) NOT NULL,
    Data_Transacao DATE NOT NULL,
    Descricao VARCHAR(255),
    Status INT,
    FOREIGN KEY (Numero_Cartao) REFERENCES my_database.Clientes(Numero_Cartao) ON DELETE CASCADE
);

-- Create Categorias table
CREATE TABLE IF NOT EXISTS my_database.Categorias (
    Id_Categoria INT AUTO_INCREMENT PRIMARY KEY,
    Descricao_Categoria VARCHAR(100) NOT NULL
);

-- Insert default values into Categorias if they don't exist
INSERT INTO my_database.Categorias (Descricao_Categoria)
SELECT 'Alta' FROM DUAL WHERE NOT EXISTS (SELECT 1 FROM my_database.Categorias WHERE Descricao_Categoria = 'Alta')
UNION ALL
SELECT 'Média' FROM DUAL WHERE NOT EXISTS (SELECT 1 FROM my_database.Categorias WHERE Descricao_Categoria = 'Média')
UNION ALL
SELECT 'Baixa' FROM DUAL WHERE NOT EXISTS (SELECT 1 FROM my_database.Categorias WHERE Descricao_Categoria = 'Baixa');

-- Create Transacoes_Categorias table
CREATE TABLE IF NOT EXISTS my_database.Transacoes_Categorias (
    Id INT AUTO_INCREMENT PRIMARY KEY,
    Id_Transacao INT,
    Id_Categoria INT,
    FOREIGN KEY (Id_Transacao) REFERENCES my_database.Transacoes(Id_Transacao) ON DELETE CASCADE,
    FOREIGN KEY (Id_Categoria) REFERENCES my_database.Categorias(Id_Categoria) ON DELETE CASCADE
);

-- Create Auditoria table
CREATE TABLE IF NOT EXISTS my_database.Auditoria (
    Id_Auditoria INT AUTO_INCREMENT PRIMARY KEY,
    Descricao VARCHAR(255),
    Data_Auditoria TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    Id_Transacao INT,
    FOREIGN KEY (Id_Transacao) REFERENCES my_database.Transacoes(Id_Transacao) ON DELETE SET NULL
);

-- Drop view if it exists
DROP VIEW IF EXISTS my_database.vw_TransacoesComCategoria;

-- Create view vw_TransacoesComCategoria
CREATE VIEW my_database.vw_TransacoesComCategoria AS
SELECT c.Nome_Cliente, 
       t.Numero_Cartao, 
       t.Valor_Transacao, 
       t.Data_Transacao, 
       cat.Descricao_Categoria AS Categoria
FROM my_database.Transacoes t
JOIN my_database.Clientes c ON t.Numero_Cartao = c.Numero_Cartao
LEFT JOIN my_database.Transacoes_Categorias tc ON t.Id_Transacao = tc.Id_Transacao
LEFT JOIN my_database.Categorias cat ON tc.Id_Categoria = cat.Id_Categoria;

DELIMITER $$

-- Drop the function if it exists
DROP FUNCTION IF EXISTS my_database.CategorizarTransacao$$

-- Create the function
CREATE FUNCTION my_database.CategorizarTransacao(Valor DECIMAL(10,2))
RETURNS VARCHAR(10)
DETERMINISTIC
BEGIN
    IF Valor > 1000 THEN
        RETURN 'Alta';
    ELSEIF Valor BETWEEN 500 AND 1000 THEN
        RETURN 'Média';
    ELSE
        RETURN 'Baixa';
    END IF;
END$$
