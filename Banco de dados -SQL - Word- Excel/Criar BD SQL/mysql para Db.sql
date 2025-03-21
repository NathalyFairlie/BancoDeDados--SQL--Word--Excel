CREATE DATABASE cadastrodb;
USE cadastrodb;
CREATE TABLE pessoas (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nome VARCHAR(100) NOT NULL,
    idade INT NOT NULL CHECK (idade BETWEEN 1 AND 99),
    profissao VARCHAR(100) NOT NULL
);
