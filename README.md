# Processador de Tabelas Excel para Documento Word

Este script processa um arquivo Excel contendo múltiplas tabelas empilhadas e as insere automaticamente em um modelo de documento Word.

## Descrição

O script executa as seguintes tarefas:
1. Lê um arquivo Excel contendo múltiplas tabelas empilhadas verticalmente
2. Extrai cada tabela individualmente como imagem
3. Insere cada imagem extraída em campos designados em um modelo de documento Word
4. Processa todas as tabelas encontradas no arquivo Excel sequencialmente

## Requisitos

- Python 3.x
- Bibliotecas necessárias:
    - openpyxl
    - python-docx
    - Pillow

## Uso

1. Prepare seu arquivo Excel com as tabelas empilhadas
2. Configure seu modelo Word com os campos designados
3. Execute o script para processar e gerar o documento final

## Observação

Certifique-se de que suas tabelas Excel estejam formatadas adequadamente e o modelo Word tenha os campos de marcação corretos para as imagens.