# Holerite_Automatizado

## Processador de Recibos em PDF

Este script em Python processa arquivos PDF de recibos, extrai informações relevantes e envia esses recibos por e-mail para funcionários específicos.

## Funcionalidades

1. **Processamento de PDFs**: Extrai informações de cada PDF, como código do funcionário e mês/ano do recibo.
2. **Criação de PDFs Recortados**: Cria uma nova versão recortada do PDF original.
3. **Envio de E-mails**: Envia o PDF recortado por e-mail para o funcionário correspondente.

## Dependências

- `os`
- `datetime`
- `time`
- `PyPDF2`
- `pandas`
- `re`
- `smtplib`
- `email`
- `config` (Módulo de configuração local)

## Estrutura do Código

### Funções

#### `extract_information(pdf_path)`

Processa um arquivo PDF, extrai informações e envia o PDF recortado por e-mail.

#### `envia_email(destinatarios, subject, text, path_pdf='', path_image='')`

Envia um e-mail com os detalhes fornecidos, anexando o PDF recortado se fornecido.

#### `email_por_codigo_funci(codigo)`

Obtém o e-mail e nome do funcionário pelo código (Função placeholder).

### Bloco Principal

Executa o script principal, processando todos os arquivos PDF no diretório de recibos.

## Configuração

O arquivo `config.py` deve conter as seguintes variáveis:

- `DIRETORIO_RECIBOS`: Diretório onde os PDFs de recibos estão armazenados.
- `DIRETORIO_RECIBOS_PROCESSADOS`: Diretório onde os PDFs processados serão movidos.
- `EMAIL_HOST`: Host do servidor SMTP.
- `EMAIL_USER`: Usuário do e-mail para envio.
- `EMAIL_PASS`: Senha do e-mail para envio.
- `EXECUCAO_TEMPO_ESPERA_SEGUNDOS`: Tempo de espera antes de encerrar a execução (em segundos).

## Como Usar

1. Cria uma pasta com nome da organização. `config.DIRETORIO_RECIBOS`.
2. Dentro da pasta da organização, teremos dois arquivos Python que serão  `main.py`e `config.DIRETORIO_RECIBOS`.
3. Teremos uma pasta normal chamada   `Recibo_pagamento` a qual você recebendo o arquivo de pagamento você coloca dentro da pasta.
4. Ao fazer os passos anteriores, temos um arquivo em Excel a qual tem os dados dos colaboradores a qual o programa irá ler.
5. Por fim, apenas clicar duas vezes no `run.bat`  que ele irá rodar o código e todos os arquivos processados, estarão na pasta processados.
   

