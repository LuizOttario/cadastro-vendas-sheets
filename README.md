# Cadastro de vendas com Google Sheets

Projeto da seção 70 do curso de Python + RPA.

Tela em Tkinter ligada a uma planilha do Google com as colunas Estado, Cidade, Vendedor, Produto e Preço. Dá pra cadastrar, editar e excluir linhas direto pela tela, e tudo é salvo na planilha.

## O que tem

- Cadastro de uma venda nova pelo formulário
- Edição: clicando numa linha da tabela os dados vão pro formulário, aí é só mudar e clicar em "Salvar alterações"
- Exclusão da linha selecionada (pede confirmação)
- Total do preço e quantidade de linhas no rodapé, que acompanham o filtro
- Filtro por combobox: escolhe a coluna (Estado, Cidade, Vendedor ou Produto) e depois o valor
- Exportar para Excel o que está na tela, com opção de separar em abas (por exemplo uma aba por estado)

O preço aceita `1500`, `1500,00` ou `1.500,00`.

## Arquivos

- `app.py`: a tela
- `planilha.py`: leitura e escrita no Google Sheets e a exportação pro Excel

## Configurando o acesso ao Google Sheets

1. No [Google Cloud Console](https://console.cloud.google.com/) crie um projeto e ative a **Google Sheets API** e a **Google Drive API**.
2. Em "Credenciais" crie uma **conta de serviço**, gere uma chave do tipo JSON e salve o arquivo nesta pasta como `credenciais.json`.
3. Crie a planilha no Google Sheets com o cabeçalho `Estado | Cidade | Vendedor | Produto | Preço` na primeira linha (se ela estiver vazia o programa coloca o cabeçalho sozinho).
4. Compartilhe a planilha com o e-mail da conta de serviço (está dentro do `credenciais.json`, no campo `client_email`) como editor.
5. Copie o `config.exemplo.json` para `config.json` e coloque o nome da planilha. Se os dados não estiverem na primeira aba, coloque o nome da aba em `aba`.

O `credenciais.json` e o `config.json` estão no `.gitignore` pra não subirem pro GitHub.

## Rodando

```bash
pip install -r requirements.txt
python app.py
```

No Linux pode ser que o Tkinter precise ser instalado à parte (`sudo apt install python3-tk` no Ubuntu, `sudo pacman -S tk` no Arch).
