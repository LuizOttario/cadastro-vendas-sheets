"""Tudo que conversa com o Google Sheets e com o Excel fica aqui."""
import gspread
import pandas as pd

COLUNAS = ["Estado", "Cidade", "Vendedor", "Produto", "Preço"]


def converter_preco(texto):
    """Aceita 1234.5, 1234,50 ou 1.234,50 e devolve float."""
    texto = str(texto).strip().replace("R$", "").strip()
    if not texto:
        raise ValueError("Preço vazio")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    return float(texto)


def formatar_moeda(valor):
    texto = f"{valor:,.2f}"
    return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")


class Planilha:
    def __init__(self, arquivo_credenciais, nome_planilha, aba=None):
        cliente = gspread.service_account(filename=arquivo_credenciais)
        planilha = cliente.open(nome_planilha)
        self.aba = planilha.worksheet(aba) if aba else planilha.sheet1
        self._garantir_cabecalho()

    def _garantir_cabecalho(self):
        primeira_linha = self.aba.row_values(1)
        if not primeira_linha:
            self.aba.append_row(COLUNAS)

    def listar(self):
        """Devolve uma lista de (numero_da_linha_na_planilha, dict)."""
        valores = self.aba.get_all_values()
        registros = []
        for numero, linha in enumerate(valores[1:], start=2):
            if not any(linha):
                continue
            linha = (linha + [""] * len(COLUNAS))[: len(COLUNAS)]
            registro = dict(zip(COLUNAS, linha))
            try:
                registro["Preço"] = converter_preco(registro["Preço"])
            except ValueError:
                registro["Preço"] = 0.0
            registros.append((numero, registro))
        return registros

    def adicionar(self, registro):
        self.aba.append_row([registro[c] for c in COLUNAS], value_input_option="RAW")

    def atualizar(self, numero_linha, registro):
        self.aba.update(
            range_name=f"A{numero_linha}:E{numero_linha}",
            values=[[registro[c] for c in COLUNAS]],
            value_input_option="RAW",
        )

    def excluir(self, numero_linha):
        self.aba.delete_rows(numero_linha)


def exportar_excel(registros, caminho, separar_por=None):
    """
    Salva os registros num .xlsx. Se separar_por for o nome de uma coluna,
    cria uma aba para cada valor dela (ex.: uma aba por Estado).
    """
    tabela = pd.DataFrame(registros, columns=COLUNAS)

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        if separar_por:
            for valor, grupo in tabela.groupby(separar_por):
                # nome de aba no excel tem limite de 31 caracteres e não aceita alguns símbolos
                nome_aba = str(valor)
                for c in "[]:*?/\\":
                    nome_aba = nome_aba.replace(c, "-")
                grupo.to_excel(writer, sheet_name=nome_aba[:31] or "vazio", index=False)
        else:
            tabela.to_excel(writer, sheet_name="Vendas", index=False)
