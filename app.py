import json
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from planilha import COLUNAS, Planilha, converter_preco, exportar_excel, formatar_moeda

PASTA = Path(__file__).parent
TODOS = "(todos)"


class App:
    def __init__(self, root, planilha):
        self.root = root
        self.planilha = planilha
        self.registros = []          # tudo que veio da planilha: (linha, dict)
        self.linha_selecionada = None

        root.title("Cadastro de vendas - Google Sheets")
        root.geometry("900x600")
        root.minsize(800, 500)

        self.montar_formulario()
        self.montar_filtro()
        self.montar_tabela()
        self.montar_rodape()

        self.recarregar()

    # ---------- montagem da tela ----------

    def montar_formulario(self):
        quadro = ttk.LabelFrame(self.root, text="Dados da venda")
        quadro.pack(fill="x", padx=10, pady=(10, 5))

        self.campos = {}
        for i, coluna in enumerate(COLUNAS):
            ttk.Label(quadro, text=coluna).grid(row=0, column=i, sticky="w", padx=5, pady=(5, 0))
            entrada = ttk.Entry(quadro, width=18)
            entrada.grid(row=1, column=i, padx=5, pady=(0, 5), sticky="ew")
            self.campos[coluna] = entrada
            quadro.columnconfigure(i, weight=1)

        botoes = ttk.Frame(quadro)
        botoes.grid(row=2, column=0, columnspan=len(COLUNAS), sticky="w", padx=5, pady=5)
        ttk.Button(botoes, text="Cadastrar", command=self.cadastrar).pack(side="left", padx=(0, 5))
        self.btn_salvar = ttk.Button(botoes, text="Salvar alterações", command=self.salvar, state="disabled")
        self.btn_salvar.pack(side="left", padx=5)
        self.btn_excluir = ttk.Button(botoes, text="Excluir", command=self.excluir, state="disabled")
        self.btn_excluir.pack(side="left", padx=5)
        ttk.Button(botoes, text="Limpar", command=self.limpar_formulario).pack(side="left", padx=5)

    def montar_filtro(self):
        quadro = ttk.Frame(self.root)
        quadro.pack(fill="x", padx=10, pady=5)

        ttk.Label(quadro, text="Filtrar por").pack(side="left")
        self.cb_coluna = ttk.Combobox(quadro, values=COLUNAS[:-1], state="readonly", width=12)
        self.cb_coluna.set("Estado")
        self.cb_coluna.pack(side="left", padx=5)
        self.cb_coluna.bind("<<ComboboxSelected>>", lambda e: self.atualizar_valores_filtro())

        self.cb_valor = ttk.Combobox(quadro, state="readonly", width=25)
        self.cb_valor.pack(side="left", padx=5)
        self.cb_valor.bind("<<ComboboxSelected>>", lambda e: self.mostrar())

        ttk.Button(quadro, text="Atualizar", command=self.recarregar).pack(side="right")

    def montar_tabela(self):
        quadro = ttk.Frame(self.root)
        quadro.pack(fill="both", expand=True, padx=10, pady=5)

        self.tabela = ttk.Treeview(quadro, columns=COLUNAS, show="headings", selectmode="browse")
        for coluna in COLUNAS:
            self.tabela.heading(coluna, text=coluna)
            self.tabela.column(coluna, width=150, anchor="e" if coluna == "Preço" else "w")

        barra = ttk.Scrollbar(quadro, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=barra.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")

        self.tabela.bind("<<TreeviewSelect>>", self.ao_selecionar)

    def montar_rodape(self):
        quadro = ttk.Frame(self.root)
        quadro.pack(fill="x", padx=10, pady=(5, 10))

        self.lbl_qtd = ttk.Label(quadro, text="Linhas: 0")
        self.lbl_qtd.pack(side="left")
        self.lbl_total = ttk.Label(quadro, text="Total: R$ 0,00", font=("TkDefaultFont", 10, "bold"))
        self.lbl_total.pack(side="left", padx=20)

        ttk.Button(quadro, text="Exportar para Excel", command=self.exportar).pack(side="right")
        self.cb_abas = ttk.Combobox(quadro, values=COLUNAS[:-1], state="readonly", width=12)
        self.cb_abas.set("Estado")
        self.cb_abas.pack(side="right", padx=5)
        self.var_separar = tk.BooleanVar(value=False)
        ttk.Checkbutton(quadro, text="Separar em abas por", variable=self.var_separar).pack(side="right")

    # ---------- dados ----------

    def recarregar(self):
        try:
            self.registros = self.planilha.listar()
        except Exception as erro:
            messagebox.showerror("Erro", f"Não consegui ler a planilha:\n{erro}")
            return
        self.atualizar_valores_filtro(manter=True)
        self.limpar_formulario()

    def atualizar_valores_filtro(self, manter=False):
        coluna = self.cb_coluna.get()
        valores = sorted({r[coluna] for _, r in self.registros if r[coluna]})
        atual = self.cb_valor.get()
        self.cb_valor["values"] = [TODOS] + valores
        if not (manter and atual in valores):
            self.cb_valor.set(TODOS)
        self.mostrar()

    def registros_filtrados(self):
        coluna, valor = self.cb_coluna.get(), self.cb_valor.get()
        if valor in ("", TODOS):
            return self.registros
        return [(n, r) for n, r in self.registros if r[coluna] == valor]

    def mostrar(self):
        self.tabela.delete(*self.tabela.get_children())
        filtrados = self.registros_filtrados()
        total = 0.0
        for numero, r in filtrados:
            valores = [r[c] for c in COLUNAS[:-1]] + [formatar_moeda(r["Preço"])]
            # o iid do item é o número da linha na planilha, assim sei o que editar/excluir
            self.tabela.insert("", "end", iid=str(numero), values=valores)
            total += r["Preço"]

        self.lbl_qtd.config(text=f"Linhas: {len(filtrados)}")
        self.lbl_total.config(text=f"Total: {formatar_moeda(total)}")

    def ler_formulario(self):
        registro = {c: self.campos[c].get().strip() for c in COLUNAS}
        vazios = [c for c, v in registro.items() if not v]
        if vazios:
            messagebox.showwarning("Atenção", "Preencha: " + ", ".join(vazios))
            return None
        try:
            registro["Preço"] = converter_preco(registro["Preço"])
        except ValueError:
            messagebox.showwarning("Atenção", "Preço inválido. Use algo como 1500 ou 1.500,00")
            return None
        return registro

    def limpar_formulario(self):
        for entrada in self.campos.values():
            entrada.delete(0, "end")
        self.linha_selecionada = None
        self.btn_salvar.config(state="disabled")
        self.btn_excluir.config(state="disabled")
        if self.tabela.selection():
            self.tabela.selection_remove(self.tabela.selection())

    def ao_selecionar(self, _evento=None):
        selecao = self.tabela.selection()
        if not selecao:
            return
        numero = int(selecao[0])
        registro = dict(next(r for n, r in self.registros if n == numero))
        registro["Preço"] = f"{registro['Preço']:.2f}".replace(".", ",")

        for c in COLUNAS:
            self.campos[c].delete(0, "end")
            self.campos[c].insert(0, registro[c])

        self.linha_selecionada = numero
        self.btn_salvar.config(state="normal")
        self.btn_excluir.config(state="normal")

    # ---------- ações ----------

    def cadastrar(self):
        registro = self.ler_formulario()
        if registro is None:
            return
        self.planilha.adicionar(registro)
        self.recarregar()

    def salvar(self):
        if self.linha_selecionada is None:
            return
        registro = self.ler_formulario()
        if registro is None:
            return
        self.planilha.atualizar(self.linha_selecionada, registro)
        self.recarregar()

    def excluir(self):
        if self.linha_selecionada is None:
            return
        if not messagebox.askyesno("Excluir", "Excluir a linha selecionada da planilha?"):
            return
        self.planilha.excluir(self.linha_selecionada)
        self.recarregar()

    def exportar(self):
        caminho = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="vendas.xlsx",
        )
        if not caminho:
            return
        separar_por = self.cb_abas.get() if self.var_separar.get() else None
        registros = [r for _, r in self.registros_filtrados()]
        exportar_excel(registros, caminho, separar_por)
        messagebox.showinfo("Pronto", f"Arquivo salvo em:\n{caminho}")


def main():
    arquivo_config = PASTA / "config.json"
    if not arquivo_config.exists():
        raise SystemExit("Crie o config.json a partir do config.exemplo.json (veja o README)")
    config = json.loads(arquivo_config.read_text(encoding="utf-8"))

    planilha = Planilha(
        PASTA / config["credenciais"],
        config["planilha"],
        config.get("aba"),
    )

    root = tk.Tk()
    App(root, planilha)
    root.mainloop()


if __name__ == "__main__":
    main()
