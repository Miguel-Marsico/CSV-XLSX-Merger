import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, ttk, simpledialog
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import os
import threading
from tkinter import Menu
import chardet

class Spreadsheetmerger:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV/XLSX Merger")
        self.root.geometry("850x700")
        self.root.resizable(True, True)
        self.root.configure(bg="#f5f7fa")

        self.style = ttk.Style()
        self.configure_styles()

        self.input_folder = tk.StringVar()
        self.output_filename = tk.StringVar()
        self.output_type = tk.StringVar(value="csv")
        self.csv_delimiter = tk.StringVar(value=",")
        self.remove_duplicates = tk.BooleanVar(value=False)
        self.merged_df = None
        self.progress_running = False
        self.duplicates_removed = 0

        self.history = []
        self.history_position = -1
        self.max_history = 15

        self.center_window()
        self.create_widgets()
        self.create_transform_menu()

    def open_github(self):
        webbrowser.open_new("https://github.com/Miguel-Marsico/")

    def center_window(self):
        """Centraliza a janela na tela"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def configure_styles(self):
        """Configura os estilos visuais"""
        self.style.configure("TFrame", background="#f5f7fa")
        self.style.configure("TLabel", background="#f5f7fa", font=("Segoe UI", 10))
        self.style.configure("TButton", font=("Segoe UI", 10), padding=6)
        self.style.configure(
            "TRadiobutton", background="#f5f7fa", font=("Segoe UI", 10)
        )
        self.style.configure(
            "TCheckbutton", background="#f5f7fa", font=("Segoe UI", 10)
        )
        self.style.configure("TNotebook", background="#f5f7fa", borderwidth=0)
        self.style.configure("TNotebook.Tab", padding=[10, 4], font=("Segoe UI", 9))
        self.style.configure(
            "Status.Horizontal.TProgressbar",
            thickness=20,
            troughcolor="#e0e0e0",
            background="#4a6da7",
            troughrelief="flat",
            relief="flat",
        )

        self.style.map(
            "TButton",
            foreground=[("pressed", "white"), ("active", "white")],
            background=[("pressed", "#3a5a8f"), ("active", "#5c7fbf")],
        )

    def create_widgets(self):
        main_container = ttk.Frame(self.root)
        main_container.pack(fill="both", expand=True, padx=20, pady=10)

        center_frame = ttk.Frame(main_container)
        center_frame.pack(expand=True, fill="both")

        self.create_header(center_frame)

        self.create_main_content(center_frame)

        self.create_status_bar(main_container)

    def create_transform_menu(self):
        """Cria o menu de transformação de dados"""
        menubar = Menu(self.root)

        transform_menu = Menu(menubar, tearoff=0)
        transform_menu.add_command(
            label="Renomear Colunas", command=self.rename_columns
        )
        transform_menu.add_command(
            label="Reordenar Colunas", command=self.reorder_columns
        )
        transform_menu.add_command(
            label="Converter Tipos de Dados", command=self.convert_data_types
        )
        transform_menu.add_separator()
        transform_menu.add_command(
            label="Criar Nova Coluna", command=self.create_new_column
        )

    def create_header(self, parent):
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill="x", pady=(0, 15))

        title = ttk.Label(
            header_frame,
            text="📊 CSV/XLSX Merger",
            font=("Segoe UI", 16, "bold"),
            foreground="#2c3e50",
        )
        title.pack(side="left")

        github_label = tk.Label(
            header_frame,
            text="By Miguel Marsico",
            font=("Segoe UI", 9, "underline"),
            fg="blue",
            cursor="hand2",
            bg="#f5f7fa",
        )
        github_label.pack(side="left", padx=(10, 0))
        github_label.bind("<Button-1>", lambda e: self.open_github())

        help_btn = ttk.Button(
            header_frame, text="Ajuda", command=self.show_help, style="TButton"
        )
        help_btn.pack(side="right")

    def create_main_content(self, parent):
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill="both", expand=True)

        config_frame = ttk.Frame(self.notebook)
        self.notebook.add(config_frame, text="Configuração")

        self.create_config_form(config_frame)

        self.preview_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.preview_frame, text="Pré-visualização", state="hidden")

        self.create_preview_area()

    def create_config_form(self, parent):
        form_frame = ttk.Frame(parent)
        form_frame.pack(padx=20, pady=10, fill="x")

        ttk.Label(form_frame, text="Pasta com os arquivos:").grid(
            row=0, column=0, sticky="w", pady=(0, 5)
        )
        folder_entry = ttk.Entry(form_frame, textvariable=self.input_folder, width=60)
        folder_entry.grid(row=1, column=0, sticky="we", pady=(0, 15))
        ttk.Button(form_frame, text="Procurar", command=self.browse_folder).grid(
            row=1, column=1, padx=(10, 0)
        )

        ttk.Label(form_frame, text="Nome do arquivo final:").grid(
            row=2, column=0, sticky="w", pady=(0, 5)
        )
        ttk.Entry(form_frame, textvariable=self.output_filename, width=60).grid(
            row=3, column=0, columnspan=2, sticky="we", pady=(0, 15)
        )

        type_frame = ttk.Frame(form_frame)
        type_frame.grid(row=4, column=0, columnspan=2, sticky="we", pady=(0, 15))
        ttk.Label(type_frame, text="Tipo do arquivo final:").grid(
            row=0, column=0, sticky="w", pady=(0, 5)
        )
        ttk.Radiobutton(
            type_frame,
            text="CSV (.csv)",
            variable=self.output_type,
            value="csv",
            command=self.toggle_csv_options,
        ).grid(row=1, column=0, sticky="w")
        ttk.Radiobutton(
            type_frame,
            text="Excel (.xlsx)",
            variable=self.output_type,
            value="xlsx",
            command=self.toggle_csv_options,
        ).grid(row=1, column=1, sticky="w", padx=(20, 0))

        self.csv_options_frame = ttk.Frame(form_frame)
        self.csv_options_frame.grid(
            row=5, column=0, columnspan=2, sticky="we", pady=(0, 15)
        )

        ttk.Label(self.csv_options_frame, text="Delimitador do arquivo final:").grid(
            row=0, column=0, sticky="w", pady=(0, 5)
        )
        delimiter_combo = ttk.Combobox(
            self.csv_options_frame,
            textvariable=self.csv_delimiter,
            width=5,
            values=[",", ";", "|", "Tab"],
        )
        delimiter_combo.grid(row=0, column=1, sticky="w", padx=(10, 0), pady=(0, 5))
        delimiter_combo.current(0)

        self.toggle_csv_options()

        ttk.Checkbutton(
            form_frame, text="Remover linhas repetidas", variable=self.remove_duplicates
        ).grid(row=6, column=0, sticky="w", pady=(15, 0))

        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=(20, 0), sticky="we")
        self.process_button = ttk.Button(
            button_frame, text="Processar Arquivos", command=self.start_processing
        )
        self.process_button.pack(side="left", padx=(0, 10), ipadx=20)

        terms_label = tk.Label(
            form_frame,
            text="Ao utilizar você concorda com os ",
            font=("Segoe UI", 8),
            bg="#f5f7fa",
            fg="#2c3e50",
        )
        terms_label.grid(row=8, column=0, sticky="w", pady=(15, 0))

        link_label = tk.Label(
            form_frame,
            text="Termos de Uso",
            font=("Segoe UI", 8, "underline"),
            bg="#f5f7fa",
            fg="blue",
            cursor="hand2",
        )
        link_label.grid(row=8, column=0, sticky="e", padx=(0, 120), pady=(15, 0))
        link_label.bind("<Button-1>", lambda e: self.show_terms())

    def toggle_csv_options(self):
        """Mostra ou esconde opções CSV dependendo do tipo de arquivo selecionado"""
        if self.output_type.get() == "csv":
            for widget in self.csv_options_frame.winfo_children():
                widget.grid()
        else:
            for widget in self.csv_options_frame.winfo_children():
                widget.grid_remove()

    def show_terms(self):
        termos = (
            "Termos de Uso\n\n"
            'Este software é fornecido "no estado em que se encontra", sem garantias de qualquer tipo, '
            "expressas ou implícitas, incluindo, mas não se limitando a garantias de comercialização, "
            "adequação a um propósito específico ou não violação.\n\n"
            "Ao utilizar este aplicativo, você concorda que o autor não será responsável por quaisquer danos "
            "diretos, indiretos, incidentais ou consequenciais decorrentes do uso ou da incapacidade de usar este software.\n\n"
            "Você é livre para usar, modificar e distribuir este aplicativo, desde que mantenha a atribuição de autoria."
        )
        messagebox.showinfo("Termos de Uso", termos)

    def create_preview_area(self):
        controls_frame = ttk.Frame(self.preview_frame)
        controls_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(controls_frame, text="Mostrar:", font=("Segoe UI", 10)).pack(
            side="left"
        )

        self.rows_to_show = tk.StringVar(value="10")
        rows_menu = ttk.OptionMenu(
            controls_frame,
            self.rows_to_show,
            "10",
            "10",
            "20",
            "50",
            "100",
            "500",
            command=lambda _: self.update_preview(),
        )
        rows_menu.pack(side="left", padx=5)

        history_frame = ttk.Frame(controls_frame)
        history_frame.pack(side="left", padx=20)

        self.undo_button = ttk.Button(
            history_frame,
            text="↩ Desfazer",
            command=self.undo_transformation,
            state="disabled",
        )
        self.undo_button.pack(side="left", padx=5)

        buttons_frame = ttk.Frame(controls_frame)
        buttons_frame.pack(side="right")

        transform_btn = ttk.Button(
            buttons_frame,
            text="Transformação",
            command=self.show_transform_menu,
            style="TButton",
        )
        transform_btn.pack(side="left", padx=5)

        save_btn = ttk.Button(
            buttons_frame,
            text="Salvar Arquivo",
            command=self.save_file,
            style="TButton",
        )
        save_btn.pack(side="left", padx=5)

        self.preview_text = ScrolledText(
            self.preview_frame,
            wrap=tk.NONE,
            font=("Consolas", 9),
            bg="white",
            padx=10,
            pady=10,
        )
        self.preview_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def show_transform_menu(self):
        """Mostra o menu de transformação como um menu popup"""
        menu = tk.Menu(self.root, tearoff=0)

        org_menu = tk.Menu(menu, tearoff=0)
        org_menu.add_command(label="Renomear Colunas", command=self.rename_columns)
        org_menu.add_command(label="Reordenar Colunas", command=self.reorder_columns)
        menu.add_cascade(label="Organização", menu=org_menu)

        conv_menu = tk.Menu(menu, tearoff=0)
        conv_menu.add_command(
            label="Converter Tipos de Dados", command=self.convert_data_types
        )
        menu.add_cascade(label="Conversão", menu=conv_menu)

        calc_menu = tk.Menu(menu, tearoff=0)
        calc_menu.add_command(label="Criar Nova Coluna", command=self.create_new_column)
        menu.add_cascade(label="Adicionar", menu=calc_menu)

        try:
            menu.tk_popup(self.root.winfo_pointerx(), self.root.winfo_pointery())
        finally:
            menu.grab_release()

    def create_status_bar(self, parent):
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill="x", pady=(5, 0))

        self.progress_bar = ttk.Progressbar(
            status_frame,
            orient="horizontal",
            mode="determinate",
            style="Status.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill="x")

        self.status_text = ttk.Label(
            status_frame,
            text="Pronto",
            anchor="w",
            font=("Segoe UI", 9),
            foreground="#2c3e50",
        )
        self.status_text.pack(fill="x")

    def save_state(self):
        """Salva o estado atual do DataFrame no histórico"""
        if self.merged_df is None:
            return

        if len(self.history) > self.max_history:
            self.history = self.history[-(self.max_history) :]
            self.history_position = len(self.history) - 1

        if self.history_position < len(self.history) - 1:
            self.history = self.history[: self.history_position + 1]

        self.history.append(self.merged_df.copy())
        self.history_position = len(self.history) - 1

        self.update_history_buttons()

    def undo_transformation(self):
        """Desfaz a última transformação"""
        if self.history_position > 0:
            self.history_position -= 1
            self.merged_df = self.history[self.history_position].copy()
            self.update_preview()
            self.update_history_buttons()
            self.update_progress(100, "Operação desfeita", "green")

    def update_history_buttons(self):
        """Atualiza o estado dos botões de histórico"""
        if self.history_position > 0:
            self.undo_button.config(state="normal")
        else:
            self.undo_button.config(state="disabled")

    def rename_columns(self):
        if self.merged_df is None:
            messagebox.showwarning(
                "Aviso", "Nenhum dado disponível para transformação."
            )
            return

        self.save_state()

        rename_window = tk.Toplevel(self.root)
        rename_window.title("Renomear Colunas")
        rename_window.geometry("500x400")

        columns_frame = ttk.Frame(rename_window)
        columns_frame.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(
            columns_frame,
            text="Colunas Atuais → Novos Nomes",
            font=("Segoe UI", 10, "bold"),
        ).pack()

        canvas = tk.Canvas(columns_frame)
        scrollbar = ttk.Scrollbar(
            columns_frame, orient="vertical", command=canvas.yview
        )
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.rename_entries = {}
        for idx, col in enumerate(self.merged_df.columns):
            ttk.Label(scrollable_frame, text=col).grid(
                row=idx, column=0, padx=5, pady=2, sticky="e"
            )
            entry = ttk.Entry(scrollable_frame)
            entry.insert(0, col)
            entry.grid(row=idx, column=1, padx=5, pady=2, sticky="we")
            self.rename_entries[col] = entry

        button_frame = ttk.Frame(rename_window)
        button_frame.pack(fill="x", padx=10, pady=10)

        ttk.Button(button_frame, text="Cancelar", command=rename_window.destroy).pack(
            side="right", padx=5
        )
        ttk.Button(button_frame, text="Aplicar", command=self.apply_renaming).pack(
            side="right", padx=5
        )

    def apply_renaming(self):
        new_names = {}
        for old_name, entry in self.rename_entries.items():
            new_name = entry.get().strip()
            if new_name:
                new_names[old_name] = new_name

        if new_names:
            self.merged_df.rename(columns=new_names, inplace=True)
            messagebox.showinfo("Sucesso", "Colunas renomeadas com sucesso!")
            self.update_preview()

        for window in self.root.winfo_children():
            if isinstance(window, tk.Toplevel) and window.title() == "Renomear Colunas":
                window.destroy()
                break

    def reorder_columns(self):
        if self.merged_df is None:
            messagebox.showwarning(
                "Aviso", "Nenhum dado disponível para transformação."
            )
            return

        self.save_state()

        reorder_window = tk.Toplevel(self.root)
        reorder_window.title("Reordenar Colunas")
        reorder_window.geometry("500x500")

        ttk.Label(reorder_window, text="Reordenar Colunas", font=("Segoe UI", 10)).pack(
            pady=5
        )

        listbox = tk.Listbox(reorder_window, selectmode=tk.SINGLE, height=15)
        for col in self.merged_df.columns:
            listbox.insert(tk.END, col)
        listbox.pack(fill="both", expand=True, padx=10, pady=5)

        button_frame = ttk.Frame(reorder_window)
        button_frame.pack(fill="x", padx=10, pady=10)

        ttk.Button(
            button_frame,
            text="Mover para Cima",
            command=lambda: self.move_item(listbox, -1),
        ).pack(side="left", padx=5)
        ttk.Button(
            button_frame,
            text="Mover para Baixo",
            command=lambda: self.move_item(listbox, 1),
        ).pack(side="left", padx=5)

        ttk.Button(button_frame, text="Cancelar", command=reorder_window.destroy).pack(
            side="right", padx=5
        )
        ttk.Button(
            button_frame,
            text="Aplicar",
            command=lambda: self.apply_reordering(listbox, reorder_window),
        ).pack(side="right", padx=5)

    def move_item(self, listbox, direction):
        selected = listbox.curselection()
        if not selected:
            return

        pos = selected[0]
        if (direction < 0 and pos == 0) or (
            direction > 0 and pos == listbox.size() - 1
        ):
            return

        new_pos = pos + direction
        text = listbox.get(pos)
        listbox.delete(pos)
        listbox.insert(new_pos, text)
        listbox.select_set(new_pos)

    def apply_reordering(self, listbox, window):
        new_order = [listbox.get(i) for i in range(listbox.size())]
        self.merged_df = self.merged_df[new_order]
        messagebox.showinfo("Sucesso", "Colunas reordenadas com sucesso!")
        self.update_preview()
        window.destroy()

    def convert_data_types(self):
        if self.merged_df is None:
            messagebox.showwarning(
                "Aviso", "Nenhum dado disponível para transformação."
            )
            return

        self.save_state()

        convert_window = tk.Toplevel(self.root)
        convert_window.title("Converter Tipos de Dados")
        convert_window.geometry("500x400")

        columns_frame = ttk.Frame(convert_window)
        columns_frame.pack(fill="both", expand=True, padx=10, pady=10)

        ttk.Label(
            columns_frame,
            text="Coluna → Tipo Atual → Novo Tipo",
            font=("Segoe UI", 10, "bold"),
        ).pack()

        canvas = tk.Canvas(columns_frame)
        scrollbar = ttk.Scrollbar(
            columns_frame, orient="vertical", command=canvas.yview
        )
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.type_vars = {}
        type_options = ["Manter", "Texto", "Inteiro", "Decimal", "Data"]

        for idx, col in enumerate(self.merged_df.columns):
            current_type = str(self.merged_df[col].dtype)
            ttk.Label(scrollable_frame, text=col).grid(
                row=idx, column=0, padx=5, pady=2, sticky="e"
            )
            ttk.Label(scrollable_frame, text=current_type).grid(
                row=idx, column=1, padx=5, pady=2
            )

            var = tk.StringVar(value="Manter")
            option = ttk.OptionMenu(scrollable_frame, var, *type_options)
            option.grid(row=idx, column=2, padx=5, pady=2, sticky="we")
            self.type_vars[col] = var

        button_frame = ttk.Frame(convert_window)
        button_frame.pack(fill="x", padx=10, pady=10)

        ttk.Button(button_frame, text="Cancelar", command=convert_window.destroy).pack(
            side="right", padx=5
        )
        ttk.Button(
            button_frame, text="Aplicar", command=self.apply_type_conversion
        ).pack(side="right", padx=5)

    def apply_type_conversion(self):
        for col, var in self.type_vars.items():
            new_type = var.get()

            try:
                if new_type == "Texto":
                    self.merged_df[col] = self.merged_df[col].astype(str)
                elif new_type == "Inteiro":
                    self.merged_df[col] = pd.to_numeric(
                        self.merged_df[col], errors="coerce"
                    ).astype("Int64")
                elif new_type == "Decimal":
                    self.merged_df[col] = pd.to_numeric(
                        self.merged_df[col], errors="coerce"
                    )
                elif new_type == "Data":
                    self.merged_df[col] = pd.to_datetime(
                        self.merged_df[col], errors="coerce"
                    )
            except Exception as e:
                messagebox.showerror(
                    "Erro", f"Erro ao converter coluna {col}: {str(e)}"
                )
                continue

        messagebox.showinfo("Sucesso", "Conversão de tipos aplicada!")
        self.update_preview()

        for window in self.root.winfo_children():
            if (
                isinstance(window, tk.Toplevel)
                and window.title() == "Converter Tipos de Dados"
            ):
                window.destroy()
                break

    def create_new_column(self):
        if self.merged_df is None:
            messagebox.showwarning(
                "Aviso", "Nenhum dado disponível para transformação."
            )
            return

        self.save_state()

        col_name = simpledialog.askstring("Nova Coluna", "Nome da nova coluna:")
        if not col_name:
            return

        value = simpledialog.askstring(
            "Nova Coluna", "Valor ou fórmula (use $ para referenciar colunas):"
        )
        if value is None:
            return

        try:
            if value.startswith("=") or "$" in value:
                for col in self.merged_df.columns:
                    if f"${col}" in value:
                        value = value.replace(f"${col}", f"self.merged_df['{col}']")

                try:
                    self.merged_df[col_name] = eval(value[1:])
                except:
                    self.merged_df[col_name] = value
            else:
                self.merged_df[col_name] = value

            messagebox.showinfo("Sucesso", f"Coluna '{col_name}' criada com sucesso!")
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao criar coluna: {str(e)}")

    def start_processing(self):
        """Inicia o processamento em uma thread separada"""
        if not self.progress_running:
            self.progress_running = True
            self.process_button.config(state="disabled")
            threading.Thread(target=self.process_files, daemon=True).start()

    def update_progress(self, value, message, color="#2c3e50"):
        """Atualiza a barra de progresso e mensagem de status"""
        self.progress_bar["value"] = value
        self.status_text.config(text=message, foreground=color)
        self.root.update_idletasks()

    def show_help(self):
        help_text = """
    📌 Como usar a CSV/XLSX Merger:

    1. COLETAR ARQUIVOS:
       - Coloque todos os arquivos (CSV ou XLSX) que deseja unificar em uma pasta
       - Todos os arquivos devem ter a mesma estrutura de colunas

    2. CONFIGURAR:
       - Clique em 'Procurar' e selecione a pasta com os arquivos
       - Digite um nome para o arquivo final
       - Escolha o formato de saída (CSV ou XLSX)

       OPÇÕES PARA CSV:
       - Selecione o delimitador (vírgula, ponto-e-vírgula, etc.)
       - O programa detecta automaticamente a codificação dos arquivos

    3. PROCESSAR:
       - Marque 'Remover linhas repetidas' se necessário
       - Clique em 'Processar Arquivos'
       - Verifique a pré-visualização na aba correspondente

    4. TRANSFORMAR DADOS (opcional):
       - Use o menu 'Transformação' para:
         * Renomear/Reordenar colunas
         * Converter tipos de dados
         * Criar novas colunas
       - Use o botão 'Desfazer' para reverter transformações indesejadas

    5. SALVAR:
       - Clique em 'Salvar Arquivo' quando estiver satisfeito
       - Para CSV: o delimitador selecionado será aplicado

    🔍 DICAS AVANÇADAS:
    - O programa tenta automaticamente estas codificações de arquivo:
      UTF-8, Latin1, ISO-8859-1, CP1252, UTF-16
    - Use 'Tab' como delimitador para arquivos TSV
    - As transformações podem ser desfeitas em qualquer momento
    - Você pode visualizar diferentes quantidades de linhas na pré-visualização

    ⚠️ LIMITAÇÕES:
    - Para arquivos muito grandes, algumas operações podem demorar
    """
        help_window = tk.Toplevel(self.root)
        help_window.title("Como usar")
        help_window.iconbitmap('Icone.ico')
        help_window.geometry("700x500")

        text_frame = ttk.Frame(help_window)
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)

        text = ScrolledText(text_frame, wrap=tk.WORD, font=("Segoe UI", 10))
        text.pack(fill="both", expand=True)
        text.insert(tk.END, help_text.strip())
        text.config(state="disabled")

        close_btn = ttk.Button(help_window, text="Fechar", command=help_window.destroy)
        close_btn.pack(pady=10)

    def process_files(self):
        folder = self.input_folder.get()

        if not os.path.isdir(folder):
            messagebox.showerror("Erro", "Pasta inválida.")
            self.update_progress(0, "Erro: Pasta inválida", "red")
            self.progress_running = False
            self.process_button.config(state="normal")
            return

        files = []
        for f in os.listdir(folder):
            if f.lower().endswith(('.csv', '.xlsx')):
                files.append(f)

        if not files:
            messagebox.showerror("Erro", "Nenhum arquivo CSV ou XLSX encontrado na pasta.")
            self.update_progress(0, "Erro: Nenhum arquivo encontrado", "red")
            self.progress_running = False
            self.process_button.config(state="normal")
            return

        self.update_progress(10, "Iniciando processamento...")

        dfs = []
        header = None
        total_files = len(files)

        for idx, file in enumerate(files):
            if not self.progress_running:
                break

            path = os.path.join(folder, file)
            try:
                progress = 10 + (idx / total_files) * 80
                self.update_progress(progress, f"Processando {idx + 1}/{total_files}: {file[:20]}...")

                if file.lower().endswith('.csv'):
                    df = self.read_csv_with_encodings(path)
                else:
                    df = pd.read_excel(path, engine="openpyxl")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao ler {file}: {str(e)}")
                self.update_progress(0, f"Erro ao processar {file}", "red")
                self.progress_running = False
                self.process_button.config(state="normal")
                return

            if idx == 0:
                header = df.columns.tolist()
            else:
                if df.columns.tolist() != header:
                    messagebox.showerror("Erro", f"Cabeçalho incompatível no arquivo: {file}")
                    self.update_progress(0, f"Erro: Cabeçalho incompatível em {file}", "red")
                    self.progress_running = False
                    self.process_button.config(state="normal")
                    return
            dfs.append(df)

        self.merged_df = pd.concat(dfs, ignore_index=True)

        self.history = [self.merged_df.copy()]
        self.history_position = 0
        self.update_history_buttons()

        self.duplicates_removed = 0
        if self.remove_duplicates.get():
            initial_count = len(self.merged_df)
            self.merged_df.drop_duplicates(inplace=True)
            self.duplicates_removed = initial_count - len(self.merged_df)

        self.update_progress(95, "Finalizando...")

        messagebox.showinfo("Sucesso", f"{len(files)} arquivos combinados com sucesso.")
        self.update_progress(100, f"Pronto - {len(files)} arquivos processados", "green")

        self.notebook.tab(1, state="normal")
        self.notebook.select(1)
        self.update_preview()

        self.progress_running = False
        self.process_button.config(state="normal")

    def read_csv_with_encodings(self, file_path, encodings=None):
        """Tenta ler um CSV com múltiplas codificações automaticamente."""
        if encodings is None:
            encodings = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252', 'utf-16']

        with open(file_path, 'rb') as f:
            rawdata = f.read(10000)
            result = chardet.detect(rawdata)
            detected_encoding = result['encoding']

            if detected_encoding:
                try:
                    return pd.read_csv(file_path, encoding=detected_encoding)
                except:
                    pass

        for encoding in encodings:
            try:
                return pd.read_csv(file_path, encoding=encoding)
            except UnicodeDecodeError:
                continue

        raise ValueError(f"Não foi possível ler o arquivo com as codificações tentadas: {encodings}")

    def save_file(self):
        if self.merged_df is None:
            messagebox.showerror("Erro", "Nenhum dado disponível para salvar.")
            return

        filename = self.output_filename.get().strip()
        if not filename:
            messagebox.showerror("Erro", "Nome do arquivo final não pode estar vazio.")
            self.update_progress(0, "Erro: Nome do arquivo vazio", "red")
            return

        file_ext = self.output_type.get()
        filetypes = [("CSV Files", "*.csv")] if file_ext == "csv" else [("Excel Files", "*.xlsx")]
        save_path = filedialog.asksaveasfilename(defaultextension=f".{file_ext}",
                                                 filetypes=filetypes,
                                                 initialfile=filename)

        if save_path:
            try:
                self.update_progress(0, "Salvando arquivo...")

                if file_ext == "csv":
                    delimiter = "\t" if self.csv_delimiter.get() == "Tab" else self.csv_delimiter.get()
                    self.merged_df.to_csv(save_path, index=False, sep=delimiter)
                else:
                    self.merged_df.to_excel(save_path, index=False, engine="openpyxl")

                messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso.")
                self.update_progress(100, f"Arquivo salvo em: {save_path}", "green")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {str(e)}")
                self.update_progress(0, "Erro ao salvar arquivo", "red")

    def update_preview(self, event=None):
        if self.merged_df is None:
            return

        try:
            n_rows = int(self.rows_to_show.get())
        except ValueError:
            n_rows = 10

        self.preview_text.delete(1.0, tk.END)
        sample_df = self.merged_df.head(n_rows)
        df_string = sample_df.to_string(index=False)

        self.preview_text.insert(tk.END, df_string)

        info_text = f"\n\n► Total de linhas: {len(self.merged_df)}"
        info_text += f"\n► Total de colunas: {len(self.merged_df.columns)}"
        info_text += f"\n► Colunas: {', '.join(self.merged_df.columns)}"

        if self.remove_duplicates.get() and self.duplicates_removed > 0:
            info_text += f"\n► Linhas repetidas removidas: {self.duplicates_removed}"

        self.preview_text.insert(tk.END, info_text)

    def browse_folder(self):
        """Abre o diálogo para selecionar uma pasta"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.input_folder.set(folder_selected)
            self.update_progress(0, f"Pasta selecionada: {folder_selected}")

if __name__ == "__main__":
    root = tk.Tk()
    try:
        root.iconbitmap('Icone.ico')
    except:
        pass
    app = Spreadsheetmerger(root)
    root.mainloop()
