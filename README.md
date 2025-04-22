<h1>
    📊 CSV-XLSX-Merger
</h1>

## 📋 Tópicos

<div>
 • <a href="#-sobre">Sobre</a> </br>
 • <a href="#-funcionalidades">Funcionalidades</a> </br>
 • <a href="#-ferramentas">Ferramentas</a> </br>
 • <a href="#-como-executar-o-projeto">Como executar o projeto</a> </br>   
 • <a href="#-licença">Licença</a></br>
</div>

## 📗 Sobre

**Quem nunca precisou juntar planilhas espalhadas e acabou se perdendo no Ctrl+C e Ctrl+V?**  
O CSV/XLSX Merger resolve isso com poucos cliques: une, limpa, transforma e exporta seus dados de forma rápida e sem dor de cabeça.

**Veja funcionando na prática no [Youtube](https://pandas.pydata.org/)**

## ✨ Funcionalidades

- Unifica diversos arquivos `.csv` e `.xlsx` automaticamente.
- Detecta automaticamente codificação dos arquivos (UTF-8, Latin1, CP1252, etc.).
- Identifica o padrão de colunas automaticamente.
- Remove linhas duplicadas com um clique.
- Renomeia e reordena colunas dinamicamente.
- Converte tipos de dados (texto, número, data).
- Exporta para Excel ou CSV com delimitador personalizado (`,` `;` `|` `Tab`).
- Possui histórico com suporte a **desfazer alterações**.
- Pré-visualização interativa com estatísticas.
    
<div style="display: flex; gap: 10px;">
  <img src="https://github.com/user-attachments/assets/469885e1-a76d-420a-b257-23167954fe65" alt="Tela de configuração" width="48%" />
  <img src="https://github.com/user-attachments/assets/8a36a07c-e857-4139-b1ed-c2b00e33102a" alt="Pré-visualização" width="48%" />
</div>

## 🔧 Ferramentas

### 🐍 **Aplicação (Python + Tkinter)**

- [Pandas](https://pandas.pydata.org/) – manipulação de dados
- [Tkinter](https://docs.python.org/3/library/tkinter.html) – interface gráfica
- [ttk](https://docs.python.org/3/library/tkinter.ttk.html) – widgets estilizados
- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/) – leitura de arquivos Excel
- [Chardet](https://pypi.org/project/chardet/) – detecção de codificação de arquivos

### 🛠️ **Utilitários**

- Editor de código: **[Pycharm](https://www.jetbrains.com/pt-br/pycharm/)** 

## ▶ Como executar o projeto

#### Criando um ambiente virtual:

1 - Navegue até o diretório onde deseja criar o ambiente virtual:

```bash
 cd /path/to/your/project
```

2 - Crie um ambiente virtual:

```bash
 python3 -m venv name
```

3 - Ative o ambiente virtual:

```bash
 name\Scripts\activate
```

#### Instalação de bibliotecas:

```bash
 pip install tkinter
```

```bash
 pip install webbrowser
```

```bash
 pip install threading
```

```bash
 pip install openpyxl 
```

```bash
 pip install pandas
```

```bash
 pip install chardet
```

#### Importação de bibliotecas:

```bash
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, ttk, simpledialog
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import os
import threading
from tkinter import Menu
import chardet
```
## 📜 Licença

### Este projeto está sob licença do MIT.
<br>
Desenvolvido por Miguel Marsico 👋🏻

