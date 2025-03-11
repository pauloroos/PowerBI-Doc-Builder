# 📊 PowerBI Doc Builder

![Python](https://img.shields.io/badge/Python-3.10+-blue?style=flat-square&logo=python)
![Power BI](https://img.shields.io/badge/Power%20BI-Documentation-orange?style=flat-square&logo=powerbi)
![Tkinter](https://img.shields.io/badge/Tkinter-GUI-success?style=flat-square)
![Status](https://img.shields.io/badge/status-Ativo-brightgreen?style=flat-square)

> Geração automatizada de documentação técnica para arquivos `.PBIX` do Power BI — com análise via DAX Studio, descrição por IA (Gemini API) e diagrama relacional!

---

## 🚀 Funcionalidades
- 🧠 **Geração de descrições com IA (Google Gemini)**
- 📄 **Criação de documento Word com formatação profissional**
- 🧱 **Tabela de medidas, colunas, partições, relacionamentos, parâmetros e grupos de cálculo**
- 🖼️ **Inserção de diagrama relacional gerado automaticamente (Graphviz)**
- 💻 **Interface amigável com Tkinter + ttkbootstrap**
- ⚙️ **Configuração de caminhos e API com salvamento persistente**
- 🔍 Suporte a múltiplos arquivos `.pbix` em lote


---

## 🖼️ Capturas de Tela
### 🧭 Tela Inicial do Aplicativo
<img src="https://github.com/pauloroos/PowerBI-Doc-Builder/blob/2bde9f2e4f652220a5ccd44756776f288132bbc8/assets/aplicativo.png" alt="Tela inicial do app" width="700"/>

### ⚙️ Menu de Configurações
<img src="https://github.com/pauloroos/PowerBI-Doc-Builder/blob/2bde9f2e4f652220a5ccd44756776f288132bbc8/assets/configuracoes.png" alt="Menu de configurações do app" width="700"/>


---

## 🧪 Tecnologias Utilizadas
- Python, Tkinter, TTKBootstrap
- Pandas, Graphviz, Pillow
- python-docx, psutil, subprocess
- Google Generative AI (Gemini)

---

## 🤖 IA no Projeto
O projeto utiliza a **API do Gemini** para gerar descrições automáticas dos dashboards com base nos dados extraídos (tabelas, colunas, medidas e relacionamentos). Isso permite um overview claro e objetivo sobre o propósito do modelo semântico.

---

## 📁 Estrutura de Saída
Ao processar um ou mais arquivos `.pbix`, será gerado:

```
📁 Resultado
 ┣ 📁 Arquivos
 ┃ ┗ 📁 <nome_arquivo_pbix>
 ┃    ┣ columns.csv
 ┃    ┣ measures.csv
 ┃    ┣ ...
 ┣ 📁 Documentacao
 ┃ ┣ <nome_arquivo_pbix>.docx
 ┃ ┣ <nome_arquivo_pbix> - DIAGRAMA.png
```

---

## ⚙️ Como Usar
1. **Clone o repositório:**

```bash
git clone https://github.com/pauloroos/PowerBI-Doc-Builder.git
cd PowerBI-Doc-Builder
```

2. **Instale as dependências:**

```bash
pip install -r requirements.txt
```

3. **Configure os caminhos:**
- DAX Studio (`dscmd.exe`)
- DLL do Analysis Services (`Microsoft.AnalysisServices.dll`)
- Executável do Power BI Desktop (`PBIDesktop.exe`)
- API Key do Gemini (obtenha em https://aistudio.google.com/app/apikey)

4. **Execute o aplicativo:**

```bash
python PowerBIDocBuilderApp.py
```

---

## 🧩 Instalação e Configuração de Requisitos
Antes de rodar o PowerBI Doc Builder, certifique-se de instalar e configurar os seguintes itens:

### 📦 1. Python 3.10+
Instale a versão mais recente do Python em: https://www.python.org/downloads/

### 🖥️ 2. Power BI Desktop

1. Baixe e instale o Power BI Desktop: https://powerbi.microsoft.com/pt-br/desktop/
2. Para encontrar o caminho do executável:
   - Abra o Power BI Desktop
   - Pressione `Ctrl + Shift + Esc` para abrir o **Gerenciador de Tarefas**
   - Vá em **Detalhes** > localize `PBIDesktop.exe` > clique com o botão direito > **Abrir local do arquivo**
   - Clique com o botão direito em `PBIDesktop.exe` > **Copiar como caminho**
   - ⚠️ Remova as aspas duplas ao colar no app!

### 🧪 3. DAX Studio

1. Baixe em: https://daxstudio.org
2. Caminho do `dscmd.exe`: `C:\Program Files\DAX Studio\dscmd.exe`
3. Caminho da DLL: `C:\Program Files\DAX Studio\bin\Microsoft.AnalysisServices.dll`

### 🧠 4. Gemini API

1. Acesse https://aistudio.google.com/app/apikey
2. Gere sua chave de API do Gemini
3. Cole a chave no campo correspondente da janela de configurações

### 🔁 5. Graphviz

1. Acesse: https://graphviz.org/download/
2. Baixe e instale a versão para Windows
3. Durante a instalação, **adicione o Graphviz ao PATH**
4. Teste no terminal com:

```bash
dot -V
```

---

## 📦 Geração de Executável
Este projeto inclui um script `create_app.bat` que gera automaticamente um executável `.exe` com todos os arquivos necessários incluídos (imagens, configurações e módulos).

### ▶️ Como usar:

1. Certifique-se de ter o **PyInstaller** instalado:

```bash
pip install pyinstaller
```

2. Dê **duplo clique** no arquivo `create_app.bat` ou execute pelo terminal.

O executável será gerado na pasta `dist/` com o nome `PowerBIDocBuilderApp.exe`.

### 📂 O script inclui:

- Ícone personalizado (`icon.ico`)
- Pasta `core/` com os módulos auxiliares
- Imagens e logotipo da pasta `assets/`
- Arquivo `config.json` com as configurações persistidas

O conteúdo do script:

```bat
@echo off
cd /d %~dp0

REM Gera o executável com ícone, arquivos core, imagens e config incluídos
python -m PyInstaller --noconsole --onefile PowerBIDocBuilderApp.py ^
--icon=assets\icon.ico ^
--add-data "config.json;." ^
--add-data "assets\icon.png;assets" ^
--add-data "assets\seu-logo.png;assets" ^
--add-data "assets\image.png;assets" ^
--add-data "assets\icon.ico;assets" ^
--add-data "core\helpers.py;core" ^
--add-data "core\ai_description.py;core" ^
--add-data "core\diagram_generator.py;core" ^
--add-data "core\pbi_extractor.py;core"

pause
```

## 📁 Pasta Auxiliar
O repositório também inclui uma pasta chamada `auxiliar`, que contém:

- 📘 **Manual do Usuário** (`manual.docx`): documentação com orientações detalhadas de uso
- ⚙️ **Executáveis testados**: versões específicas do DAX Studio, Graphviz e outros utilizados durante os testes, garantindo maior compatibilidade

Essa pasta é útil para quem deseja replicar o ambiente exatamente como foi validado.

## 🙏 Referência
Este projeto foi inspirado e adaptado a partir de:

- [pbi-docs](https://github.com/alisonpezzott/pbi-docs) — criado por [@alisonpezzott](https://github.com/alisonpezzott)

Agradecimentos pela base de extração e estrutura inicial, que contribuíram significativamente para o desenvolvimento do PowerBI Doc Builder.

## 📜 Licença
MIT License © Paulo Roos
---

## 👤 Autor
**Paulo Roos**  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-paulo--roosf-blue?logo=linkedin&style=flat-square)](https://www.linkedin.com/in/pauloroosf)  
[![GitHub](https://img.shields.io/badge/GitHub-paulo--roos-black?logo=github&style=flat-square)](https://github.com/pauloroos)

---

## ⭐ Contribua
Achou útil? Deixe uma ⭐ no repositório e compartilhe com a comunidade Power BI!

---


---


---


---