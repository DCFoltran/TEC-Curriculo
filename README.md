# TEC-Currículo

Scripts de automatização de extração de dados e Prompt de análise dos Projetos Pedagógicos

## 📦 Instalação

Para executar o **TEC-Currículo**, é necessário configurar um ambiente computacional adequado e instalar dependências essenciais para garantir o correto funcionamento dos scripts, integridade dos dados e reprodutibilidade da pesquisa. 

---

## ⚙️ Requisitos

* **Sistema operacional**: Windows 10/11, Linux ou macOS
* **Python**: versão 3.10 ou superior
* **Internet**: necessária para coleta de dados e uso do Google Gemini
* **Conta Google** com acesso ao **Google Gemini (Gems)**
* **Editor de planilhas**: Excel ou LibreOffice Calc

---

## 🐍 Instalação do Python

Verifique se o Python já está instalado:

```bash
python --version
# ou
py --version
```

Caso não esteja instalado, faça o download:
👉 https://www.python.org/downloads/

---

## 📚 Instalação das dependências

Instale os módulos necessários via `pip`:

```bash
pip install pandas
pip install openpyxl
pip install tkinter
```

### Principais bibliotecas

* **pandas** → manipulação de dados
* **openpyxl** → geração de arquivos XLSX
* **tkinter** → interface gráfica (geralmente já incluída)

---

## 🤖 Configuração do Google Gemini (Gems)

1. Acesse o menu **Gems** no Google Gemini
2. Crie um novo agente preenchendo:

* **Nome**: `Analisador de PPC`
* **Instruções**: inserir o prompt de análise definido na pesquisa
* **Ferramenta padrão**: `Deep Research`

---

## ✅ Resultado esperado

Após a configuração:

* Execução correta dos scripts
* Processamento consistente dos dados
* Geração de relatórios estruturados
* Integração com análise via IA (Gemini)

---
