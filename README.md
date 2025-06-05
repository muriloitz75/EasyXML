# EasyXML - Leitor de Notas Fiscais XML 📄➡️📊

Este projeto permite extrair informações de notas fiscais eletrônicas (NF-e) em XML e salvar os dados em um arquivo Excel.

## 🚀 Como usar

1. **Instale o Python** (se ainda não tiver)  
2. **Clone este repositório**  
   ```bash
   git clone https://github.com/caioreis29974/EasyXML.git
   cd EasyXML
   ```
3. **Instale as dependências**
   ```
   pip install -r requirements.txt
   ```
4. **Coloque os arquivos XML na pasta nfs/**
5. **Execute o script**
   ```
   python main.py
   ```
6. **Verifique a planilha gerada em notas_processadas/NotasFiscais.xlsx**

## 📂 Estrutura de Arquivos

- **EasyXML/**
  - `nfs/` → Coloque seus arquivos XML aqui
  - `notas_processadas/` → A planilha será salva automaticamente aqui
  - `main.py` → Código principal
  - `requirements.txt` → Bibliotecas necessárias
  - `README.md` → Instruções de uso

## 🔧 Tecnologias usadas

- Python 3
- xmltodict para processar XML
- pandas para manipulação de dados
- openpyxl para salvar arquivos Excel

## 👨‍💻 Criado por CaioXyZ
