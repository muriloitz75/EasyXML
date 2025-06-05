# EasyXML - Processador de Notas Fiscais

## Sobre o Aplicativo

EasyXML é uma ferramenta para processar arquivos XML de notas fiscais eletrônicas (NF-e) e notas fiscais de serviço eletrônicas (NFS-e), extraindo informações relevantes para um arquivo Excel.

## Funcionalidades

- Processa arquivos XML de notas fiscais
- Extrai informações como número da nota, data de emissão, competência, valores, etc.
- Gera um arquivo Excel organizado com os dados extraídos
- Formata a competência como MM/AAAA
- Calcula o valor líquido (Base de Cálculo - Valor ISS)
- Interface gráfica amigável

## Como Usar

### Usando o Executável

1. Faça o download do arquivo executável na pasta `dist/EasyXML`
2. Execute o arquivo `EasyXML.exe`
3. Na interface, selecione o diretório contendo os arquivos XML
4. Selecione o diretório de saída para o arquivo Excel
5. Clique em "Processar XMLs"
6. Após o processamento, o arquivo Excel será gerado no diretório de saída

### Estrutura de Diretórios

- **nfs/**: Diretório onde devem ser colocados os arquivos XML das notas fiscais
- **Notas_Processadas/**: Diretório onde será salvo o arquivo Excel com os dados extraídos

## Requisitos do Sistema

- Windows 7/8/10/11
- Não requer instalação de Python ou outras dependências

## Informações Técnicas

O aplicativo foi desenvolvido em Python e convertido para executável usando PyInstaller. Ele utiliza as seguintes bibliotecas:

- xmltodict: Para converter XML em dicionários Python
- pandas: Para manipulação de dados e criação de DataFrames
- openpyxl: Para exportar dados para Excel
- tkinter: Para a interface gráfica

## Solução de Problemas

Se você encontrar algum problema ao usar o aplicativo, verifique:

1. Se os arquivos XML estão no formato correto (NF-e ou NFS-e)
2. Se você tem permissão para ler/escrever nos diretórios selecionados
3. Se o arquivo Excel de saída não está aberto em outro programa

Um arquivo de log é gerado na mesma pasta do executável para ajudar na solução de problemas.

## Contato

Para sugestões, dúvidas ou relatos de problemas, entre em contato com o desenvolvedor.

---

© 2025 EasyXML - Todos os direitos reservados
