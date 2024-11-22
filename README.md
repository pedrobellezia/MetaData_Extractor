# Processador de Metadados de Mídia

Este módulo Python processa arquivos de mídia para extrair metadados, criar relatórios em Excel e gerar gráficos.

## Funcionalidades

- Navega por diretórios e subdiretórios especificados.
- Extrai metadados de arquivos de mídia.
- Salva os metadados em um arquivo Excel.
- Salva os metadados em um arquivo JSON.
- Gera um gráfico de barras das extensões de arquivos e insere no arquivo Excel.

## Pré-requisitos

- Python 3.x
- Módulo `pandas`
- Módulo `matplotlib`
- Módulo `pymediainfo`
- Módulo `openpyxl`

## Instalação

1. Clone este repositório:
   ```bash
   git clone https://github.com/seu_usuario/media-metadata-processor.git
   ```

2. Instale os módulos necessários:
   ```bash
   pip install -r requirements.txt
   ```

## Uso

1. Especifique os diretórios a serem processados e os nomes dos arquivos de saída no script:
   ```python
   # Diretórios a serem processados
   DIRETORIO = [r"C:\Users\seu_usuario\Documents\media"]
   # Nome do arquivo Excel
   EXCEL_NAME = "relatorio_metadados.xlsx"
   # Nome do arquivo JSON
   JSON_NAME = "relatorio_metadados.json"
   ```

2. Execute o script:
   ```bash
   python DataExtractor.py
   ```
   
## Exemplo

```json
[
    {
        "nome_arquivo": "example.mp4",
        "data_criacao": "2023-07-19",
        "data_modi": "2023-07-19",
        "duracao": "00:09:13",
        "size": "392.02 MB",
        "caminho": "C:\\Users\\seu_usuario\\Documents\\media\\example.mp4",
        "extensao": ".mp4"
    }
]
```

## Licença

Este projeto está licenciado sob a Licença MIT.
