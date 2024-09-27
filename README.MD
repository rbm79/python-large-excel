# Processador de Excel de Alto Desempenho

Este script Python foi projetado para processar eficientemente arquivos Excel grandes, extraindo colunas específicas e criando um novo arquivo Excel otimizado.

## Características

- Leitura rápida de arquivos Excel grandes
- Extração de colunas específicas
- Escrita eficiente para um novo arquivo Excel
- Medição de desempenho (tempo e uso de memória)
- Logging detalhado do processo

## Requisitos

- Python 3.6+
- openpyxl
- psutil

Para instalar as dependências:

```powershell 
PS C:\repositorios\python-excel>pip install openpyxl psutil
```

## Uso

1. Coloque o script no mesmo diretório que o seu arquivo Excel de entrada.

2. Modifique as seguintes variáveis no script conforme necessário:
   - `input_file`: Nome do seu arquivo Excel de entrada
   - Ajuste os índices em `data.append((row[4], row[10]))` para selecionar as colunas desejadas

3. Execute o script:

```powershell
PS C:\repositorios\python-excel> & C:/Python312/python.exe c:/repositorios/python-excel/excel-process-tuple.py
```


4. O script irá gerar:
   - Um novo arquivo Excel com as colunas extraídas
   - Um arquivo de log com informações de desempenho
   - Exemplo de saida para o processamento de um arquivo de 800 mil linhas:

```powershell
PS C:\repositorios\python-excel> & C:/Python312/python.exe c:/repositorios/python-excel/excel-process-tuple.py
Iniciando leitura do arquivo Excel...
Leitura concluída - Tempo: 89.49 segundos, Memória usada: 682.73 MB
Iniciando escrita para o novo arquivo Excel...
Escrita concluída - Tempo: 163.37 segundos, Memória usada: 689.36 MB
Processo finalizado - Tempo: 163.57 segundos, Memória usada: 27.32 MB
Processo concluído. Novo arquivo Excel salvo como output_20240927_130628.xlsx
Log salvo em log_20240927_130628.txt
```

## Como Funciona

1. **Leitura do Excel**: O script lê todo o arquivo Excel de entrada e armazena os dados selecionados em uma lista de tuplas em memória.

2. **Processamento**: As colunas especificadas são extraídas durante a leitura.

3. **Escrita**: Os dados processados são escritos em um novo arquivo Excel.

4. **Logging**: O tempo de execução e o uso de memória são registrados após cada fase principal.

## Otimização de Desempenho

- Usa o modo `read_only` do openpyxl para leitura eficiente
- Emprega o modo `write_only` para escrita otimizada
- Utiliza estruturas de dados eficientes (lista de tuplas) para armazenamento em memória

## Considerações

- Este script é otimizado para velocidade, carregando todos os dados na memória. Para arquivos extremamente grandes, pode ser necessário ajustar a abordagem para processamento em lotes.
- O uso de memória pode ser significativo para arquivos muito grandes.

## Personalização

- Ajuste os índices das colunas em `data.append((row[4], row[10]))` por exemplo, para extrair diferentes colunas.
- Modifique o cabeçalho `new_ws.append(["NOME", "VALOR"])` para refletir as colunas escolhidas.

## Saída

- **Novo Excel**: `output_[timestamp].xlsx`
- **Arquivo de Log**: `log_[timestamp].txt`

O arquivo de log contém informações sobre o tempo de execução e uso de memória para cada fase do processo.

## Suporte

Para questões ou sugestões de melhorias, por favor, abra uma issue no repositório do projeto.

