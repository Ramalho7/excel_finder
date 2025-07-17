# Excel Finder

## Descrição

O **Excel Finder** é uma aplicação desktop desenvolvida em Python que permite buscar rapidamente por um termo específico em todos os arquivos Excel de uma pasta. Com uma interface moderna e intuitiva baseada em `customtkinter`, o programa facilita a localização de informações em múltiplas planilhas, otimizando o tempo do usuário.

## Funcionalidades

- Seleção fácil de uma pasta contendo arquivos Excel (`.xlsx` ou `.xls`)
- Busca exata por um termo em todas as planilhas de todos os arquivos da pasta
- Exibição dos resultados detalhados:
  - Valor encontrado
  - Nome do arquivo
  - Nome da planilha
  - Localização da célula
- Abertura automática do arquivo Excel ao encontrar o termo
- Início da busca ao pressionar o botão ou a tecla **Enter**
- Mensagens de aviso caso a pasta ou o termo não sejam informados

## Tecnologias Utilizadas

- **Python**: Linguagem principal
- **openpyxl**: Manipulação de arquivos Excel
- **customtkinter**: Interface gráfica
- **tkinter**: Diálogos e mensagens
- **os**: Manipulação de arquivos e diretórios

## Como Usar

1. **Instale as dependências**:
   ```bash
   pip install openpyxl customtkinter
   ```

2. **Execute o programa**:
   ```bash
   python app.py
   ```

3. **Passos na interface**:
   - Clique em **"Selecione a pasta"** e escolha a pasta com os arquivos Excel.
   - Digite o termo desejado no campo de busca.
   - Clique em **"Buscar"** ou pressione **Enter**.
   - Veja os resultados na caixa de texto. Clique no resultado para abrir o arquivo correspondente.

## Como gerar um executável

Se desejar distribuir o programa sem depender do Python instalado:

1. Instale o PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Gere o executável:
   ```bash
   pyinstaller --onefile --noconsole app.py
   ```
3. O executável estará na pasta `dist/`.

## Estrutura do Projeto

```
├── app.py
├── readme.md
├── .gitignore
└── dist/
```

## Licença

Distribuído sob a licença MIT.