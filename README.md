# Conversor de Arquivos Excel - XLS/XLT/XLSX para XLSB

Esta aplicação permite converter arquivos Excel (como `.xls`, `.xlsx`, e `.xltx`) para o formato binário `.xlsb`. O usuário pode selecionar os arquivos através de uma interface gráfica (GUI), e o programa irá convertê-los para `.xlsb`, removendo o arquivo original, se assim desejado.

## Funcionalidades
- Converte arquivos `.xls`, `.xlsx` e `.xltx` para `.xlsb`.
- Permite a seleção de múltiplos arquivos por meio de uma interface gráfica.
- Remove o arquivo original após a conversão (opcional).
- Exibe o processo de conversão em uma interface amigável.

## Requisitos
Para rodar este programa, é necessário ter as seguintes dependências instaladas:

- Python 3.x
- Tkinter (geralmente já vem incluído com o Python)
- Biblioteca `win32com` (parte do pacote `pywin32`)

### Instalando as Dependências

Primeiro, instale o Python e as bibliotecas necessárias:

```bash
pip install pywin32

Como Usar
Clone o repositório ou faça o download dos arquivos de código:

bash
Copiar código
git clone https://github.com/seu-usuario/nome-do-repositorio.git
cd nome-do-repositorio
Execute a aplicação: Execute o arquivo Python para iniciar a interface gráfica e começar a converter os arquivos:

bash
Copiar código
python main.py
Use a Interface Gráfica:

Clique em Selecionar Arquivos para escolher os arquivos Excel que você deseja converter.
Clique em Converter Arquivos para iniciar o processo de conversão.
Os arquivos convertidos para .xlsb serão salvos no mesmo diretório dos arquivos originais, e os arquivos originais podem ser removidos automaticamente, se configurado.
Limpe a lista de arquivos clicando em Limpar Lista se necessário.

Tratamento de Erros
Se ocorrerem problemas, como arquivos não encontrados ou problemas de acesso, eles serão exibidos na área de saída de log da GUI. Alguns erros comuns incluem:

Arquivo não encontrado.
Excel não instalado ou inacessível.
O arquivo está sendo usado por outro programa.
