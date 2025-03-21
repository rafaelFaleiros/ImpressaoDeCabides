# Automacao de Impressao de Cabides

## Descricao
Este programa foi desenvolvido para facilitar a automacao do processo de impressao de arquivos em massa, permitindo a selecao de arquivos, configuracao do numero de copias e escolha da impressora de forma eficiente.

## Funcionalidades
- Listagem automatica dos arquivos disponiveis para impressao.
- Campo de busca para localizar rapidamente um arquivo.
- Selecao individual de arquivos para impressao.
- Definicao do numero de copias para cada arquivo (valor padrao: 1).
- Suporte a diferentes tamanhos de papel.
- Interface grafica intuitiva.

## Requisitos
- Python 3.13 ou superior
- Bibliotecas necessarias (instalaveis via `pip install -r requirements.txt`)
- Impressora configurada no sistema

## Instalacao
1. Clone ou baixe o repositório.
2. Instale as dependencias com:
   ```sh
   pip install -r requirements.txt
   ```
3. Execute o programa:
   ```sh
   python programa.py
   ```

## Gerando o Executavel
Para criar um executavel do programa, utilize o **PyInstaller**:
```sh
pyinstaller --onefile --windowed --icon=impressora.ico programa.py
```
O executavel sera gerado na pasta `dist/`.

## Resolucao de Problemas
### O icone do executavel nao aparece
1. Certifique-se de que o arquivo `impressora.ico` esta no formato correto.
2. Execute o comando `pyinstaller` novamente usando `--clean` para limpar os arquivos antigos:
   ```sh
   pyinstaller --clean --onefile --windowed --icon=impressora.ico programa.py
   ```
3. Se o icone nao aparecer imediatamente, reinicie o Windows Explorer ou o PC.

### Ao buscar um arquivo, os selecionados anteriormente sao desmarcados
Isso ocorre porque a busca recarrega a lista. Para corrigir, implemente um mecanismo para armazenar a selecao antes de atualizar a interface.

## Autor
Desenvolvido por [Seu Nome].

## Licenca
Este projeto esta licenciado sob a [MIT License](LICENSE).

