# Tutorial: Criando o Executável (.exe)

Para transformar sua automação em um arquivo que funciona com um clique (sem precisar abrir o terminal), usaremos o `PyInstaller`.

## Passo 1: Instalar o PyInstaller
Abra o seu terminal (PowerShell) e instale a ferramenta usando o Python 3:
```powershell
python3 -m pip install pyinstaller
```

## Passo 2: Criar o Executável
Abra o seu terminal na pasta do projeto e rode o comando abaixo:

```powershell
python3 -m PyInstaller --noconsole --onefile --add-data "graph_email_monitor.py;." --name "automacao-email-alex" nf_automation_gui.py
```

### O que o comando faz?
*   `--noconsole`: Esconde a janela preta do terminal quando você abrir o programa.
*   `--onefile`: Gera apenas um arquivo `.exe` final (mais limpo).
*   `--add-data`: Garante que o script de monitoramento seja incluído no pacote.
*   `--name`: O nome que você quer dar ao programa (`automacao-email-alex`).

## Dica: "pyinstaller não é reconhecido"
Se o comando `pyinstaller` der erro de "não reconhecido", tente usar o prefixo `python -m`:

```powershell
python -m PyInstaller [RESTO DO COMANDO]
```

## Passo 3: Onde encontrar o arquivo?
Após o comando terminar, uma pasta chamada **`dist`** será criada. 
Seu executável **`automacao-email-alex.exe`** estará lá!

> [!TIP]
> **Inteligência de Caminho**: O programa agora é "esperto". Mesmo que você execute o arquivo diretamente de dentro da pasta `dist`, ele automaticamente salvará as notas na pasta raiz (`Automacao-Email-Alex`).

---

> [!IMPORTANT]
> **Dependências Externas**: 
> Lembre-se que o computador que for rodar o executável ainda precisa ter o **pdftotext** instalado no Windows (PATH), conforme as instruções da V1, pois o script chama ele externamente para ler os PDFs.

> [!TIP]
> Você pode criar um atalho para este arquivo `.exe` na sua Área de Trabalho para facilitar o acesso diário.
