# Tutorial: Instalação do pdftotext (Xpdf) no Windows

O programa de automação precisa do **pdftotext** para conseguir ler as informações de dentro das Notas Fiscais em PDF. Siga estes passos para instalar em um novo computador:

## Passo 1: Baixar as ferramentas
1. Acesse o site oficial do Xpdf: [www.xpdfreader.com/download.html](https://www.xpdfreader.com/download.html)
2. Procure pela seção **"Xpdf command line tools"** para Windows.
3. Baixe o arquivo ZIP (geralmente chamado `xpdf-tools-win-4.xx.zip`).

## Passo 2: Extrair os arquivos
1. Abra o arquivo ZIP baixado.
2. Entre na pasta `bin64` (para Windows 64 bits) ou `bin32` (para 32 bits).
3. Copie todos os arquivos dessa pasta para um local definitivo no seu computador.
   * **Sugestão**: Crie uma pasta em `C:\Program Files\xpdf-tools` e cole os arquivos lá.

## Passo 3: Adicionar ao PATH do Windows
Para que o programa encontre o `pdftotext` de qualquer lugar, você precisa avisar o Windows onde ele está:

1. Clique no botão **Iniciar** e digite: `variáveis de ambiente`.
2. Selecione a opção **"Editar as variáveis de ambiente do sistema"**.
3. Na janela que abrir, clique no botão **"Variáveis de Ambiente..."** (embaixo à direita).
4. Na seção **"Variáveis do sistema"** (a de baixo), procure pela variável chamada **`Path`** e clique duas vezes nela (ou clique nela e depois em **Editar**).
5. Na nova janela, clique no botão **Novo**.
6. Cole o caminho da pasta onde você salvou os arquivos (ex: `C:\Program Files\xpdf-tools`).
7. Clique em **OK** em todas as janelas para fechar.

## Passo 4: Verificar se funcionou
1. Abra um novo terminal (**PowerShell** ou **Prompt de Comando**).
   * *Atenção: Se o terminal já estava aberto, feche e abra de novo.*
2. Digite o seguinte comando e dê Enter:
   ```powershell
   pdftotext -v
   ```
3. Se aparecer uma mensagem com a versão do Xpdf (ex: `pdftotext version 4.04`), a instalação foi um **sucesso!** 🎉

---
**Agora seu computador está pronto para rodar a Automação Email Alex!**
