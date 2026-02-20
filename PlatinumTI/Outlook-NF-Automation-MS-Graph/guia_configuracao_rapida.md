# Guia de Configuração: Automação E-mail Alex

Para rodar este programa em um novo computador, siga os passos abaixo:

## 1. Configurar o Leitor de PDF (Obrigatório)
O programa precisa de um componente externo para ler as Notas Fiscais.
1. Abra a pasta `Dependencias\Poppler\bin` que você deve enviar junto com o programa.
2. Copie o caminho dessa pasta (algo como `C:\...\Pacote-Distribuicao\Dependencias\Poppler\bin`).
3. Adicione esse caminho ao **PATH** do Windows:
   * No menu Iniciar, digite "Variáveis de ambiente" e selecione "Editar as variáveis de ambiente do sistema".
   * Clique em **Variáveis de Ambiente**.
   * Em "Variáveis do Sistema", selecione **Path** e clique em **Editar**.
   * Clique em **Novo** e cole o caminho da pasta `bin`.
   * Clique em OK em todas as janelas.

## 2. Iniciar o Programa
1. Clique duas vezes no arquivo **`automacao-email-alex.exe`**.
2. O Windows mostrará um aviso (SmartScreen). Clique em **"Mais informações"** e depois em **"Executar assim mesmo"**.

## 3. Primeiro Acesso (Autenticação)
1. Clique em **Iniciar Sincronização**.
2. O navegador abrirá para você fazer login na sua conta Microsoft.
3. Após o login, a página dará um erro (normal). **Copie a URL inteira da barra de endereços**.
4. Volte para o programa, cole a URL no campo indicado e confirme.

---

**Pronto!** O programa criará uma pasta chamada `Notas-Salvas-xml-email` na raiz do projeto para organizar os arquivos.
