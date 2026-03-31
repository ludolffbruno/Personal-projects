# Guia de Configuração: Automação E-mail Alex

Para rodar este programa em um novo computador, siga os passos abaixo:

## 1. Configurar o Leitor de PDF (Obrigatório)
O programa precisa de um componente externo para ler as Notas Fiscais.
1. Se o programa não encontrar o `pdftotext` no sistema, ele tentará usar a pasta `xpdf-tools-win-4.06` que acompanha o projeto.
2. Para melhor performance, recomenda-se instalar o `pdftotext` seguindo o `guia_instalacao_pdftotext.md`.

## 2. Iniciar o Programa
1. Clique duas vezes no arquivo **`automacao-email-alex.exe`**.
2. O Windows mostrará um aviso (SmartScreen). Clique em **"Mais informações"** e depois em **"Executar assim mesmo"**.

## 3. Primeiro Acesso (Autenticação)
1. Clique em **Iniciar Sincronização**.
2. O navegador abrirá para você fazer login na sua conta Microsoft.
3. Após o login bem-sucedido, o navegador tentará abrir uma página de "localhost" que dará erro (isto é esperado).
4. **Copie a URL inteira da barra de endereços** desse erro.
5. Volte para o programa, cole a URL no campo indicado e confirme.

---

**Pronto!** O programa processará as notas e criará pastas individuais em `Notas-Salvas-xml-email`.
