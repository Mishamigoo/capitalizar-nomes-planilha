# Capitalizar Nomes em Google Sheets

Este repositório contém um script em Google Apps Script que capitaliza nomes próprios em uma planilha do Google Sheets.

## Funcionalidades

- Coloca a **primeira letra maiúscula** de cada nome.
- Deixa preposições como `de`, `da`, `do`, `dos`, `e` em **minúsculo**, exceto se estiverem no início.
- Trata palavras como **"viúvo", "viúva", "casado", "solteiro"** para manter em minúsculo.
- Corrige automaticamente a grafia de `viuvo` e `viuva` para `viúvo` e `viúva`.
- Trata nomes com **apóstrofos** de forma especial:
  - `"D'Albuquerque"` vira `"d'Albuquerque"` (minúsculo antes do apóstrofo)
  - `"sant'anna"` vira `"Sant'Anna"` (primeira letra maiúscula após o apóstrofo)

## Como usar

1. Acesse o menu `Extensões > Apps Script` em sua planilha do Google Sheets.
2. Copie o conteúdo de `capitalizarNomes.gs` para o editor.
3. Salve e execute a função `capitalizarIntervalo()`.
4. O intervalo `E3:L559` será atualizado com os nomes formatados corretamente.

---

Feito com carinho pra deixar sua planilha linda ✨
