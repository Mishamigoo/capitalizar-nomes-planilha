// Google Apps Script (.gs) – também pode ser usado como JavaScript (.js)

// Função que capitaliza nomes, deixa preposições e palavras específicas em minúsculo e trata exceções,
// incluindo nomes com apóstrofos como "d'Albuquerque" ou "Sant'Anna".
function CapitalizarNome(texto) {
  if (typeof texto !== 'string') return texto;

  const preposicoes = ['de', 'da', 'do', 'das', 'dos', 'e'];
  const palavrasMinusculas = ['casado', 'casada', 'solteiro', 'solteira', 'viúvo', 'viúva'];
  const excecoesFixas = {
    "sem identificação": "Sem identificação"
  };

  const textoNormalizado = texto.trim().toLowerCase();
  if (excecoesFixas[textoNormalizado]) return excecoesFixas[textoNormalizado];

  let palavras = texto.replace(/viuvo/g, "viúvo").replace(/viuva/g, "viúva").split(' ');

  palavras = palavras.map((palavra, i) => {
    const palavraLimpa = palavra.toLowerCase().replace(/[^\wáéíóúâêôãõç()']/gi, '');

    if (palavrasMinusculas.includes(palavraLimpa)) {
      return palavra.toLowerCase();
    }

    if (palavra.includes("'")) {
      const partes = palavra.split("'");
      if (partes.length === 2 && partes[0].toLowerCase() === 'd') {
        return "d'" + partes[1].charAt(0).toUpperCase() + partes[1].slice(1).toLowerCase();
      }
      return partes.map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join("'");
    }

    if (i === 0 || !preposicoes.includes(palavraLimpa)) {
      return palavra.charAt(0).toUpperCase() + palavra.slice(1).toLowerCase();
    }

    return palavra.toLowerCase();
  });

  return palavras.join(' ');
}

// Função que aplica a CapitalizarNome ao intervalo E3:L559 de uma planilha do Google Sheets
function capitalizarIntervalo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange("E3:L559");
  const values = range.getValues();

  values.forEach((linha, i) => {
    linha.forEach((celula, j) => {
      if (typeof celula === "string") {
        values[i][j] = CapitalizarNome(celula);
      }
    });
  });

  range.setValues(values);
}
