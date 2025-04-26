// Função que capitaliza nomes, deixa preposições e palavras específicas em minúsculo e trata exceções
function CapitalizarNome(texto) {
  if (typeof texto !== 'string') return texto;

  const preposicoes = ['de', 'da', 'do', 'das', 'dos', 'e'];
  const palavrasMinusculas = ['casado', 'casada', 'solteiro', 'solteira', 'viúvo', 'viúva'];
  const excecoesFixas = {
    "sem identificação": "Sem identificação"
  };

  // Normaliza o texto e aplica as exceções
  const textoNormalizado = texto.trim().toLowerCase();
  if (excecoesFixas[textoNormalizado]) return excecoesFixas[textoNormalizado];

  // Substitui palavras específicas como "viuvo" e "viuva" por "viúvo" e "viúva"
  let palavras = texto.replace(/\bviuvo\b/g, "viúvo").replace(/\bviuva\b/g, "viúva").split(' ');

  // Itera sobre as palavras e aplica as regras de capitalização
  palavras = palavras.map((palavra, i) => {
    const palavraLimpa = palavra.toLowerCase().replace(/[^\wáéíóúâêôãõç()']/gi, ''); // Remove pontuação (mantém parênteses e apóstrofos)

    // Se for uma preposição ou palavra específica para minúsculo, mantém em minúsculo
    if (palavrasMinusculas.includes(palavraLimpa)) {
      return palavra.toLowerCase();
    }

    // Verifica palavras com apóstrofo (como "d'Albuquerque" ou "sant'Anna") para capitalizar corretamente
    if (palavra.includes("'")) {
      let partes = palavra.split("'"); // Divide a palavra pelo apóstrofo

      // A parte antes do apóstrofo vai em minúsculo (caso seja preposição como "de")
      // A parte depois do apóstrofo começa com maiúscula
      let parteAntesApostrofo = partes[0].toLowerCase();
      let parteDepoisApostrofo = partes[1].charAt(0).toUpperCase() + partes[1].slice(1).toLowerCase();

      // Caso "D'" ou outras variantes, mantém "d" minúsculo
      if (parteAntesApostrofo === 'd') {
        return parteAntesApostrofo + "'" + parteDepoisApostrofo; // Ex: "d'Albuquerque"
      }

      // Caso contrário, capitaliza a primeira parte antes do apóstrofo
      return parteAntesApostrofo.charAt(0).toUpperCase() + parteAntesApostrofo.slice(1) + "'" + parteDepoisApostrofo; // Ex: "Sant'Anna"
    }

    // Se não for uma preposição, capitaliza a palavra
    if (i === 0 || !preposicoes.includes(palavraLimpa)) {
      return palavra.charAt(0).toUpperCase() + palavra.slice(1).toLowerCase();
    }

    // Mantém em minúsculo se for preposição
    return palavra.toLowerCase();
  });

  return palavras.join(' ');
}

// Função que aplica a CapitalizarNome ao intervalo E3:L559
function capitalizarIntervalo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange("F566:G568");
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
