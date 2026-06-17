/**
 * ==============================================================================
 * INSTRUÇÕES PARA CONFIGURAÇÃO DA SINCRONIZAÇÃO (GRAVAÇÃO AUTOMÁTICA)
 * ==============================================================================
 * 
 * Siga os passos abaixo para conectar o seu Dashboard à planilha de forma bidirecional:
 * 
 * 1. Abra a sua planilha do Google Sheets:
 *    https://docs.google.com/spreadsheets/d/1_Lh54dqjscfNBpRsW81qtN42xOkMPlSKshU7anh5aR4
 * 
 * 2. No menu superior, clique em "Extensões" (Extensions) -> "Apps Script".
 * 
 * 3. Apague qualquer código existente no editor e cole TODO o código Javascript
 *    que está no bloco "CÓDIGO DO SCRIPT" abaixo.
 * 
 * 4. Salve o projeto clicando no ícone do disquete (ou Ctrl + S) e dê um nome
 *    (ex: "Sincronizador Natto").
 * 
 * 5. Clique no botão azul "Implantar" (Deploy) no canto superior direito -> "Nova implantação" (New deployment).
 * 
 * 6. Na janela que abrir, clique na engrenagem de configuração de tipo ao lado de
 *    "Selecionar tipo" e selecione "App da Web" (Web app).
 * 
 * 7. Preencha as configurações da seguinte forma:
 *    - Descrição: "Sincronização Natto v1"
 *    - Executar como: "Eu" (Seu e-mail do Google)
 *    - Quem tem acesso: "Qualquer pessoa" (Anyone) <-- IMPORTANTE!
 * 
 * 8. Clique em "Implantar" (Deploy). O Google solicitará permissão de acesso à sua
 *    planilha. Siga as instruções na tela e autorize (vá em "Avançado" e "Acessar...").
 * 
 * 9. O Google gerará uma "URL do app da Web". Copie essa URL.
 * 
 * 10. Abra o seu Dashboard no navegador, clique no ícone de engrenagem (⚙️) no cabeçalho,
 *     cole a URL que você copiou e clique em "Salvar URL".
 * 
 * Pronto! Agora, sempre que você adicionar um prazo, condicionante ou licença
 * pelo dashboard, as modificações serão enviadas em tempo real para a sua Planilha Google.
 */

// ==============================================================================
// CÓDIGO DO SCRIPT (COLE ESTE BLOCO NO EDITOR DO GOOGLE APPS SCRIPT)
// ==============================================================================

function getSheetCaseInsensitive(doc, name) {
  const sheets = doc.getSheets();
  const lowerName = name.toLowerCase().trim();
  for (let i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName().toLowerCase().trim();
    if (sheetName === lowerName) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * doGet — leitura bidirecional: o dashboard faz GET para carregar os dados
 * atuais direto da planilha (evita depender de CSV público ou CORS no POST).
 */
function doGet(e) {
  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();

    // Lê aba Condicionantes
    const condSheet = getSheetCaseInsensitive(doc, "Condicionantes");
    const condData  = condSheet ? condSheet.getDataRange().getValues() : [];

    // Lê aba Licenças e Autorizações (tenta variações do nome)
    const licSheet = getSheetCaseInsensitive(doc, "Licenças e Autorizações")
                  || getSheetCaseInsensitive(doc, "Licencas e Autorizacoes")
                  || getSheetCaseInsensitive(doc, "Licenças");
    const licData  = licSheet  ? licSheet.getDataRange().getValues()  : [];

    // Converte cada linha para array de strings (igual ao que o parseCSV retornaria).
    // Datas vindas do Sheets como Date object são formatadas como DD/MM/YYYY.
    const cellToStr = (cell) => {
      if (cell === null || cell === undefined || cell === "") return "";
      if (cell instanceof Date) {
        const d = String(cell.getDate()).padStart(2, "0");
        const m = String(cell.getMonth() + 1).padStart(2, "0");
        const y = cell.getFullYear();
        return `${d}/${m}/${y}`;
      }
      return String(cell).trim();
    };
    const toStrRows = (rows) => rows.map(row => row.map(cellToStr));

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      condicionantes: toStrRows(condData),
      licencas: toStrRows(licData)
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data;
    
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    
    // Mapeamento dos códigos internos de layout para os nomes das abas/unidades na planilha
    const unitNames = {
      "bj": "Belo Jardim",
      "pe": "Pesqueira",
      "it": "Itapissuma",
      "ga": "Garanhuns"
    };
    
    if (action === "syncCondicionantes") {
      let sheet = getSheetCaseInsensitive(doc, "Condicionantes");
      if (!sheet) {
        // Cria a aba automaticamente se não existir
        sheet = doc.insertSheet("Condicionantes");
      }
      
      const headers = ["Unidade", "CNPJ", "Item", "Condicionante / Obrigação", "Data Prazo", "Protocolo / Obs.", "Status"];
      const rows = [headers];
      const cnpjs = payload.cnpjs || {};
      
      for (const un in data) {
        const items = data[un];
        items.forEach(item => {
          item.prazos.forEach(prazo => {
            // Converte a data de YYYY-MM-DD de volta para DD/MM/YYYY
            const dateParts = prazo.data.split("-");
            let displayDate = prazo.data;
            if (dateParts.length === 3) {
              displayDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
            }
            
            const statusText = prazo.status === "cumprido" ? "✅ Cumprido" : "⏳ Pendente";
            
            rows.push([
              unitNames[un] || un,
              cnpjs[un] || "",
              item.num,
              item.desc,
              displayDate,
              prazo.obs || "",
              statusText
            ]);
          });
        });
      }
      
      sheet.clearContents();
      sheet.getRange(1, 1, rows.length, 7).setValues(rows);
      return buildResponse({ success: true });
    }

    if (action === "syncLicencas") {
      let sheet = getSheetCaseInsensitive(doc, "Licenças e Autorizações")
               || getSheetCaseInsensitive(doc, "Licencas e Autorizacoes");
      if (!sheet) {
        // Cria a aba automaticamente se não existir
        sheet = doc.insertSheet("Licenças e Autorizações");
      }
      
      const headers = ["Unidade", "CNPJ", "Documento / Licença", "Órgão", "Código / Nº", "Emissão", "Validade", "Status"];
      const rows = [headers];
      const cnpjs = payload.cnpjs || {};
      
      for (const un in data) {
        const lics = data[un];
        lics.forEach(lic => {
          const convertToDMY = (isoDate) => {
            if (!isoDate) return "";
            const parts = isoDate.split("-");
            if (parts.length === 3) {
              return `${parts[2]}/${parts[1]}/${parts[0]}`;
            }
            return isoDate;
          };
          
          // Calcula automaticamente a string de status conforme a regra do negócio
          // Usa new Date(y, m-1, d) para evitar bug de fuso horário com strings ISO.
          let statusText = "⬜ Ausente";
          if (lic.validade) {
            const hoje = new Date(); hoje.setHours(0,0,0,0);
            const vp = lic.validade.split("-");
            const val = new Date(parseInt(vp[0]), parseInt(vp[1])-1, parseInt(vp[2]));
            const diff = Math.round((val - hoje) / 86400000);
            if (diff < 0) statusText = "🔴 Vencido";
            else if (diff <= 90) statusText = "⚠️ Próx. Venc.";
            else statusText = "✅ Válido";
          }
          
          rows.push([
            unitNames[un] || un,
            cnpjs[un] || "",
            lic.nome,
            lic.orgao || "",
            lic.codigo || "",
            convertToDMY(lic.emissao),
            convertToDMY(lic.validade),
            statusText
          ]);
        });
      }
      
      sheet.clearContents();
      sheet.getRange(1, 1, rows.length, 8).setValues(rows);
      return buildResponse({ success: true });
    }

    return buildResponse({ success: false, error: "Acao nao suportada." });
  } catch (err) {
    return buildResponse({ success: false, error: err.toString() });
  }
}

// Retorna JSON com headers CORS para que o dashboard possa ler a resposta sem no-cors.
function buildResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
