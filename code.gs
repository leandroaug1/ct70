/**
 * ============================================================================
 * BACKEND - API PARA GITHUB PAGES (Code.gs)
 * ============================================================================
 */

const DB_CONFIG = {
  ID: "1brTjhihRXDJncbKGxq9M3aAJmi1E4UytB_F5KDxRS3E", 
  TABELA: "CT_70_Status" 
};

// Responde a requisições de LEITURA (GET)
function doGet(e) {
  return handleRequest('GET_DASHBOARD', null);
}

// Responde a requisições de ESCRITA (POST)
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    return handleRequest(payload.action, payload.data);
  } catch (error) {
    return createJSONOutput({ success: false, error: "JSON inválido: " + error.message });
  }
}

// Roteador Central
function handleRequest(action, data) {
  const lock = LockService.getScriptLock();
  try {
    const service = new CT70Service();
    let result;

    // Adiciona delay para evitar conflito em escritas
    if (action !== 'GET_DASHBOARD') lock.tryLock(5000);

    switch (action) {
      case 'GET_DASHBOARD':
        result = service.getDashboardData();
        break;
      case 'VALIDATE_TICKET':
        result = service.validarTickets(data.ids);
        break;
      case 'UPDATE_STATUS':
        result = service.atualizarStatus(data.id, data.status);
        break;
      case 'DELETE_TICKET':
        result = service.excluirTicket(data.id);
        break;
      case 'CLEAR_VALIDATIONS':
        result = service.limparValidacoes();
        break;
      default:
        throw new Error(`Ação desconhecida: ${action}`);
    }
    
    return createJSONOutput({ success: true, data: result });

  } catch (e) {
    return createJSONOutput({ success: false, error: e.message });
  } finally {
    lock.releaseLock();
  }
}

// Cria a resposta JSON formatada para o navegador
function createJSONOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- MANTIVE A LÓGICA DE NEGÓCIO INTACTA ---
class CT70Service {
  constructor() { this.repo = new Repository(DB_CONFIG.ID, DB_CONFIG.TABELA); }

  getDashboardData() {
    const rawData = this.repo.findAll();
    if (!rawData || rawData.length < 2) return [];
    
    const headers = rawData[0].map(h => String(h).trim().toLowerCase());
    const rows = rawData.slice(1);

    const getIdx = (keys) => {
      for (const k of keys) { const i = headers.indexOf(k.toLowerCase()); if (i > -1) return i; }
      return -1;
    };

    const idx = {
      id: getIdx(['ordem', 'os', 'id']),
      statusFinal: getIdx(['status final']),
      obsFluxo: getIdx(['observação fluxo', 'observacao fluxo']),
      permanencia: getIdx(['permanência', 'permanencia']),
      aging: getIdx(['aging fluxo', 'aging']),
      status: getIdx(['status', 'obs custom']),
      responsavel: getIdx(['resp. fluxo', 'resp fluxo']),
      stf: getIdx(['st_f', 'stf']),
      topPerm: getIdx(['top permanência', 'top permanencia'])
    };

    if (idx.id === -1) throw new Error("Coluna 'Ordem' não encontrada.");

    return rows.map((r, i) => {
      if (!r[idx.id]) return null;
      const val = (ix) => (ix > -1 ? r[ix] : "");
      return {
        id: r[idx.id],
        statusFinal: val(idx.statusFinal),
        obsFluxo: String(val(idx.obsFluxo)),
        permanencia: Number(val(idx.permanencia)) || 0,
        aging: String(val(idx.aging)),
        status: String(val(idx.status)),
        responsavel: val(idx.responsavel),
        topPerm: String(val(idx.topPerm)),
        isValidado: String(val(idx.topPerm)).includes('*')
      };
    }).filter(i => i !== null);
  }

  validarTickets(ids) {
    const colTop = this.repo.findColIndex(['top permanência', 'top permanencia']);
    const colObs = this.repo.findColIndex(['observação fluxo', 'observacao fluxo']);
    if (colTop === -1) throw new Error("Crie a coluna 'Top Permanência' na planilha.");
    ids.forEach(id => {
      this.repo.updateCell(id, colTop, v => v.startsWith('*') ? v : '*' + v);
      if (colObs !== -1) this.repo.updateCell(id, colObs, () => "Tiquete Validado");
    });
    return true;
  }
  atualizarStatus(id, st) {
    const col = this.repo.findColIndex(['status']);
    if (col === -1) throw new Error("Coluna 'Status' não encontrada.");
    this.repo.updateCell(id, col, () => st);
    return true;
  }
  excluirTicket(id) { this.repo.deleteRow(id); return true; }
  limparValidacoes() {
    const col = this.repo.findColIndex(['top permanência']);
    if (col === -1) return false;
    this.repo.updateColumn(col, v => v.startsWith('*'), v => v.replace(/^\*/, ''));
    return true;
  }
}

class Repository {
  constructor(id, tab) {
    this.ss = SpreadsheetApp.openById(id);
    this.sheet = this.ss.getSheetByName(tab);
    if (!this.sheet) throw new Error(`Aba '${tab}' não encontrada.`);
  }
  findAll() { return this.sheet.getDataRange().getValues(); }
  findColIndex(names) {
    const h = this.sheet.getRange(1, 1, 1, this.sheet.getLastColumn()).getValues()[0];
    const n = h.map(x => String(x).trim().toLowerCase());
    for (const name of names) { const i = n.indexOf(name); if (i > -1) return i + 1; }
    return -1;
  }
  findRowIndex(uid) {
    const col = this.findColIndex(['ordem', 'os', 'id']);
    if (col === -1) return -1;
    const d = this.sheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) { if (String(d[i][col-1]).trim() == String(uid).trim()) return i + 1; }
    return -1;
  }
  updateCell(id, col, cb) { const r = this.findRowIndex(id); if (r > -1) { const c = this.sheet.getRange(r, col); c.setValue(cb(String(c.getValue()))); }}
  deleteRow(id) { const r = this.findRowIndex(id); if (r > -1) this.sheet.deleteRow(r); }
  updateColumn(col, cond, trans) {
    const lr = this.sheet.getLastRow(); if (lr < 2) return;
    const rng = this.sheet.getRange(2, col, lr-1, 1); const v = rng.getValues();
    let chg = false; for(let i=0;i<v.length;i++){ if(cond(String(v[i][0]))){ v[i][0]=trans(String(v[i][0])); chg=true;}}
    if(chg) rng.setValues(v);
  }
}
