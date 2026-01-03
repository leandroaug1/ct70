/**
 * BACKEND - SISTEMA ERP CT70 (Versão Blindada)
 */

const DB_CONFIG = {
  ID: "1brTjhihRXDJncbKGxq9M3aAJmi1E4UytB_F5KDxRS3E", 
  TABELA: "CT_70_Status" 
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ERP Manager | CT70')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function api(endpoint, payload) {
  const lock = LockService.getScriptLock();
  try {
    const service = new CT70Service();
    switch (endpoint) {
      case 'GET_DASHBOARD': return response(service.getDashboardData());
      case 'VALIDATE_TICKET': 
        lock.tryLock(5000); return response(service.validarTickets(payload.ids));
      case 'UPDATE_STATUS': 
        lock.tryLock(5000); return response(service.atualizarStatus(payload.id, payload.status));
      case 'DELETE_TICKET': 
        lock.tryLock(5000); return response(service.excluirTicket(payload.id));
      case 'CLEAR_VALIDATIONS': 
        lock.tryLock(5000); return response(service.limparValidacoes());
      default: throw new Error(`Endpoint ${endpoint} desconhecido.`);
    }
  } catch (e) {
    return errorResponse(e.message);
  } finally {
    lock.releaseLock();
  }
}

function response(data) { return { success: true, data: data }; }
function errorResponse(msg) { return { success: false, error: msg }; }

class CT70Service {
  constructor() { this.repo = new Repository(DB_CONFIG.ID, DB_CONFIG.TABELA); }

  getDashboardData() {
    const rawData = this.repo.findAll();
    if (!rawData || rawData.length < 2) return [];

    const headers = rawData[0].map(h => String(h).trim().toLowerCase());
    const rows = rawData.slice(1);

    // Mapeamento flexível de colunas
    const getIdx = (keys) => {
      for (const k of keys) {
        const i = headers.indexOf(k.toLowerCase());
        if (i > -1) return i;
      }
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
      topPerm: getIdx(['top permanência', 'top permanencia']) // Coluna que estava faltando
    };

    if (idx.id === -1) throw new Error("Coluna 'Ordem' não encontrada na planilha.");

    return rows.map((r, i) => {
      if (!r[idx.id]) return null;
      const val = (ix) => (ix > -1 ? r[ix] : "");
      
      const obj = {
        id: r[idx.id],
        statusFinal: val(idx.statusFinal),
        obsFluxo: String(val(idx.obsFluxo)),
        permanencia: Number(val(idx.permanencia)) || 0,
        aging: String(val(idx.aging)),
        status: String(val(idx.status)),
        responsavel: val(idx.responsavel),
        topPerm: String(val(idx.topPerm))
      };
      // Se a coluna não existir, considera não validado em vez de erro
      obj.isValidado = obj.topPerm.includes('*');
      return obj;
    }).filter(i => i !== null);
  }

  validarTickets(ids) {
    const colTop = this.repo.findColIndex(['top permanência', 'top permanencia']);
    const colObs = this.repo.findColIndex(['observação fluxo', 'observacao fluxo']);
    
    if (colTop === -1) throw new Error("Crie a coluna 'Top Permanência' na planilha para validar.");

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
    const col = this.repo.findColIndex(['top permanência', 'top permanencia']);
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
    const lastCol = this.sheet.getLastColumn();
    if (lastCol === 0) return -1;
    const headers = this.sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const norms = headers.map(h => String(h).trim().toLowerCase());
    for (const n of names) {
      const i = norms.indexOf(n);
      if (i > -1) return i + 1;
    }
    return -1;
  }
  
  findRowIndex(uid) {
    const col = this.findColIndex(['ordem', 'os', 'id']);
    if (col === -1) return -1;
    const data = this.sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][col-1]).trim() == String(uid).trim()) return i + 1;
    }
    return -1;
  }

  updateCell(id, col, cb) {
    const row = this.findRowIndex(id);
    if (row > -1) {
      const cell = this.sheet.getRange(row, col);
      cell.setValue(cb(String(cell.getValue())));
    }
  }

  deleteRow(id) {
    const row = this.findRowIndex(id);
    if (row > -1) this.sheet.deleteRow(row);
  }

  updateColumn(col, cond, trans) {
    const lr = this.sheet.getLastRow();
    if (lr < 2) return;
    const rng = this.sheet.getRange(2, col, lr-1, 1);
    const vals = rng.getValues();
    let chg = false;
    for(let i=0; i<vals.length; i++) {
      if(cond(String(vals[i][0]))) { vals[i][0] = trans(String(vals[i][0])); chg = true; }
    }
    if(chg) rng.setValues(vals);
  }
}
