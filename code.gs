/**
 * ============================================================================
 * BACKEND - SISTEMA ERP CT70 (Code.gs)
 * ============================================================================
 */

const DB_CONFIG = {
  // ID da sua planilha
  ID: "1brTjhihRXDJncbKGxq9M3aAJmi1E4UytB_F5KDxRS3E", 
  // Nome exato da aba (Deve ser igual na planilha)
  TABELA: "CT_70_Status" 
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ERP Manager | CT70')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * API DO SISTEMA
 */
function api(endpoint, payload) {
  const lock = LockService.getScriptLock();
  try {
    const service = new CT70Service();
    
    switch (endpoint) {
      case 'GET_DASHBOARD':
        return response(service.getDashboardData());
        
      case 'VALIDATE_TICKET':
        lock.tryLock(5000); 
        return response(service.validarTickets(payload.ids));
        
      case 'UPDATE_STATUS':
        lock.tryLock(5000);
        return response(service.atualizarStatus(payload.id, payload.status));
        
      case 'DELETE_TICKET':
        lock.tryLock(5000);
        return response(service.excluirTicket(payload.id));
        
      case 'CLEAR_VALIDATIONS':
        lock.tryLock(5000);
        return response(service.limparValidacoes());
        
      default:
        throw new Error(`Endpoint ${endpoint} desconhecido.`);
    }
  } catch (e) {
    Logger.log("ERRO API: " + e.message);
    return errorResponse(e.message);
  } finally {
    lock.releaseLock();
  }
}

function response(data) { return { success: true, data: data }; }
function errorResponse(msg) { return { success: false, error: msg }; }

// =========================================================================
// SERVICE - REGRAS DE NEGÓCIO COM PROTEÇÃO DE ERRO
// =========================================================================
class CT70Service {
  constructor() {
    this.repo = new Repository(DB_CONFIG.ID, DB_CONFIG.TABELA);
  }

  getDashboardData() {
    const rawData = this.repo.findAll();
    
    if (!rawData || rawData.length < 2) {
      Logger.log("Aviso: Planilha vazia ou sem cabeçalho.");
      return [];
    }

    const headers = rawData[0].map(h => String(h).trim().toLowerCase());
    const rows = rawData.slice(1);

    // Mapeamento Inteligente (Aceita variações de nome)
    const mapIndex = (chaves) => {
      for (let chave of chaves) {
        const idx = headers.indexOf(chave.toLowerCase());
        if (idx !== -1) return idx;
      }
      return -1;
    };

    const idx = {
      id: mapIndex(['ordem', 'os', 'id', 'ticket']),
      statusFinal: mapIndex(['status final']),
      obsFluxo: mapIndex(['observação fluxo', 'observacao fluxo', 'obs fluxo']),
      permanencia: mapIndex(['permanência', 'permanencia']),
      aging: mapIndex(['aging fluxo', 'aging']),
      status: mapIndex(['status', 'obs custom']),
      responsavel: mapIndex(['resp. fluxo', 'resp fluxo', 'responsavel']),
      stf: mapIndex(['st_f', 'stf']),
      topPerm: mapIndex(['top permanência', 'top permanencia'])
    };

    if (idx.id === -1) {
      throw new Error(`Coluna 'Ordem' não encontrada na planilha. Cabeçalhos lidos: ${headers.join(', ')}`);
    }

    return rows.map((r, i) => {
      // Pula linhas sem ID
      if (!r[idx.id]) return null;
      
      const getVal = (index) => (index > -1 && r[index] !== undefined) ? r[index] : "";

      const obj = { 
        rowIndex: i + 2,
        id: r[idx.id],
        statusFinal: getVal(idx.statusFinal),
        obsFluxo: String(getVal(idx.obsFluxo)),
        permanencia: Number(getVal(idx.permanencia)) || 0,
        aging: getVal(idx.aging) instanceof Date ? getVal(idx.aging).toLocaleDateString() : String(getVal(idx.aging)),
        status: String(getVal(idx.status)),
        responsavel: getVal(idx.responsavel),
        stf: getVal(idx.stf),
        topPerm: String(getVal(idx.topPerm))
      };

      obj.isValidado = obj.topPerm.includes('*');
      return obj;
    }).filter(item => item !== null);
  }

  validarTickets(ids) {
    const colTop = this.repo.findColIndex(['top permanência', 'top permanencia']);
    
    // Se não achar a coluna, lança erro explicativo
    if (colTop === -1) throw new Error("Crie a coluna 'Top Permanência' na planilha (Coluna I) para validar.");

    const colObs = this.repo.findColIndex(['observação fluxo', 'observacao fluxo']);

    ids.forEach(id => {
      this.repo.updateCell(id, colTop, val => val.startsWith('*') ? val : '*' + val);
      if (colObs !== -1) {
        this.repo.updateCell(id, colObs, () => "Tiquete Validado");
      }
    });
    return true;
  }

  atualizarStatus(id, novoStatus) {
    const colStatus = this.repo.findColIndex(['status']);
    if (colStatus === -1) throw new Error("Coluna 'Status' não encontrada na planilha.");
    
    this.repo.updateCell(id, colStatus, () => novoStatus);
    return true;
  }

  excluirTicket(id) {
    this.repo.deleteRow(id);
    return true;
  }

  limparValidacoes() {
    const colTop = this.repo.findColIndex(['top permanência', 'top permanencia']);
    if (colTop === -1) return false;
    
    this.repo.updateColumn(colTop, val => val.startsWith('*'), val => val.replace(/^\*/, ''));
    return true;
  }
}

// =========================================================================
// REPOSITORY
// =========================================================================
class Repository {
  constructor(ssId, sheetName) {
    try {
      this.ss = SpreadsheetApp.openById(ssId);
      this.sheet = this.ss.getSheetByName(sheetName);
      if (!this.sheet) throw new Error(`Aba '${sheetName}' não existe.`);
    } catch(e) {
      throw new Error(`Erro Planilha: ${e.message}`);
    }
  }

  findAll() {
    return this.sheet.getDataRange().getValues();
  }

  findColIndex(possiveisNomes) {
    const lastCol = this.sheet.getLastColumn();
    if (lastCol === 0) return -1;
    
    const headers = this.sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const headersNorm = headers.map(h => String(h).trim().toLowerCase());
    
    for (let nome of possiveisNomes) {
      const idx = headersNorm.indexOf(nome.toLowerCase());
      if (idx !== -1) return idx + 1; 
    }
    return -1;
  }

  findRowIndex(uniqueId) {
    const colId = this.findColIndex(['ordem', 'os', 'id', 'ticket']);
    if (colId === -1) return -1;
    
    const data = this.sheet.getDataRange().getValues();
    const colArrayIdx = colId - 1; 

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colArrayIdx]).trim() == String(uniqueId).trim()) return i + 1;
    }
    return -1;
  }

  updateCell(id, colIndex, callbackValue) {
    const row = this.findRowIndex(id);
    if (row === -1) return;
    const cell = this.sheet.getRange(row, colIndex);
    const val = String(cell.getValue());
    cell.setValue(typeof callbackValue === 'function' ? callbackValue(val) : callbackValue);
  }

  deleteRow(id) {
    const row = this.findRowIndex(id);
    if (row !== -1) this.sheet.deleteRow(row);
  }

  updateColumn(colIndex, conditionFn, transformFn) {
    const lastRow = this.sheet.getLastRow();
    if (lastRow < 2) return;
    
    const range = this.sheet.getRange(2, colIndex, lastRow - 1, 1);
    const values = range.getValues();
    let changed = false;

    for (let i = 0; i < values.length; i++) {
      if (conditionFn(String(values[i][0]))) {
        values[i][0] = transformFn(String(values[i][0]));
        changed = true;
      }
    }
    if (changed) range.setValues(values);
  }
}
