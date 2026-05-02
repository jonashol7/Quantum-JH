// ????????????????????????????????????????????????????????????????
//  QUANTUM JH ? Gest?o de Estoque | Codigo.gs
//
//  COMO USAR:
//  1. Cole este arquivo em C?digo.gs
//  2. Cole os outros 3 arquivos HTML (Index, Estilo, Codigo_js)
//  3. Implantar > Nova implanta??o > App da Web
//     Executar como: Eu | Acesso: Qualquer pessoa
// ????????????????????????????????????????????????????????????????

var SHEET_PROD = "Produtos";
var SHEET_MOV  = "Movimentacoes";

// ?? Serve o HTML ? SEM setup() para n?o travar ??????????????????
function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Stockify")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ?? getTodos: cria planilhas se n?o existirem + retorna dados ????
function getTodos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Garante aba Produtos
  var pS = ss.getSheetByName(SHEET_PROD);
  if (!pS) {
    pS = ss.insertSheet(SHEET_PROD);
    pS.getRange(1,1,1,6).setValues([["id","name","qty","min","cost","cat"]]);
    pS.getRange(1,1,1,6).setFontWeight("bold");
    pS.getRange(2,1,5,6).setValues([
      [uid(),"Cremalheira",   12, 5,  85.00,  "Mecanica"],
      [uid(),"Correia",        8,10,  32.50,  "Transmissao"],
      [uid(),"Rolamento",     25,15,  18.90,  "Mecanica"],
      [uid(),"Caixa completa", 3, 3, 520.00,  "Componentes"],
      [uid(),"Motor Eletrico", 0, 2,1240.00,  "Eletrica"]
    ]);
  }

  // Garante aba Movimentacoes
  var mS = ss.getSheetByName(SHEET_MOV);
  if (!mS) {
    mS = ss.insertSheet(SHEET_MOV);
    mS.getRange(1,1,1,6).setValues([["id","prod_id","type","qty","obs","date"]]);
    mS.getRange(1,1,1,6).setFontWeight("bold");
  }

  var pD = pS.getDataRange().getValues();
  var mD = mS.getDataRange().getValues();

  return {
    products:  pD.slice(1).map(function(r){ return toObj(pD[0], r); }),
    movements: mD.slice(1).map(function(r){ return toObj(mD[0], r); })
  };
}

// ?? salvarProduto ????????????????????????????????????????????????
function salvarProduto(id, name, qty, min, cost, cat) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var pS   = ss.getSheetByName(SHEET_PROD);
    var mS   = ss.getSheetByName(SHEET_MOV);
    var data = pS.getDataRange().getValues();
    qty  = Number(qty)  || 0;
    min  = Number(min)  || 0;
    cost = Number(cost) || 0;
    name = String(name  || "").trim();
    cat  = String(cat   || "").trim();
    var isEdit = id && id !== "null" && String(id).trim() !== "";
    if (isEdit) {
      var sid = String(id).trim();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === sid) {
          var oldQty  = Number(data[i][2]) || 0;
          var oldCost = Number(data[i][4]) || 0;
          pS.getRange(i+1,1,1,6).setValues([[sid,name,qty,min,cost,cat]]);
          if (oldQty !== qty)
            mS.appendRow([uid(),sid,"adj",Math.abs(qty-oldQty),"Ajuste via edicao",new Date().toISOString()]);
          if (oldCost !== cost && oldCost > 0)
            mS.appendRow([uid(),sid,"price",cost,"Preco anterior: "+oldCost.toFixed(2),new Date().toISOString()]);
          return "ok";
        }
      }
      return "erro: nao encontrado id=" + id;
    } else {
      var nid = uid();
      pS.appendRow([nid,name,qty,min,cost,cat]);
      if (qty > 0) mS.appendRow([uid(),nid,"in",qty,"Estoque inicial",new Date().toISOString()]);
      return "ok";
    }
  } catch(e) { return "erro: " + e.message; }
}

// ?? excluirProduto ???????????????????????????????????????????????
function excluirProduto(id) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var pS   = ss.getSheetByName(SHEET_PROD);
    var data = pS.getDataRange().getValues();
    var sid  = String(id).trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === sid) {
        pS.deleteRow(i + 1);
        return "ok";
      }
    }
    return "erro: nao encontrado";
  } catch(e) { return "erro: " + e.message; }
}

// ?? registrarMovimentacao ????????????????????????????????????????
function registrarMovimentacao(prodId, type, qty, obs, newQty) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var pS    = ss.getSheetByName(SHEET_PROD);
    var mS    = ss.getSheetByName(SHEET_MOV);
    var data  = pS.getDataRange().getValues();
    var sid   = String(prodId).trim();
    qty    = Number(qty)    || 0;
    newQty = Number(newQty) || 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === sid) {
        pS.getRange(i+1,3).setValue(newQty);
        mS.appendRow([uid(),sid,type,qty,obs||"",new Date().toISOString()]);
        return "ok";
      }
    }
    return "erro: nao encontrado";
  } catch(e) { return "erro: " + e.message; }
}

// ?? helpers ??????????????????????????????????????????????????????
function uid() { return Utilities.getUuid(); }

function toObj(headers, row) {
  var o = {};
  headers.forEach(function(h,i){ o[String(h)] = row[i]; });
  return o;
}

// ── Importação em lote ───────────────────────────────────────────
function importarProdutos(lista) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pS = ss.getSheetByName(SHEET_PROD);
    var mS = ss.getSheetByName(SHEET_MOV);
    var dataAtual = pS.getDataRange().getValues();
    var existentes = {};
    for (var i = 1; i < dataAtual.length; i++) {
      existentes[String(dataAtual[i][1]).trim().toLowerCase()] = i + 1;
    }
    var inseridos = 0, atualizados = 0, erros = [];
    for (var j = 0; j < lista.length; j++) {
      var p = lista[j];
      try {
        var nm = String(p.name || "").trim();
        if (!nm) { erros.push("Linha " + (j+2) + ": nome vazio"); continue; }
        var qty  = Number(p.qty)  || 0;
        var min  = Number(p.min)  || 0;
        var cost = Number(p.cost) || 0;
        var cat  = String(p.cat || "").trim();
        var key  = nm.toLowerCase();
        if (existentes[key]) {
          var row = existentes[key];
          var sid = String(dataAtual[row-1][0]).trim();
          pS.getRange(row, 1, 1, 6).setValues([[sid, nm, qty, min, cost, cat]]);
          atualizados++;
        } else {
          var nid = uid();
          pS.appendRow([nid, nm, qty, min, cost, cat]);
          if (qty > 0) mS.appendRow([uid(), nid, "in", qty, "Importação inicial", new Date().toISOString()]);
          inseridos++;
        }
      } catch(e) {
        erros.push("Linha " + (j+2) + ": " + e.message);
      }
    }
    return JSON.stringify({ ok: true, inseridos: inseridos, atualizados: atualizados, erros: erros });
  } catch(e) {
    return JSON.stringify({ ok: false, erro: e.message });
  }
}
