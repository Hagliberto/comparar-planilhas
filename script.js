// script.js
// Carrega e converte múltiplos arquivos XLSX em arrays de objetos
async function readSheets(files) {
    const data = [];
    for (let f of files) {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      data.push(...json);
    }
    return data;
  }
  
  // Renderiza um array de objetos em tabela HTML dentro de um container
  function renderTable(container, data, title) {
    const section = document.createElement('section');
    section.innerHTML = `<h2>${title} (${data.length} linhas)</h2>`;
    if (!data.length) {
      container.appendChild(section);
      return;
    }
    const table = document.createElement('table');
    // Cabeçalho
    const thead = table.createTHead();
    const headerRow = thead.insertRow();
    Object.keys(data[0]).forEach(col => {
      const th = document.createElement('th');
      th.textContent = col;
      headerRow.appendChild(th);
    });
    // Corpo
    const tbody = table.createTBody();
    data.forEach(row => {
      const tr = tbody.insertRow();
      Object.keys(row).forEach(col => {
        const td = tr.insertCell();
        const val = row[col];
        td.textContent = Array.isArray(val)
          ? `${val[0]} → ${val[1]}`
          : val;
      });
    });
    section.appendChild(table);
    container.appendChild(section);
  }
  
  // Função de comparação com tolerância numérica e checagem de strings
  function compareData(nucleo, unidades) {
    const diffs = [];
    const mapUni = unidades.reduce((m, row) => {
      m[row.Matricula] = row;
      return m;
    }, {});
    for (let n of nucleo) {
      const u = mapUni[n.Matricula] || {};
      const diff = { Matricula: n.Matricula };
      let hasDiff = false;
      Object.keys(n).forEach(key => {
        const aVal = n[key];
        const bVal = u[key];
        const a = parseFloat(aVal) || 0;
        const b = parseFloat(bVal) || 0;
        if (!isNaN(a) && !isNaN(b) && (Math.abs(a - b) > 0.01)) {
          diff[key] = [a, b];
          hasDiff = true;
        } else if (isNaN(a) || isNaN(b)) {
          if ((aVal || '').toString().trim() !== (bVal || '').toString().trim()) {
            diff[key] = [aVal, bVal];
            hasDiff = true;
          }
        }
      });
      if (hasDiff) diffs.push(diff);
    }
    return diffs;
  }
  
  // Evento de clique para processar e exibir dados + gerar XLSX com 3 abas
  document.getElementById('btnCompare').onclick = async () => {
    const nucleoFiles = document.getElementById('nucleo').files;
    const uniFiles    = document.getElementById('unidades').files;
    const nucleoData  = await readSheets(nucleoFiles);
    const uniData     = await readSheets(uniFiles);
    const diffs       = compareData(nucleoData, uniData);
  
    // Limpa e renderiza as três tabelas
    const div = document.getElementById('resultado');
    div.innerHTML = '';
    renderTable(div, nucleoData, 'Dados do Núcleo');
    renderTable(div, uniData,    'Dados das Unidades');
    renderTable(div, diffs,      'Diferenças Encontradas');
  
    // Cria planilha XLSX com 3 abas usando ExcelJS
    const wb = new ExcelJS.Workbook();
  
    // Aba Núcleo
    const wsN = wb.addWorksheet('Núcleo');
    if (nucleoData.length) {
      wsN.columns = Object.keys(nucleoData[0]).map(h => ({ header: h, key: h }));
      nucleoData.forEach(r => wsN.addRow(r));
    }
  
    // Aba Unidades
    const wsU = wb.addWorksheet('Unidades');
    if (uniData.length) {
      wsU.columns = Object.keys(uniData[0]).map(h => ({ header: h, key: h }));
      uniData.forEach(r => wsU.addRow(r));
    }
  
    // Aba Diferenças
    const wsD = wb.addWorksheet('Diferenças');
    if (diffs.length) {
      wsD.columns = Object.keys(diffs[0]).map(h => ({ header: h, key: h }));
      diffs.forEach(r => wsD.addRow(r));
    }
  
    // Exporta e dispara o download
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = 'comparacao_verbas.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  };
  