let headerHtml = '';
const headerStatus = document.getElementById('headerStatus');
const headerDocxInput = document.getElementById('headerDocxInput');
const exportFullDocxBtn = document.getElementById('exportFullDocxBtn');

headerDocxInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(event) {
        mammoth.convertToHtml({ arrayBuffer: event.target.result })
        .then(result => {
            headerHtml = result.value;
            headerStatus.textContent = "Шаблон завантажено!";
            headerStatus.style.color = "green";
        }).catch(() => {
            headerStatus.textContent = "Помилка при читанні шаблону";
            headerStatus.style.color = "red";
        });
    };
    reader.readAsArrayBuffer(file);
});

exportFullDocxBtn.addEventListener('click', exportFullDocx);

function exportFullDocx() {
    if (!headerHtml) {
        alert("Завантажте шаблон DOCX для шапки!");
        return;
    }
    if (!window.htmlDocx || !window.htmlDocx.asBlob) {
        alert("Бібліотека html-docx.js не завантажилась!");
        return;
    }
    const fio = document.getElementById('fioInput').value.trim() || '________________';
    let finalHeaderHtml = headerHtml.replace(/\{FIO\}/g, fio);

    let table = document.createElement('table');
    table.style.width = "100%";
    table.style.borderCollapse = "collapse";
    table.style.marginTop = "20px";
    let thead = document.createElement('thead');
    let headRow = document.createElement('tr');
    ["№ п/п", "Назва документу", "Дата включення", "К-ть аркушів", "Примітка"].forEach(text => {
        let th = document.createElement('th');
        th.textContent = text;
        th.style.border = "1px solid #000";
        th.style.padding = "6px";
        th.style.background = "#f2f2f2";
        headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);

    let tbody = document.createElement('tbody');
    document.querySelectorAll('#docBody tr').forEach((row, idx) => {
        let tr = document.createElement('tr');
        for(let i=0;i<5;i++) {
            let td = document.createElement('td');
            if(i === 0) td.textContent = idx+1;
            else if(i === 1) td.textContent = row.children[1].textContent;
            else if(i === 2) td.textContent = formatDate(row.children[2].querySelector('input').value);
            else if(i === 3) td.textContent = row.children[3].querySelector('input').value;
            else if(i === 4) td.textContent = row.children[4].textContent;
            td.style.border = "1px solid #000";
            td.style.padding = "6px";
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    let fullHtml = `<html><head><meta charset="utf-8"></head><body>
      ${finalHeaderHtml}<br/><br/>
      ${table.outerHTML}
      </body></html>`;

    let converted = window.htmlDocx.asBlob(fullHtml, {orientation: 'portrait'});
    let link = document.createElement('a');
    link.href = URL.createObjectURL(converted);
    link.download = 'export.docx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ====== Остальной функционал ======

let uniqueDocuments = JSON.parse(localStorage.getItem('documents')) || [
  "Заява про надання соціальних послуг ХХХ",
  "Акт оцінки потреб сім’ї ХХХ",
  "Висновок оцінки потреб сім’ї ХХХ",
  "Копія свідоцтва про народження дитини ХХХ",
  "Копія картки платника податків ХХХ",
  "Копія паспорту ХХХ",
  "Копія карти платника податків YYY",
  "Копія паспорту YYY",
  "Декларація про доходи та майновий стан особи, яка потребує надання соціальних послуг ХХХ",
  "Рішення про надання соціальних послуг",
  "Рішення про припинення надання соціальних послуг",
  "Повідомлення про сім’ю яка перебуває в складних життєвих обставинах №425/01-17",
  "Копія свідоцтва про народження Кічук А.В.",
  "Копія картки платника податків Кічук О.С.",
  "Копія довідки про внесення відомостей до єдиного державного демографічного реєстру. №1303026-2023",
  "Копія ID Картки Кічук О.С.",
  "Копія рішення Виконавчого Комітету Таращанської міської ради №305-VIII"
];
let tableDocuments = JSON.parse(localStorage.getItem('tableDocuments')) || [];

const today = new Date().toISOString().split('T')[0];
const docBody = document.getElementById('docBody');
const docList = document.getElementById('docList');
const editModal = document.getElementById('editModal');
let currentEditIndex = null;
let currentEditSource = null;

document.getElementById('replaceBtn').addEventListener('click', applyReplace);
document.getElementById('replaceDateBtn').addEventListener('click', applyDateReplace);
document.getElementById('copyTableBtn').addEventListener('click', copyTable);
document.getElementById('exportTableBtn').addEventListener('click', exportTable);
document.getElementById('addNewDocumentBtn').addEventListener('click', addNewDocument);
document.getElementById('searchDoc').addEventListener('input', searchDocuments);
document.getElementById('saveEditedDocumentBtn').addEventListener('click', saveEditedDocument);
document.getElementById('closeModalBtn').addEventListener('click', closeModal);

document.getElementById('tableContainer').addEventListener('dragover', function(e) { e.preventDefault(); });
document.getElementById('tableContainer').addEventListener('drop', handleDrop);
document.getElementById('listContainer').addEventListener('dragover', function(e) { e.preventDefault(); });
document.getElementById('listContainer').addEventListener('drop', handleRemove);

window.addEventListener('DOMContentLoaded', () => {
  renderList();
  renderTable();
});

function saveDocuments() {
  localStorage.setItem('documents', JSON.stringify(uniqueDocuments));
  localStorage.setItem('tableDocuments', JSON.stringify(getTableData()));
}

function getTableData() {
  const rows = [...docBody.querySelectorAll('tr')];
  return rows.map(row => {
    return {
      name: row.children[1].textContent,
      date: row.children[2].querySelector('input').value,
      sheets: row.children[3].querySelector('input').value,
      note: row.children[4].textContent
    };
  });
}

function renderList(filteredDocs = uniqueDocuments) {
  docList.innerHTML = '';
  filteredDocs.forEach((doc, index) => {
    const item = document.createElement('div');
    item.classList.add('draggable-item');
    item.textContent = doc;
    item.draggable = true;
    item.dataset.docIndex = index;
    item.addEventListener('dragstart', (e) => {
      e.dataTransfer.setData('text/plain', doc);
      e.dataTransfer.setData('source', 'list');
    });

    const btnContainer = document.createElement('div');
    const editBtn = document.createElement('button');
    editBtn.textContent = '✎';
    editBtn.className = 'edit-btn';
    editBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      currentEditIndex = index;
      currentEditSource = 'list';
      document.getElementById('editDocName').value = doc;
      editModal.style.display = 'block';
    });
    const delBtn = document.createElement('button');
    delBtn.textContent = '✖';
    delBtn.className = 'delete-btn';
    delBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      uniqueDocuments.splice(index, 1);
      saveDocuments();
      renderList();
    });
    btnContainer.appendChild(editBtn);
    btnContainer.appendChild(delBtn);
    item.appendChild(btnContainer);
    docList.appendChild(item);
  });
}

function renderTable() {
  docBody.innerHTML = '';
  if (tableDocuments.length > 0) {
    tableDocuments.forEach(doc => {
      addDocumentToTable(doc.name, doc.date, doc.sheets, doc.note);
    });
  }
  updateIndexes();
}

function addDocumentToTable(name, date = today, sheets = '1', note = '') {
  const row = document.createElement('tr');
  row.draggable = true;
  row.addEventListener('dragstart', (e) => {
    e.dataTransfer.setData('text/plain', name);
    e.dataTransfer.setData('source', 'table');
    row.classList.add('dragging');
  });
  row.addEventListener('dragend', () => row.classList.remove('dragging'));

  const editBtn = document.createElement('button');
  editBtn.textContent = '✎';
  editBtn.className = 'edit-btn';
  editBtn.addEventListener('click', function() {
    currentEditIndex = [...docBody.querySelectorAll('tr')].indexOf(row);
    currentEditSource = 'table';
    document.getElementById('editDocName').value = row.children[1].textContent;
    editModal.style.display = 'block';
  });

  const delBtn = document.createElement('button');
  delBtn.textContent = '✖';
  delBtn.className = 'delete-btn';
  delBtn.addEventListener('click', function() {
    row.remove();
    updateIndexes();
    saveDocuments();
  });

  const actionTd = document.createElement('td');
  actionTd.appendChild(editBtn);
  actionTd.appendChild(delBtn);

  row.innerHTML = `
    <td></td>
    <td>${name}</td>
    <td><input type="date" value="${date}"/></td>
    <td><input type="number" value="${sheets}" min="1" style="width: 100%"/></td>
    <td>${note}</td>
  `;
  row.appendChild(actionTd);
  docBody.appendChild(row);
  updateIndexes();
}

function saveEditedDocument() {
  const newName = document.getElementById('editDocName').value.trim();
  if (newName) {
    if (currentEditSource === 'list') {
      uniqueDocuments[currentEditIndex] = newName;
      renderList();
    } else if (currentEditSource === 'table') {
      const rows = [...docBody.querySelectorAll('tr')];
      rows[currentEditIndex].children[1].textContent = newName;
    }
    saveDocuments();
    closeModal();
  }
}

function closeModal() {
  editModal.style.display = 'none';
}

function handleDrop(e) {
  e.preventDefault();
  const doc = e.dataTransfer.getData('text/plain');
  const source = e.dataTransfer.getData('source');

  if (source === 'table') {
    const dragged = [...docBody.querySelectorAll('tr')].find(r => r.classList.contains('dragging'));
    const dropTarget = e.target.closest('tr');
    if (dragged && dropTarget && dragged !== dropTarget) {
      docBody.insertBefore(dragged, dropTarget.nextSibling);
    }
    updateIndexes();
    saveDocuments();
    return;
  }

  addDocumentToTable(doc);
  saveDocuments();
}

function handleRemove(e) {
  e.preventDefault();
  const doc = e.dataTransfer.getData('text/plain');
  const source = e.dataTransfer.getData('source');
  if (source === 'table') {
    const rows = [...docBody.querySelectorAll('tr')];
    const rowToRemove = rows.find(r => r.children[1].textContent === doc);
    if (rowToRemove) rowToRemove.remove();
    updateIndexes();
    saveDocuments();
  }
}

function updateIndexes() {
  document.querySelectorAll('#docBody tr').forEach((row, index) => {
    row.children[0].textContent = index + 1;
  });
}

function addNewDocument() {
  const newDoc = document.getElementById('newDoc').value.trim();
  if (newDoc) {
    uniqueDocuments.push(newDoc);
    saveDocuments();
    renderList();
    document.getElementById('newDoc').value = '';
  }
}

function applyReplace() {
  const replaceFrom = document.getElementById('replaceValueFrom').value.trim();
  const replaceTo = document.getElementById('replaceValueTo').value.trim();
  if (!replaceFrom || !replaceTo) return;
  const rows = [...docBody.querySelectorAll('tr')];
  rows.forEach(row => {
    const nameCell = row.children[1];
    nameCell.textContent = nameCell.textContent.replace(new RegExp(replaceFrom, 'g'), replaceTo);
  });
  saveDocuments();
}

function applyDateReplace() {
  const replaceDate = document.getElementById('replaceDateValue').value.trim();
  if (!replaceDate) return;
  const rows = [...docBody.querySelectorAll('tr')];
  rows.forEach(row => {
    const dateCell = row.children[2].querySelector('input');
    dateCell.value = replaceDate;
  });
  saveDocuments();
}

function searchDocuments() {
  const searchValue = document.getElementById('searchDoc').value.toLowerCase();
  const filteredDocs = uniqueDocuments.filter(doc => doc.toLowerCase().includes(searchValue));
  renderList(filteredDocs);
}

function formatDate(dateString) {
  if (!dateString) return '';
  const date = new Date(dateString);
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

function copyTable() {
  const replaceValue = document.getElementById('replaceValueFrom').value.trim();
  const rows = [...docBody.querySelectorAll('tr')];

  let result = '';
  rows.forEach((row, index) => {
    const name = row.children[1].textContent
                      .replace(new RegExp(replaceValue, 'g'), document.getElementById('replaceValueTo').value.trim());
    const date = formatDate(row.children[2].querySelector('input').value);
    const sheets = row.children[3].querySelector('input').value;

    result += `${index + 1}.\n${name}\n${date}\n${sheets}\n\n\n`;
  });

  navigator.clipboard.writeText(result).then(() => {
    alert("Таблицю скопійовано у потрібному форматі!");
  });
}

function exportTable() {
  const format = document.querySelector('input[name="exportFormat"]:checked').value;
  const replaceFrom = document.getElementById('replaceValueFrom').value.trim();
  const replaceTo = document.getElementById('replaceValueTo').value.trim();
  const rows = [...docBody.querySelectorAll('tr')];

  let content = '';
  let filename = 'documents_export';

  if (format === 'txt') {
    filename += '.txt';
    rows.forEach((row, index) => {
      const name = row.children[1].textContent.replace(new RegExp(replaceFrom, 'g'), replaceTo);
      const date = formatDate(row.children[2].querySelector('input').value);
      const sheets = row.children[3].querySelector('input').value;
      const note = row.children[4].textContent;

      content += `${index + 1}.\n${name}\n${date}\n${sheets}\n${note}\n\n\n`;
    });
  } 
  else if (format === 'csv') {
    filename += '.csv';
    content = "№ п/п;Назва документу;Дата включення;К-ть аркушів;Примітка\n";
    rows.forEach((row, index) => {
      const name = row.children[1].textContent.replace(new RegExp(replaceFrom, 'g'), replaceTo);
      const date = formatDate(row.children[2].querySelector('input').value);
      const sheets = row.children[3].querySelector('input').value;
      const note = row.children[4].textContent;

      content += `${index + 1};"${name}";${date};${sheets};"${note}"\n`;
    });
  }
  else if (format === 'custom') {
    filename += '.txt';
    const template = document.getElementById('customFormat').value;
    rows.forEach((row, index) => {
      const name = row.children[1].textContent.replace(new RegExp(replaceFrom, 'g'), replaceTo);
      const date = formatDate(row.children[2].querySelector('input').value);
      const sheets = row.children[3].querySelector('input').value;
      const note = row.children[4].textContent;

      let rowContent = template
        .replace(/{index}/g, index + 1)
        .replace(/{name}/g, name)
        .replace(/{date}/g, date)
        .replace(/{sheets}/g, sheets)
        .replace(/{note}/g, note);
      content += rowContent;
    });
  }

  const blob = new Blob([content], { type: format === 'csv' ? 'text/csv;charset=utf-8;' : 'text/plain;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);

  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  link.style.visibility = 'hidden';

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

window.onclick = function(event) {
  if (event.target == editModal) {
    closeModal();
  }
};
