/**
 * កូដដំណើរការមុខងារ UI ប្រព័ន្ធគ្រប់គ្រងឯកសារស្តង់ដាសាលាបឋមសិក្សាគំរូ (app.js)
 * គ្រប់គ្រងការជ្រើសរើសស្តង់ដា ស្វែងរក មើលគំរូឯកសារ កែសម្រួល បោះពុម្ព និងទាញយកជា Word
 */

let currentStandardId = 1;
let currentActiveDoc = null;
let currentEvidenceSubCode = null;
let currentEvidenceList = [];
let currentPrintOrientation = 'portrait';

document.addEventListener('DOMContentLoaded', () => {
  initEvidenceDB();
  initStandardButtons();
  renderStandardView(currentStandardId);
  initSearch();
  initModalActions();
  initEvidenceModalActions();
});

/**
 * បង្កើតប៊ូតុងស្តង់ដាទាំង ៥ និងកំណត់ Event Listener
 */
function initStandardButtons() {
  const container = document.getElementById('standardButtonsContainer');
  if (!container) return;

  container.innerHTML = '';

  STANDARDS_DATA.forEach(std => {
    // គណនាចំនួនសូចនាករគន្លឹះ និងសូចនាកររង
    const keyCount = std.keyIndicators.length;
    let subCount = 0;
    std.keyIndicators.forEach(k => subCount += k.subIndicators.length);

    const btn = document.createElement('button');
    btn.className = `standard-btn ${std.themeClass} ${std.id === currentStandardId ? 'active' : ''}`;
    btn.dataset.id = std.id;
    btn.onclick = () => selectStandard(std.id);

    btn.innerHTML = `
      <div class="std-top-row">
        <div class="std-icon-box">
          <i class="fas ${std.icon}"></i>
        </div>
        <span class="std-badge">ស្តង់ដា ${std.code}</span>
      </div>
      <div>
        <div class="std-title">${std.shortTitle}</div>
        <div class="std-subtitle">${std.description.substring(0, 50)}...</div>
      </div>
      <div class="std-count-info">
        <span><i class="fas fa-layer-group"></i> ${khmerNumber(keyCount)} សូចនាករគន្លឹះ</span>
        <span><i class="fas fa-file-alt"></i> <strong>${khmerNumber(subCount)}</strong> សូចនាកររង</span>
      </div>
    `;

    container.appendChild(btn);
  });
}

/**
 * ជ្រើសរើសស្តង់ដាណាមួយ
 */
function selectStandard(stdId) {
  currentStandardId = parseInt(stdId);

  // Update active button UI
  document.querySelectorAll('.standard-btn').forEach(btn => {
    if (parseInt(btn.dataset.id) === currentStandardId) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });

  // Clear search input if user switches standard
  const searchInput = document.getElementById('searchInput');
  if (searchInput) searchInput.value = '';

  renderStandardView(currentStandardId);
}

/**
 * បង្ហាញបញ្ជីសូចនាករ និងទម្រង់ឯកសារក្នុងស្តង់ដាដែលបានជ្រើស
 */
function renderStandardView(stdId) {
  const container = document.getElementById('mainContentArea');
  if (!container) return;

  const std = getStandardById(stdId);
  if (!std) return;

  // គណនាសូចនាកររងសរុបក្នុងស្តង់ដានេះ
  let totalSubInStd = 0;
  std.keyIndicators.forEach(k => totalSubInStd += k.subIndicators.length);

  let html = `
    <div class="section-banner">
      <div class="section-title-wrap">
        <div class="section-icon" style="background: ${std.color};">
          <i class="fas ${std.icon}"></i>
        </div>
        <div class="section-info">
          <h2>${std.title}</h2>
          <p>${std.description}</p>
        </div>
      </div>
      <div class="header-badges">
        <span class="stat-pill">
          <i class="fas fa-folder-open"></i> សូចនាករគន្លឹះ៖ <strong>${khmerNumber(std.keyIndicators.length)}</strong>
        </span>
        <span class="stat-pill">
          <i class="fas fa-file-contract"></i> សូចនាកររង / ឯកសារ៖ <strong>${khmerNumber(totalSubInStd)}</strong>
        </span>
      </div>
    </div>
  `;

  // Render each key indicator group
  std.keyIndicators.forEach(keyInd => {
    html += `
      <div class="key-indicator-group">
        <div class="key-indicator-header">
          <h3>
            <i class="fas fa-check-circle" style="color: ${std.color};"></i>
            ${keyInd.title}
          </h3>
          <span style="font-size: 0.82rem; color: #64748b; font-weight: 600;">
            ${khmerNumber(keyInd.subIndicators.length)} សូចនាកររង
          </span>
        </div>
        <div class="sub-indicators-grid">
    `;

    keyInd.subIndicators.forEach(sub => {
      const isCustomized = !!localStorage.getItem('saved_doc_' + sub.code);
      const evCount = getEvidenceCount(sub.code);
      const evBadgeHtml = evCount > 0 ? `<span class="evidence-badge-num">${khmerNumber(evCount)}</span>` : '';
      html += `
        <div class="sub-indicator-card" id="card-${sub.code.replace(/\./g, '-')}">
          <div>
            <div class="card-header-row">
              <span class="card-code-badge">
                <i class="fas fa-tag"></i> ${sub.code}
              </span>
              <div style="display: flex; gap: 0.4rem; align-items: center;">
                ${isCustomized ? '<span class="card-edited-badge" title="មានទិន្នន័យបានកែប្រែ និងរក្សាទុក"><i class="fas fa-check-circle"></i> បានកែប្រែ</span>' : ''}
                ${evCount > 0 ? `<span class="card-evidence-badge" title="មានភស្តុតាងភ្ជាប់ចំនួន ${khmerNumber(evCount)}"><i class="fas fa-images"></i> ភស្តុតាង (${khmerNumber(evCount)})</span>` : ''}
                <span class="card-type-tag">${sub.category || 'ឯកសារកូឡោនទី២'}</span>
              </div>
            </div>
            <h4 class="card-title">${sub.name}</h4>
            <div class="card-docs-preview">
              <div class="doc-item">
                <i class="fas fa-file-invoice"></i>
                <div>
                  <strong>ឯកសារតម្រូវ (កូឡោនទី២)៖</strong><br/>
                  <span>${sub.col2Docs}</span>
                </div>
              </div>
            </div>
          </div>
          <div class="card-actions">
            <button class="btn-card-action btn-view-doc" onclick="openDocumentModal('${sub.code}')" title="មើលគំរូទម្រង់ឯកសារ">
              <i class="fas fa-eye"></i> មើលគំរូ
            </button>
            <button class="btn-card-action btn-import-excel" onclick="openExcelImportModal('${sub.code}')" title="នាំចូលទិន្នន័យពី Excel ឬ Google Sheet បង្កើតរបាយការណ៍ស្វ័យប្រវត្តិ">
              <i class="fas fa-file-excel"></i> នាំចូល Excel
            </button>
            <button class="btn-card-action btn-evidence ${evCount > 0 ? 'has-evidence' : ''}" onclick="openEvidenceModal('${sub.code}')" title="បញ្ចូល និងគ្រប់គ្រងភស្តុតាងបញ្ជាក់ (រូបភាព/ឯកសារ)">
              <i class="fas fa-images"></i> ភស្តុតាង ${evBadgeHtml}
            </button>
            <button class="btn-card-action btn-print-doc" onclick="quickPrintDocument('${sub.code}')" title="បោះពុម្ព A4 ភ្លាមៗ">
              <i class="fas fa-print"></i> បោះពុម្ព
            </button>
          </div>
        </div>
      `;
    });

    html += `
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

/**
 * មុខងារស្វែងរកសូចនាករ (Search & Filter)
 */
function initSearch() {
  const input = document.getElementById('searchInput');
  if (!input) return;

  input.addEventListener('input', (e) => {
    const query = e.target.value.trim().toLowerCase();
    if (!query) {
      renderStandardView(currentStandardId);
      return;
    }

    const allSubs = getAllSubIndicators();
    const matched = allSubs.filter(sub => {
      return sub.code.includes(query) ||
             sub.name.toLowerCase().includes(query) ||
             sub.col2Docs.toLowerCase().includes(query) ||
             sub.templateTitle.toLowerCase().includes(query);
    });

    renderSearchResults(matched, query);
  });
}

function renderSearchResults(results, query) {
  const container = document.getElementById('mainContentArea');
  if (!container) return;

  if (results.length === 0) {
    container.innerHTML = `
      <div style="background: white; padding: 3rem; text-align: center; border-radius: 16px; border: 1px solid #e2e8f0; margin-top: 1rem;">
        <i class="fas fa-search" style="font-size: 2.5rem; color: #cbd5e1; margin-bottom: 1rem;"></i>
        <h3 style="font-size: 1.2rem; color: #334155; margin-bottom: 0.5rem;">រកមិនឃើញសូចនាករដែលត្រូវនឹង «${query}» ឡើយ</h3>
        <p style="color: #64748b; font-size: 0.9rem;">សូមសាកល្បងបញ្ចូលលេខកូដសូចនាករ ដូចជា «១.១.២» ឬពាក្យគន្លឹះ «អំណាន», «កិច្ចព្រមព្រៀង»...</p>
      </div>
    `;
    return;
  }

  let html = `
    <div class="section-banner" style="border-left: 4px solid #0284c7;">
      <div class="section-title-wrap">
        <div class="section-icon" style="background: #0284c7;">
          <i class="fas fa-search"></i>
        </div>
        <div class="section-info">
          <h2>លទ្ធផលស្វែងរក៖ «${query}»</h2>
          <p>រកឃើញសរុបចំនួន ${khmerNumber(results.length)} សូចនាកររង</p>
        </div>
      </div>
      <button class="btn-secondary" onclick="clearSearch()">
        <i class="fas fa-times"></i> បោះបង់ការស្វែងរក
      </button>
    </div>
    <div class="sub-indicators-grid">
  `;

  results.forEach(sub => {
    const isCustomized = !!localStorage.getItem('saved_doc_' + sub.code);
    const evCount = getEvidenceCount(sub.code);
    const evBadgeHtml = evCount > 0 ? `<span class="evidence-badge-num">${khmerNumber(evCount)}</span>` : '';
    html += `
      <div class="sub-indicator-card" id="card-${sub.code.replace(/\./g, '-')}">
        <div>
          <div class="card-header-row">
            <span class="card-code-badge">
              <i class="fas fa-tag"></i> ${sub.code}
            </span>
            <div style="display: flex; gap: 0.4rem; align-items: center;">
              ${isCustomized ? '<span class="card-edited-badge" title="មានទិន្នន័យបានកែប្រែ និងរក្សាទុក"><i class="fas fa-check-circle"></i> បានកែប្រែ</span>' : ''}
              ${evCount > 0 ? `<span class="card-evidence-badge" title="មានភស្តុតាងភ្ជាប់ចំនួន ${khmerNumber(evCount)}"><i class="fas fa-images"></i> ភស្តុតាង (${khmerNumber(evCount)})</span>` : ''}
              <span class="card-type-tag" style="background: #e0f2fe; color: #0369a1;">
                ស្តង់ដាទី ${khmerNumber(sub.standardId)}
              </span>
            </div>
          </div>
          <h4 class="card-title">${sub.name}</h4>
          <div class="card-docs-preview">
            <div class="doc-item">
              <i class="fas fa-file-invoice"></i>
              <div>
                <strong>ឯកសារតម្រូវ (កូឡោនទី២)៖</strong><br/>
                <span>${sub.col2Docs}</span>
              </div>
            </div>
          </div>
        </div>
        <div class="card-actions">
          <button class="btn-card-action btn-view-doc" onclick="openDocumentModal('${sub.code}')" title="មើលគំរូទម្រង់ឯកសារ">
            <i class="fas fa-eye"></i> មើលគំរូ
          </button>
          <button class="btn-card-action btn-evidence ${evCount > 0 ? 'has-evidence' : ''}" onclick="openEvidenceModal('${sub.code}')" title="បញ្ចូល និងគ្រប់គ្រងភស្តុតាងបញ្ជាក់ (រូបភាព/ឯកសារ)">
            <i class="fas fa-images"></i> ភស្តុតាង ${evBadgeHtml}
          </button>
          <button class="btn-card-action btn-print-doc" onclick="quickPrintDocument('${sub.code}')" title="បោះពុម្ព A4 ភ្លាមៗ">
            <i class="fas fa-print"></i> បោះពុម្ព
          </button>
        </div>
      </div>
    `;
  });

  html += `</div>`;
  container.innerHTML = html;
}

function clearSearch() {
  const input = document.getElementById('searchInput');
  if (input) input.value = '';
  renderStandardView(currentStandardId);
}

let hasUnsavedChanges = false;

/**
 * បង្កើត និងបង្ហាញផ្ទាំងទម្រង់ឯកសារផ្លូវការ (Document Modal)
 * គាំទ្រការកែប្រែទិន្នន័យដោយផ្ទាល់ និងរក្សាទុកក្នុង LocalStorage
 */
function openDocumentModal(subCode) {
  const sub = getSubIndicatorByCode(subCode);
  if (!sub) return;

  currentActiveDoc = sub;
  hasUnsavedChanges = false;
  const modal = document.getElementById('documentModal');
  const modalContainer = document.getElementById('documentSheetContainer');
  const modalTitle = document.getElementById('modalDocTitle');
  const saveBtn = document.getElementById('btnSaveDocument');

  if (modalTitle) {
    modalTitle.innerHTML = `<span style="color: #fde047;">[លេខកូដ៖ ${sub.code}]</span> ${sub.templateTitle}`;
  }

  // Reset Save button appearance
  if (saveBtn) {
    saveBtn.classList.remove('has-unsaved');
    saveBtn.innerHTML = '<i class="fas fa-save"></i> <span>រក្សាទុក</span>';
  }

  // ពិនិត្យមើលថាតើមានទិន្នន័យដែលបានកែប្រែ និងរក្សាទុកពីមុនមកឬទេ
  let savedHtml = localStorage.getItem('saved_doc_' + subCode);
  const savedTime = localStorage.getItem('saved_doc_time_' + subCode);

  // ធ្វើបច្ចុប្បន្នកម្មស្វ័យប្រវត្តិ ប្រសិនបើទិន្នន័យចាស់ក្នុង LocalStorage មានព័ត៌មានខេត្តកណ្តាល/លេខកូដចាស់
  if (savedHtml) {
    savedHtml = savedHtml.replace(/មន្ទីរអប់រំ យុវជន និងកីឡា ខេត្តកណ្តាល/g, 'មន្ទីរអប់រំ យុវជន និងកីឡា ខេត្តកែប')
                         .replace(/ការិយាល័យអប់រំ យុវជន និងកីឡា ស្រុកកណ្តាលស្ទឹង/g, 'ការិយាល័យអប់រំ យុវជន និងកីឡា នៃរដ្ឋបាលស្រុកដំណាក់ចង្អើរ')
                         .replace(/០៨០២០១០៤/g, '22020204012')
                         .replace(/08020104/g, '22020204012')
                         .replace(/គង់ សុផល<\/div>/g, '................................................</div>')
                         .replace(/<div class="doc-meta-item">\s*<strong>\s*សាលាបឋមសិក្សា[៖:]\s*<\/strong>[\s\S]*?<\/div>\s*/g, '');
    localStorage.setItem('saved_doc_' + subCode, savedHtml);
  }
  const hasSavedData = !!savedHtml;

  // Edit Mode Banner (Hidden on Print)
  const bannerHtml = `
    <div class="doc-edit-banner no-print">
      <div class="banner-text">
        <i class="fas fa-edit"></i>
        <span>ទម្រង់កែសម្រួលផ្ទាល់៖ លោកអ្នកអាចចុចលើទិន្នន័យក្នុងតារាង ឬអត្ថបទនានាដើម្បីកែប្រែ រួចចុចប៊ូតុង <strong>«រក្សាទុក»</strong> នៅខាងលើ។</span>
      </div>
      <div id="modalSaveStatus" class="doc-edit-status-badge ${hasSavedData ? 'status-saved' : 'status-default'}">
        <i class="fas ${hasSavedData ? 'fa-check-circle' : 'fa-circle-info'}"></i>
        <span>${hasSavedData ? 'បានរក្សាទុក ' + (savedTime ? '(' + savedTime + ')' : '') : 'ទម្រង់គំរូលំនាំដើម'}</span>
      </div>
    </div>
  `;

  // បើមានទិន្នន័យដែលបានកែប្រែ និងរក្សាទុក សូមយកទិន្នន័យនោះមកបង្ហាញ បើគ្មានទេ បង្ហាញទម្រង់ដើម
  let cleanMainHtml = hasSavedData ? savedHtml : buildDefaultDocumentHtml(sub);
  
  // ធ្វើបច្ចុប្បន្នកម្មតួនាទីហត្ថលេខាចាស់ និងដកប្រអប់ឈ្មោះសាលាដែលស្ទួនចេញ ប្រសិនបើមានក្នុងទិន្នន័យដែលធ្លាប់រក្សាទុក
  cleanMainHtml = cleanMainHtml
    .replace(/<div class="doc-meta-item">\s*<strong>\s*សាលាបឋមសិក្សា[៖:]\s*<\/strong>[\s\S]*?<\/div>\s*/g, '')
    .replace(/ប្រធានគណៈកម្មការ\s*គ\.ក\.ស\./g, 'ប្រធាន គ.គ.ស')
    .replace(/នាយកសាលាបឋមសិក្សាតាពីង/g, 'នាយកសាលា');
  
  // សម្អាត Divider ឬ Annex ចាស់ចេញពី Main Sheet ប្រសិនបើធ្លាប់បានផ្ទុក
  const sepIdx = cleanMainHtml.indexOf('class="evidence-page-separator');
  if (sepIdx !== -1) {
    cleanMainHtml = cleanMainHtml.substring(0, sepIdx);
  }
  const page2Idx = cleanMainHtml.indexOf('class="document-page-2');
  if (page2Idx !== -1) {
    cleanMainHtml = cleanMainHtml.substring(0, page2Idx);
  }
  // ដក Tag wrapper ចាស់ៗចេញប្រសិនបើមាន
  if (cleanMainHtml.includes('id="docFirstPage"')) {
    cleanMainHtml = cleanMainHtml.replace(/<div class="doc-first-page" id="docFirstPage">/g, '').replace(/<\/div>\s*$/g, '');
  }
  if (cleanMainHtml.includes('id="docPagesContainer"')) {
    cleanMainHtml = cleanMainHtml.replace(/<div class="doc-pages-container" id="docPagesContainer">/g, '').replace(/<\/div>\s*$/g, '');
  }

  const evList = getEvidenceList(subCode).filter(item => item && (item.imageData || hasLinkUrl(item) || (item.caption && item.caption.trim())));
  const hasEvidence = evList.length > 0;

  // រៀបចំសន្លឹករបាយការណ៍ (Multi-page report sheets)
  let pagesContainerHtml = '';
  if (cleanMainHtml.includes('class="doc-report-page"')) {
    pagesContainerHtml = cleanMainHtml;
  } else {
    pagesContainerHtml = `
      <div class="doc-report-page" id="docPage_1" data-page-index="1">
        ${cleanMainHtml}
      </div>
    `;
  }

  const fullSheetHtml = `
    <div class="doc-pages-container" id="docPagesContainer">
      ${pagesContainerHtml}
    </div>
    <div class="add-sheet-action-wrap no-print">
      <button type="button" class="btn-add-sheet-primary" onclick="addNewDocumentSheetPage()" title="បន្ថែមសន្លឹករបាយការណ៍ថ្មី (ទំព័រទី...)">
        <i class="fas fa-file-circle-plus"></i> <span>+ បន្ថែមសន្លឹករបាយការណ៍ថ្មី (ទំព័រទី...)</span>
      </button>
    </div>
    <div class="evidence-page-separator no-print">
      <div class="separator-line"></div>
      <div class="separator-pill">
        <i class="fas fa-images"></i> ឧបសម្ព័ន្ធភស្តុតាងបញ្ជាក់បន្ថែម (ស្ថិតនៅបន្តបន្ទាប់ពីទំព័ររបាយការណ៍) ${hasEvidence ? `(${khmerNumber(evList.length)})` : ''}
      </div>
      <div class="separator-line"></div>
    </div>
    <div class="document-page-2 ${hasEvidence ? '' : 'has-no-evidence'}" id="docEvidenceAnnex">
      ${renderEvidenceAnnexHtml(sub)}
    </div>
  `;

  modalContainer.innerHTML = bannerHtml + `<div class="document-sheet" id="printableDocument">${fullSheetHtml}</div>`;

  // ពិនិត្យចំនួនជួរឈរដើម្បីកំណត់ទិសដៅក្រដាស និងទំហំសមាមាត្រកុំឱ្យដាច់ពេលបោះពុម្ព
  const docTables = modalContainer.querySelectorAll('.doc-table');
  let maxColCount = 5;
  docTables.forEach(docTable => {
    const dataHeaders = docTable.querySelectorAll('thead th:not(.row-action-col)');
    const colCount = dataHeaders.length;
    if (colCount > maxColCount) maxColCount = colCount;
    if (colCount >= 10) {
      docTable.classList.add('table-ultra-wide');
      docTable.classList.remove('table-wide');
    } else if (colCount >= 7) {
      docTable.classList.add('table-wide');
      docTable.classList.remove('table-ultra-wide');
    } else {
      docTable.classList.remove('table-wide', 'table-ultra-wide');
    }
  });

  if (maxColCount >= 7) {
    setPrintOrientation('landscape');
  } else {
    setPrintOrientation('portrait');
  }

  // កំណត់ Event Listeners សម្រាប់ការកែសម្រួលផ្ទាល់
  initTableEditingFeatures();

  modal.classList.add('active');
  document.body.style.overflow = 'hidden';
}

/**
 * បង្កើត HTML នៃរបារឧបករណ៍បញ្ជាតារាង (Table Actions Bar)
 * រៀបចំជា ២ ជួរស្មើគ្នា (Row 1: ៦ ប៊ូតុង, Row 2: ៦ ប៊ូតុង) ទំហំប៉ុនៗគ្នាស្មើបេះបិទ
 */
function getTableActionsBarHtml() {
  return `
    <div class="table-actions-grid">
      <!-- ជួរទី១ (Row 1): ៦ ប៊ូតុង - ជួរដេក ជួរឈរ និងការបញ្ចូលទិន្នន័យ -->
      <button class="btn-table-action btn-add-row" onclick="addNewTableRow()" type="button" title="បន្ថែមជួរដេកថ្មីខាងក្រោមតារាង">
        <i class="fas fa-plus-circle"></i> <span>បន្ថែមជួរដេក</span>
      </button>
      <button class="btn-table-action btn-delete-row" onclick="deleteTableRowSmart()" type="button" title="លុបជួរដេកដែលកំពុងជ្រើស ឬជួរដេកចុងក្រោយបង្អស់នៃតារាង">
        <i class="fas fa-minus-circle"></i> <span>លុបជួរដេក</span>
      </button>
      <button class="btn-table-action btn-add-col" onclick="addNewTableColumn()" type="button" title="បន្ថែមជួរឈរថ្មីខាងស្តាំតារាង">
        <i class="fas fa-columns"></i> <span>បន្ថែមជួរឈរ</span>
      </button>
      <button class="btn-table-action btn-delete-col" onclick="openDeleteColumnModal()" type="button" title="ជ្រើសរើសជួរឈរណាមួយដើម្បីលុបចេញ">
        <i class="fas fa-trash-alt"></i> <span>លុបជួរឈរ</span>
      </button>
      <button class="btn-table-action btn-paste-clipboard" onclick="pasteClipboardToTable()" type="button" title="បិទភ្ជាប់ទិន្នន័យពី Excel ឬ Google Sheet ចូលក្នុងតារាង (Ctrl+V)">
        <i class="fas fa-paste"></i> <span>បិទភ្ជាប់ (Paste)</span>
      </button>
      <button class="btn-table-action btn-import-excel" onclick="openExcelImportModal(currentActiveDoc ? currentActiveDoc.code : null)" type="button" title="នាំចូលទិន្នន័យពី Excel ឬ Google Sheet ដើម្បីបង្កើតរបាយការណ៍ស្វ័យប្រវត្តិ">
        <i class="fas fa-file-excel"></i> <span>នាំចូល Excel</span>
      </button>

      <!-- ជួរទី២ (Row 2): ៦ ប៊ូតុង - ក្រឡា (Merge/Unmerge) និងទំហំតារាង -->
      <div class="merge-dropdown-wrap">
        <button class="btn-table-action btn-merge-cells" id="btnTableMergeCells" onclick="toggleMergeDropdown(event)" type="button" title="បញ្ចូលក្រឡាដែលបានជ្រើស (Merge) ឬបញ្ចូលជាមួយក្រឡាជិតខាង (Alt+M)">
          <i class="fas fa-object-group"></i> <span>បញ្ចូលក្រឡា (Merge)</span> <i class="fas fa-caret-down" style="font-size: 0.68rem; margin-left: 2px;"></i>
        </button>
        <div class="merge-dropdown-menu" id="mergeDropdownMenu" style="display: none;">
          <button type="button" class="merge-dropdown-item" onclick="mergeSelectedTableCells(); closeMergeDropdown();">
            <i class="fas fa-object-group" style="color: #0d9488;"></i> <span>បញ្ចូលក្រឡាដែលបានជ្រើស</span> <kbd style="margin-left:auto; font-size:0.7rem; color:#94a3b8;">Alt+M</kbd>
          </button>
          <button type="button" class="merge-dropdown-item" onclick="mergeActiveCellRight(); closeMergeDropdown();">
            <i class="fas fa-arrow-right" style="color: #0284c7;"></i> <span>បញ្ចូលជាមួយក្រឡាខាងស្តាំ</span>
          </button>
          <button type="button" class="merge-dropdown-item" onclick="mergeActiveCellDown(); closeMergeDropdown();">
            <i class="fas fa-arrow-down" style="color: #2563eb;"></i> <span>បញ្ចូលជាមួយក្រឡាខាងក្រោម</span>
          </button>
          <div style="height: 1px; background: #e2e8f0; margin: 4px 0;"></div>
          <button type="button" class="merge-dropdown-item" onclick="unmergeSelectedTableCells(); closeMergeDropdown();">
            <i class="fas fa-object-ungroup" style="color: #d97706;"></i> <span>បំបែកក្រឡា (Unmerge)</span> <kbd style="margin-left:auto; font-size:0.7rem; color:#94a3b8;">Alt+Shift+M</kbd>
          </button>
        </div>
      </div>
      <button class="btn-table-action btn-unmerge-cells" onclick="unmergeSelectedTableCells()" type="button" title="បំបែកក្រឡាដែលបាន Merge ត្រឡប់មកធម្មតាវិញ (Alt+Shift+M)">
        <i class="fas fa-object-ungroup"></i> <span>បំបែកក្រឡា</span>
      </button>
      <button class="btn-table-action btn-widen-col" onclick="adjustActiveColumnWidth(25)" type="button" title="ពង្រីកទទឹងជួរឈរដែលបានជ្រើស (+25px)">
        <i class="fas fa-arrows-alt-h"></i> <span>ពង្រីកជួរឈរ (+)</span>
      </button>
      <button class="btn-table-action btn-narrow-col" onclick="adjustActiveColumnWidth(-25)" type="button" title="បង្រួមទទឹងជួរឈរដែលបានជ្រើស (-25px)">
        <i class="fas fa-compress-alt"></i> <span>បង្រួមជួរឈរ (-)</span>
      </button>
      <button class="btn-table-action btn-autofit-cols" onclick="autoFitTableColumns()" type="button" title="លៃតម្រូវទទឹងស្វ័យប្រវត្តិតាមប្រវែងអត្ថបទ (Auto-fit)">
        <i class="fas fa-wand-magic-sparkles"></i> <span>លៃស្វ័យប្រវត្តិ</span>
      </button>
      <button class="btn-table-action btn-equal-cols" onclick="equalizeTableColumns()" type="button" title="ចែករំលែកទទឹងគ្រប់ជួរឈរឱ្យមានទំហំប៉ុនៗគ្នាស្មើភាព">
        <i class="fas fa-arrows-split-up-and-left"></i> <span>ទទឹងស្មើគ្នា</span>
      </button>
    </div>
    <div class="table-actions-footer">
      <span class="table-hint">
        <i class="fas fa-keyboard"></i> ចុចលើប្រអប់ដើម្បីកែប្រែ | អូស Mouse ឬសង្កត់ Shift លើក្រឡាដើម្បី <strong>បញ្ចូលក្រឡា (Merge)</strong> | ចុចកណ្ដុរស្ដាំ (Right-Click) លើតារាងដើម្បីបើកម៉ឺនុយរហ័ស
      </span>
      <button type="button" class="btn-table-action" onclick="addNewDocumentSheetPage()" style="margin-left: auto; background: #2563eb; color: white; border-color: #1d4ed8; padding: 0.28rem 0.85rem; border-radius: 6px; font-size: 0.8rem; font-weight: 700; display: inline-flex; align-items: center; gap: 0.4rem; cursor: pointer; white-space: nowrap;" title="បន្ថែមសន្លឹករបាយការណ៍ថ្មីមួយទំព័រទៀត (រក្សាក្បាលសំបុត្រ ក្បាលតារាង និងហត្ថលេខា)">
        <i class="fas fa-file-circle-plus"></i> <span>+ បន្ថែមសន្លឹកថ្មី</span>
      </button>
    </div>
  `;
}

/**
 * បង្កើត HTML សម្រាប់ប៊ូតុងសកម្មភាពលើជួរដេក (បន្ថែមជួរថ្មី + និងលុបជួរដេក ×)
 */
function getRowActionButtonsHtml() {
  return `
    <div class="row-action-btns">
      <button class="row-insert-btn" type="button" title="បន្ថែមជួរដេកថ្មីបន្ទាប់ពីនេះ (+)"><i class="fas fa-plus"></i></button>
      <button class="row-delete-btn" type="button" title="លុបជួរដេកនេះ (×)"><i class="fas fa-times"></i></button>
    </div>
  `;
}

/**
 * បង្កើតកូដ HTML នៃទម្រង់ឯកសារគំរូលំនាំដើម
 */
function buildDefaultDocumentHtml(sub) {
  // Table Headers
  let tableHeadersHtml = '';
  sub.headers.forEach(h => {
    tableHeadersHtml += `<th class="editable-header" contenteditable="true">${h}</th>`;
  });
  tableHeadersHtml += `<th class="row-action-col no-print" style="width: 58px; border: none; background: transparent;"></th>`;

  // Table Sample Rows
  let tableRowsHtml = '';
  sub.sampleRows.forEach((row, rowIndex) => {
    tableRowsHtml += '<tr>';
    row.forEach((cell, cellIndex) => {
      const isCenter = cellIndex === 0 || isNumericData(cell);
      tableRowsHtml += `<td class="${isCenter ? 'text-center' : ''} editable-cell" contenteditable="true">${cell}</td>`;
    });
    tableRowsHtml += `<td class="row-action-col no-print">${getRowActionButtonsHtml()}</td>`;
    tableRowsHtml += '</tr>';
  });

  return `
    <!-- លេខសម្គាល់សូចនាកររងនៅផ្នែកខាងស្តាំខាងលើ -->
    <div class="indicator-badge-top-right">
      <span class="code-label">លេខសម្គាល់</span>
      <span class="code-value">${sub.code}</span>
    </div>

    <!-- ក្បាលសំបុត្រផ្លូវការ -->
    <div class="doc-header">
      <div class="doc-header-left">
        <div class="editable" contenteditable="true">ក្រសួងអប់រំ យុវជន និងកីឡា</div>
        <div class="editable" contenteditable="true">មន្ទីរអប់រំ យុវជន និងកីឡា ខេត្តកែប</div>
        <div class="editable" contenteditable="true">ការិយាល័យអប់រំ យុវជន និងកីឡា នៃរដ្ឋបាលស្រុកដំណាក់ចង្អើរ</div>
        <div class="school-name editable" contenteditable="true">សាលាបឋមសិក្សាតាពីង</div>
        <div class="editable" style="font-size: 0.8rem; color: #64748b;" contenteditable="true">លេខកូដសាលា៖ 22020204012</div>
      </div>

      <div class="doc-header-center">
        <div class="kingdom-title">ព្រះរាជាណាចក្រកម្ពុជា</div>
        <div class="motto-title">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
        <div class="motto-line"></div>
      </div>
    </div>

    <!-- ចំណងជើងឯកសារ -->
    <div class="doc-title-container">
      <h2 class="doc-main-title editable" contenteditable="true">${sub.templateTitle}</h2>
      <div class="doc-sub-title editable" contenteditable="true">សម្រាប់គាំទ្រការអនុវត្ត «ស្តង់ដាសាលាបឋមសិក្សាគំរូ» នៃក្រសួងអប់រំ យុវជន និងកីឡា</div>
      <div class="doc-standard-note">
        <strong>${sub.standardTitle}</strong> | សូចនាកររង <strong>${sub.code}</strong>៖ ${sub.name}
      </div>
    </div>

    <!-- ព័ត៌មានលម្អិតគោល -->
    <div class="doc-meta-info">
      <div class="doc-meta-item">
        <strong>ឆ្នាំសិក្សា៖</strong>
        <div class="meta-field-wrap">
          <span class="editable doc-meta-val" id="docAcademicYearText" contenteditable="true" oninput="syncAcademicYearSelect(this.innerText)">២០២៣ - ២០២៤</span>
          <select class="doc-year-select no-print" id="docAcademicYearSelect" onchange="handleAcademicYearChange(this.value)" title="ចុចជ្រើសរើសឆ្នាំសិក្សា">
            <option value="" disabled>-- ជ្រើសរើសឆ្នាំ --</option>
            <option value="២០២២ - ២០២៣">២០២២ - ២០២៣</option>
            <option value="២០២៣ - ២០២៤" selected>២០២៣ - ២០២៤</option>
            <option value="២០២៤ - ២០២៥">២០២៤ - ២០២៥</option>
            <option value="២០២៥ - ២០២៦">២០២៥ - ២០២៦</option>
            <option value="២០២៦ - ២០២៧">២០២៦ - ២០២៧</option>
            <option value="២០២៧ - ២០២៨">២០២៧ - ២០២៨</option>
            <option value="២០២៨ - ២០២៩">២០២៨ - ២០២៩</option>
            <option value="២០២៩ - ២០៣០">២០២៩ - ២០៣០</option>
            <option value="custom">✍️ បញ្ចូលឆ្នាំផ្សេង...</option>
          </select>
        </div>
      </div>
      <div class="doc-meta-item">
        <strong>កាលបរិច្ឆេទធ្វើបច្ចុប្បន្នកម្ម៖</strong>
        <div class="meta-field-wrap">
          <span class="editable doc-meta-val" id="docUpdateDateText" contenteditable="true" oninput="syncTeacherSignatureDate(this.innerText)">ថ្ងៃទី ១៥ ខែ វិច្ឆិកា ឆ្នាំ ២០២៣</span>
          <button type="button" class="btn-calendar-trigger no-print" onclick="openNativeDatePicker()" title="ចុចជ្រើសរើសកាលបរិច្ឆេទពីប្រតិទិន">
            <i class="fas fa-calendar-alt"></i> ប្រតិទិន
          </button>
          <input type="date" id="docNativeDatePicker" class="native-date-input-hidden no-print" onchange="handleDatePicked(this.value)">
        </div>
      </div>
      <div class="doc-meta-item doc-meta-full">
        <strong>ឯកសារតម្រូវ (កូឡោនទី២)៖</strong>
        <span class="editable" contenteditable="true">${sub.col2Docs}</span>
      </div>
    </div>

    <!-- តារាងទិន្នន័យឯកសារផ្លូវការ -->
    <table class="doc-table">
      <thead>
        <tr>${tableHeadersHtml}</tr>
      </thead>
      <tbody>
        ${tableRowsHtml}
      </tbody>
    </table>

    <!-- របារបញ្ជាបន្ថែម/លុបជួរដេក និងជួរឈរ (លាក់ពេលបោះពុម្ព និងពេលទាញយក) -->
    <div class="table-actions-bar no-print">
      ${getTableActionsBarHtml()}
    </div>

    <!-- កំណត់សម្គាល់បន្ថែម -->
    <div class="editable" contenteditable="true" style="margin: 1.25rem 0; font-size: 0.84rem; color: #475569; background: #f8fafc; padding: 0.75rem 1rem; border-left: 3px solid #0284c7; border-radius: 4px;">
      <strong>* សេចក្តីណែនាំក្នុងការបំពេញ៖</strong> ទម្រង់ឯកសារនេះត្រូវបានបង្កើតឡើងដោយផ្អែកលើកូឡោនទី២ (ទំព័រ១១-៣៥) នៃសៀវភៅស្តង់ដាសាលាបឋមសិក្សាគំរូ។ លោកគ្រូ-អ្នកគ្រូ និងគណៈគ្រប់គ្រងសាលាអាចចុចកែសម្រួលទិន្នន័យផ្ទាល់លើតារាងខាងលើនេះរួចចុចប៊ូតុង «រក្សាទុក» បោះពុម្ព ឬទាញយកជាឯកសារ Word ដើម្បីចុះហត្ថលេខា និងរក្សាទុកជាភស្តុតាង។
    </div>

    <!-- ហត្ថលេខា និងកាលបរិច្ឆេទផ្លូវការ -->
    <div class="doc-signatures">
      <div class="signature-block">
        <div class="signature-date editable" contenteditable="true">បានឃើញ និងអនុម័ត</div>
        <div class="signature-role editable" contenteditable="true">ប្រធាន គ.គ.ស</div>
        <div class="signature-name editable" contenteditable="true">................................................</div>
      </div>

      <div class="signature-block">
        <div class="signature-date editable" contenteditable="true">បានឃើញ និងឯកភាព</div>
        <div class="signature-role editable" contenteditable="true">នាយកសាលា</div>
        <div class="signature-name editable" contenteditable="true">................................................</div>
      </div>

      <div class="signature-block">
        <div class="signature-date editable" id="teacherSignatureDate" contenteditable="true">ថ្ងៃទី ១៥ ខែ វិច្ឆិកា ឆ្នាំ ២០២៣</div>
        <div class="signature-role editable" contenteditable="true">គ្រូបង្រៀន / មន្ត្រីទទួលបន្ទុក</div>
        <div class="signature-name editable" contenteditable="true">................................................</div>
      </div>
    </div>
  `;
}

/* ==========================================================================
   ACADEMIC YEAR DROPDOWN & CALENDAR DATE PICKER LOGIC
   ========================================================================== */

const KHMER_MONTH_NAMES = ['មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា', 'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'];

/**
 * ទទួលយកកាលបរិច្ឆេទបច្ចុប្បន្នជាទម្រង់ខ្មែរផ្លូវការ ដោយមានឈ្មោះខែជាអក្សរ (ឧ. ថ្ងៃទី ១២ ខែកញ្ញា ឆ្នាំ ២០២៦)
 */
function getKhmerCurrentDate(date = new Date()) {
  const day = khmerNumber(date.getDate());
  const month = KHMER_MONTH_NAMES[date.getMonth()];
  const year = khmerNumber(date.getFullYear());
  return `ថ្ងៃទី ${day} ខែ${month} ឆ្នាំ ${year}`;
}

/**
 * បំប្លែងលេខខែ (ឧ. «ខែ៩», «ខែ ៩», «ខែ ០៩», «ខែ9») ទៅជាឈ្មោះខែជាអក្សរខ្មែរ (ឧ. «ខែកញ្ញា»)
 */
function formatKhmerMonthNameToWord(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return '';
  const trimmed = dateStr.trim();
  if (!trimmed) return '';

  // 1. ប្រសិនបើជាទម្រង់ ISO: YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    return formatKhmerDateFromIso(trimmed);
  }

  // 2. ប្រសិនបើជាទម្រង់ DD/MM/YYYY ឬ DD-MM-YYYY
  const slashMatch = trimmed.match(/^([០-៩\d]{1,2})[\/\-]([០-៩\d]{1,2})[\/\-]([០-៩\d]{4})$/);
  if (slashMatch) {
    let d = parseInt(slashMatch[1].replace(/[០-៩]/g, c => '០១២៣៤៥៦៧៨៩'.indexOf(c)), 10);
    let m = parseInt(slashMatch[2].replace(/[០-៩]/g, c => '០១២៣៤៥៦៧៨៩'.indexOf(c)), 10);
    let y = parseInt(slashMatch[3].replace(/[០-៩]/g, c => '០១២៣៤៥៦៧៨៩'.indexOf(c)), 10);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return `ថ្ងៃទី ${khmerNumber(d < 10 ? '០' + d : d)} ខែ${KHMER_MONTH_NAMES[m - 1]} ឆ្នាំ ${khmerNumber(y)}`;
    }
  }

  // 3. ប្រសិនបើមានឈ្មោះខែជាអក្សរខ្មែររួចហើយ
  const hasKhmerMonth = KHMER_MONTH_NAMES.some(m => trimmed.includes(m));
  if (hasKhmerMonth) {
    return trimmed.replace(/ខែ\s+([ក-៝]+)/g, (match, monthWord) => {
      if (KHMER_MONTH_NAMES.includes(monthWord)) {
        return 'ខែ' + monthWord;
      }
      return match;
    });
  }

  // 4. បំប្លែង «ខែ៩», «ខែ ៩», «ខែ ០៩», «ខែ9» ទៅជា «ខែកញ្ញា»
  return trimmed.replace(/ខែ\s*([០-៩\d]{1,2})/g, (match, monthNumStr) => {
    let latinNum = monthNumStr.replace(/[០-៩]/g, d => '០១២៣៤៥៦៧៨៩'.indexOf(d));
    let m = parseInt(latinNum, 10);
    if (m >= 1 && m <= 12) {
      return 'ខែ' + KHMER_MONTH_NAMES[m - 1];
    }
    return match;
  });
}

/**
 * បំប្លែងកាលបរិច្ឆេទពី ISO (YYYY-MM-DD) ទៅជាទម្រង់ខ្មែរផ្លូវការ
 * ឧ. 2023-11-15 -> ថ្ងៃទី ១៥ ខែ វិច្ឆិកា ឆ្នាំ ២០២៣
 */
function formatKhmerDateFromIso(isoDateStr) {
  if (!isoDateStr) return '';
  const parts = isoDateStr.split('-');
  if (parts.length !== 3) return '';
  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10);
  const d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return '';
  const dayStr = khmerNumber(d < 10 ? '០' + d : d);
  const monthStr = KHMER_MONTH_NAMES[m - 1] || '';
  const yearStr = khmerNumber(y);
  return `ថ្ងៃទី ${dayStr} ខែ${monthStr} ឆ្នាំ ${yearStr}`;
}

/**
 * បំប្លែងអត្ថបទកាលបរិច្ឆេទខ្មែរទៅជាទម្រង់ ISO (YYYY-MM-DD) សម្រាប់ដាក់ក្នុង <input type="date">
 */
function parseKhmerDateToIso(khmerStr) {
  if (!khmerStr) return '';
  let latinStr = khmerStr.replace(/[០-៩]/g, d => '០១២៣៤៥៦៧៨៩'.indexOf(d));
  const dayMatch = latinStr.match(/ថ្ងៃ(?:ទី)?\s*(\d+)/);
  const yearMatch = latinStr.match(/ឆ្នាំ\s*(\d{4})/);
  let monthIdx = -1;
  KHMER_MONTH_NAMES.forEach((m, idx) => {
    if (khmerStr.includes(m)) monthIdx = idx + 1;
  });
  if (monthIdx === -1) {
    const numMonthMatch = latinStr.match(/ខែ\s*(\d+)/);
    if (numMonthMatch) {
      const parsedNum = parseInt(numMonthMatch[1], 10);
      if (parsedNum >= 1 && parsedNum <= 12) {
        monthIdx = parsedNum;
      }
    }
  }
  if (dayMatch && yearMatch && monthIdx !== -1) {
    const y = yearMatch[1];
    const m = String(monthIdx).padStart(2, '0');
    const d = String(dayMatch[1]).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return '';
}

/**
 * បើកផ្ទាំងប្រតិទិន Date Picker
 */
function openNativeDatePicker() {
  const dateInput = document.getElementById('docNativeDatePicker');
  if (!dateInput) return;

  const updateDateEl = document.getElementById('docUpdateDateText');
  if (updateDateEl) {
    const curKhmer = updateDateEl.innerText.trim();
    const isoVal = parseKhmerDateToIso(curKhmer);
    if (isoVal) {
      dateInput.value = isoVal;
    } else {
      dateInput.value = new Date().toISOString().split('T')[0];
    }
  }

  if (typeof dateInput.showPicker === 'function') {
    dateInput.showPicker();
  } else {
    dateInput.focus();
    dateInput.click();
  }
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលអ្នកប្រើជ្រើសរើសថ្ងៃពីប្រតិទិន
 * ធ្វើបច្ចុប្បន្នកម្មទាំងកាលបរិច្ឆេទធ្វើបច្ចុប្បន្នកម្ម និងកាលបរិច្ឆេទគ្រូបង្រៀន
 */
function handleDatePicked(dateStr) {
  if (!dateStr) return;
  const khmerDate = formatKhmerDateFromIso(dateStr);
  if (!khmerDate) return;

  const updateDateEl = document.getElementById('docUpdateDateText');
  if (updateDateEl) {
    updateDateEl.innerText = khmerDate;
  }

  syncTeacherSignatureDate(khmerDate);

  const dateInput = document.getElementById('docNativeDatePicker');
  if (dateInput) {
    dateInput.setAttribute('value', dateStr);
  }

  markDocumentModified();
  showToast(`បានជ្រើសរើសកាលបរិច្ឆេទ៖ ${khmerDate}`);
}

/**
 * ធ្វើសមកាលកម្មកាលបរិច្ឆេទទៅកាន់ប្លុកហត្ថលេខាគ្រូបង្រៀន / មន្ត្រីទទួលបន្ទុក
 */
function syncTeacherSignatureDate(khmerDate) {
  if (!khmerDate) return;
  const printableArea = document.getElementById('printableDocument');
  if (!printableArea) return;

  let teacherDateEl = document.getElementById('teacherSignatureDate');
  if (!teacherDateEl) {
    const allSigDates = printableArea.querySelectorAll('.doc-signatures .signature-block .signature-date');
    if (allSigDates.length >= 3) {
      teacherDateEl = allSigDates[2];
      teacherDateEl.id = 'teacherSignatureDate';
    }
  }

  if (teacherDateEl) {
    teacherDateEl.innerText = khmerDate.trim();
    teacherDateEl.style.transition = 'background-color 0.4s ease';
    teacherDateEl.style.backgroundColor = '#fef08a';
    setTimeout(() => {
      if (teacherDateEl) teacherDateEl.style.backgroundColor = 'transparent';
    }, 1200);
  }
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលអ្នកប្រើជ្រើសរើសឆ្នាំសិក្សាពី Dropdown
 */
function handleAcademicYearChange(val) {
  if (!val) return;
  const yearText = document.getElementById('docAcademicYearText');
  const yearSelect = document.getElementById('docAcademicYearSelect');

  if (val === 'custom') {
    const currentVal = yearText ? yearText.innerText.trim() : '';
    const customVal = prompt('សូមបញ្ចូលឆ្នាំសិក្សា (ឧទាហរណ៍៖ ២០៣០ - ២០៣១)៖', currentVal);
    if (customVal && customVal.trim()) {
      const trimmed = customVal.trim();
      if (yearText) yearText.innerText = trimmed;
      if (yearSelect) {
        let exists = Array.from(yearSelect.options).some(o => o.value === trimmed);
        if (!exists) {
          const opt = document.createElement('option');
          opt.value = trimmed;
          opt.innerText = trimmed;
          yearSelect.insertBefore(opt, yearSelect.lastElementChild);
        }
        yearSelect.value = trimmed;
        Array.from(yearSelect.options).forEach(o => {
          if (o.value === trimmed) o.setAttribute('selected', 'selected');
          else o.removeAttribute('selected');
        });
      }
      markDocumentModified();
      showToast(`បានជ្រើសរើសឆ្នាំសិក្សា៖ ${trimmed}`);
    } else {
      if (yearSelect && yearText) {
        yearSelect.value = yearText.innerText.trim();
      }
    }
    return;
  }

  if (yearText) {
    yearText.innerText = val;
  }
  if (yearSelect) {
    Array.from(yearSelect.options).forEach(o => {
      if (o.value === val) o.setAttribute('selected', 'selected');
      else o.removeAttribute('selected');
    });
  }
  markDocumentModified();
  showToast(`បានជ្រើសរើសឆ្នាំសិក្សា៖ ${val}`);
}

/**
 * ធ្វើសមកាលកម្ម Dropdown ឆ្នាំសិក្សា ពេលអ្នកប្រើវាយកែលើអត្ថបទផ្ទាល់
 */
function syncAcademicYearSelect(val) {
  const trimmed = val ? val.trim() : '';
  const yearSelect = document.getElementById('docAcademicYearSelect');
  if (yearSelect && trimmed) {
    let matched = Array.from(yearSelect.options).some(o => o.value === trimmed);
    if (matched) {
      yearSelect.value = trimmed;
      Array.from(yearSelect.options).forEach(o => {
        if (o.value === trimmed) o.setAttribute('selected', 'selected');
        else o.removeAttribute('selected');
      });
    }
  }
}

/**
 * បំពាក់ Dropdown List ឆ្នាំសិក្សា និងប្រតិទិន Calendar ទៅក្នុងសន្លឹកឯកសារ (ទាំងទម្រង់ដើម និងឯកសារចាស់ដែលធ្លាប់រក្សាទុក)
 */
function enhanceDocumentPickers() {
  const printableArea = document.getElementById('printableDocument');
  if (!printableArea) return;

  const metaItems = printableArea.querySelectorAll('.doc-meta-item');
  metaItems.forEach(item => {
    const strong = item.querySelector('strong');
    if (!strong) return;
    const labelText = strong.innerText.trim();

    // 1. Academic Year Dropdown
    if (labelText.includes('ឆ្នាំសិក្សា')) {
      let yearSpan = item.querySelector('.doc-meta-val') || item.querySelector('span.editable');
      if (yearSpan) {
        if (!yearSpan.id) yearSpan.id = 'docAcademicYearText';
        yearSpan.setAttribute('oninput', 'syncAcademicYearSelect(this.innerText)');

        let yearSelect = item.querySelector('.doc-year-select');
        const curYear = yearSpan.innerText.trim() || '២០២៣ - ២០២៤';
        const years = [
          '២០២២ - ២០២៣', '២០២៣ - ២០២៤', '២០២៤ - ២០២៥',
          '២០២៥ - ២០២៦', '២០២៦ - ២០២៧', '២០២៧ - ២០២៨',
          '២០២៨ - ២០២៩', '២០២៩ - ២០៣០'
        ];

        if (!yearSelect) {
          let optionsHtml = '<option value="" disabled>-- ជ្រើសរើសឆ្នាំ --</option>';
          years.forEach(y => {
            optionsHtml += `<option value="${y}"${y === curYear ? ' selected' : ''}>${y}</option>`;
          });
          if (!years.includes(curYear) && curYear) {
            optionsHtml += `<option value="${curYear}" selected>${curYear}</option>`;
          }
          optionsHtml += '<option value="custom">✍️ បញ្ចូលឆ្នាំផ្សេង...</option>';

          let wrap = item.querySelector('.meta-field-wrap');
          if (!wrap) {
            wrap = document.createElement('div');
            wrap.className = 'meta-field-wrap';
            yearSpan.parentNode.insertBefore(wrap, yearSpan);
            wrap.appendChild(yearSpan);
          }

          const sel = document.createElement('select');
          sel.className = 'doc-year-select no-print';
          sel.id = 'docAcademicYearSelect';
          sel.setAttribute('onchange', 'handleAcademicYearChange(this.value)');
          sel.title = 'ចុចជ្រើសរើសឆ្នាំសិក្សា';
          sel.innerHTML = optionsHtml;
          wrap.appendChild(sel);
        } else {
          yearSelect.value = curYear;
        }
      }
    }

    // 2. Date Picker Calendar
    if (labelText.includes('កាលបរិច្ឆេទ')) {
      let dateSpan = item.querySelector('.doc-meta-val') || item.querySelector('span.editable');
      if (dateSpan) {
        if (!dateSpan.id) dateSpan.id = 'docUpdateDateText';
        dateSpan.setAttribute('oninput', 'syncTeacherSignatureDate(this.innerText)');

        let calBtn = item.querySelector('.btn-calendar-trigger');
        let wrap = item.querySelector('.meta-field-wrap');
        if (!wrap) {
          wrap = document.createElement('div');
          wrap.className = 'meta-field-wrap';
          dateSpan.parentNode.insertBefore(wrap, dateSpan);
          wrap.appendChild(dateSpan);
        }

        if (!calBtn) {
          const btn = document.createElement('button');
          btn.type = 'button';
          btn.className = 'btn-calendar-trigger no-print';
          btn.onclick = openNativeDatePicker;
          btn.title = 'ចុចជ្រើសរើសកាលបរិច្ឆេទពីប្រតិទិន';
          btn.innerHTML = '<i class="fas fa-calendar-alt"></i> ប្រតិទិន';
          wrap.appendChild(btn);

          const hiddenInput = document.createElement('input');
          hiddenInput.type = 'date';
          hiddenInput.id = 'docNativeDatePicker';
          hiddenInput.className = 'native-date-input-hidden no-print';
          hiddenInput.onchange = function() { handleDatePicked(this.value); };
          const iso = parseKhmerDateToIso(dateSpan.innerText.trim());
          if (iso) hiddenInput.value = iso;
          wrap.appendChild(hiddenInput);
        } else {
          const hiddenInput = document.getElementById('docNativeDatePicker');
          if (hiddenInput) {
            const iso = parseKhmerDateToIso(dateSpan.innerText.trim());
            if (iso) hiddenInput.value = iso;
          }
        }
      }
    }
  });

  // 3. Teacher Signature Date Setup & Initial Sync
  let teacherDateEl = document.getElementById('teacherSignatureDate');
  if (!teacherDateEl) {
    const allSigDates = printableArea.querySelectorAll('.doc-signatures .signature-block .signature-date');
    if (allSigDates.length >= 3) {
      teacherDateEl = allSigDates[2];
      teacherDateEl.id = 'teacherSignatureDate';
    }
  }
  const updateDateEl = document.getElementById('docUpdateDateText');
  if (teacherDateEl && updateDateEl) {
    const tVal = teacherDateEl.innerText.trim();
    if (!tVal || tVal.includes('...') || tVal.includes('......')) {
      teacherDateEl.innerText = updateDateEl.innerText.trim();
    }
  }
}

/**
 * ពិនិត្យថាតើតម្លៃក្រឡាជាទិន្នន័យលេខឬអត់ ដើម្បីតម្រឹមកណ្តាល (Center Align)
 * គាំទ្រ៖ លេខខ្មែរ (០-៩), លេខឡាតាំង (0-9), រូបិយប័ណ្ណ (រៀល, ៛, $...), ភាគរយ (%),
 * កាលបរិច្ឆេទ (dd/mm/yyyy), កម្រិតថ្នាក់ (៣ ក, ៤ ខ...), ភេទ (ស្រី, ប្រុស)...
 */
function isNumericData(val) {
  if (val === null || val === undefined) return false;
  const str = String(val).trim();
  if (!str) return false;

  // ១. ភាគរយ ឬសមាមាត្រ %
  if (str.includes('%') || str.includes('％')) return true;

  // ២. កាលបរិច្ឆេទ ឬឆមាស៖ ឧ. 15/11/2023, ១៥/១១/២០២៣, 2023-11-15, ១៥-១១-២០២៣, ឆមាសទី១, ឆមាសទី២
  if (/^[\d០-៩]{1,4}[\/\-.][\d០-៩]{1,2}([\/\-.][\d០-៩]{1,4})?$/.test(str)) return true;
  if (/^ឆមាស(ទី)?\s*[\d០-៩]+/i.test(str)) return true;

  // ៣. លេខសុទ្ធ ឬលេខមានសញ្ញាក្បៀស ចុច បូក ដក (ខ្មែរ ឬឡាតាំង)៖ ឧ. ១, ២, 100, ៦០,០០០, 1,000, 1.5, ១.៥
  if (/^[+\-–—]?\s*[\d០-៩\s.,()]+$/.test(str) && /[\d០-៩]/.test(str)) return true;

  // ៤. ចំនួនទឹកប្រាក់ ឬចំនួនលេខភ្ជាប់ខ្នាត៖ ឧ. ៦០,០០០ រៀល, ៦០,០០០ រៀល/ខែ, 60,000៛, $50, ១.៥ គ.ក្រ, ១០ នាក់, ៥ ថ្នាក់, ១ គ្រឿង
  if (/^[$៛]?\s*[\d០-៩.,\s]+\s*([%％]|រៀល(\s*\/\s*ខែ)?|៛|\$|USD|ដុល្លារ|ខែ|\/ខែ|នាក់|ថ្នាក់|គ្រឿង|គ\.ក្រ|kg|ក្បាល)?$/i.test(str) && /[\d០-៩]/.test(str)) {
    return true;
  }

  // ៥. កម្រិតថ្នាក់ ឬកម្រិតក្រីក្រខ្លីៗ៖ ឧ. "៣ ក", "៤ ខ", "1A", "ថ្នាក់ទី១", "កម្រិត ១", "កម្រិត ២", "ក្រីក្រកម្រិត ១", "ក្រីក្រកម្រិត ២"
  if (/^(ថ្នាក់(ទី)?)?\s*[\d០-៩]+\s*([ក-អA-Z])?$/i.test(str) && /[\d០-៩]/.test(str)) return true;
  if (/^(ក្រីក្រ)?\s*កម្រិត\s*[\d០-៩]+/i.test(str)) return true;

  // ៦. ភេទ៖ ស្រី, ប្រុស, ស, ប, M, F
  if (/^(ប្រុស|ស្រី|ប|ស|M|F)$/i.test(str)) return true;

  return false;
}

/**
 * ញែកទិន្នន័យតារាងចេញពី HTML ដែលបាន Copy ពី Excel ឬ Google Sheet
 * - គាំទ្រការទាញយក headers និង rows
 * - រក្សាទុក colspan និង rowspan (បើមាន)
 */
function parseHtmlTableToMatrix(htmlText) {
  if (!htmlText || (!htmlText.includes('<table') && !htmlText.includes('<tr'))) return null;
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(htmlText, 'text/html');
    const table = doc.querySelector('table');
    if (!table) return null;

    let headers = [];
    let rows = [];

    const thead = table.querySelector('thead');
    if (thead) {
      const theadTr = thead.querySelector('tr');
      if (theadTr) {
        headers = Array.from(theadTr.querySelectorAll('th, td')).map(c => {
          return {
            text: (c.textContent || '').trim().replace(/\s+/g, ' '),
            colSpan: parseInt(c.getAttribute('colspan') || '1', 10),
            rowSpan: parseInt(c.getAttribute('rowspan') || '1', 10)
          };
        });
      }
    }

    const tbody = table.querySelector('tbody') || table;
    const allTrs = Array.from(table.querySelectorAll('tr')).filter(tr => !thead || !thead.contains(tr));

    if (allTrs.length > 0) {
      let startIndex = 0;
      if (headers.length === 0) {
        const firstTr = allTrs[0];
        const firstCells = Array.from(firstTr.querySelectorAll('th, td'));
        headers = firstCells.map(c => {
          return {
            text: (c.textContent || '').trim().replace(/\s+/g, ' '),
            colSpan: parseInt(c.getAttribute('colspan') || '1', 10),
            rowSpan: parseInt(c.getAttribute('rowspan') || '1', 10)
          };
        });
        startIndex = 1;
      }

      for (let i = startIndex; i < allTrs.length; i++) {
        const tr = allTrs[i];
        const cells = Array.from(tr.querySelectorAll('th, td')).map(c => {
          return {
            text: (c.textContent || '').trim().replace(/\s+/g, ' '),
            colSpan: parseInt(c.getAttribute('colspan') || '1', 10),
            rowSpan: parseInt(c.getAttribute('rowspan') || '1', 10)
          };
        });
        if (cells.some(c => c.text.length > 0)) {
          rows.push(cells);
        }
      }
    }

    if (headers.length === 0 && rows.length > 0) {
      headers = rows[0].map((c, idx) => ({ text: `កូឡោនទី ${khmerNumber(idx + 1)}`, colSpan: 1, rowSpan: 1 }));
    }

    return (headers.length > 0) ? { headers, rows } : null;
  } catch (err) {
    console.warn('parseHtmlTableToMatrix error:', err);
    return null;
  }
}

/**
 * ផ្លាស់ប្តូរតារាងទិន្នន័យលើសន្លឹករបាយការណ៍ទាំងស្រុង (ទាំងក្បាលជួរឈរ និងជួរដេកទាំងអស់)
 * តាមអ្វីដែលបានចម្លងមកពី Excel ឬ Google Sheet
 */
function replaceDocumentTableWithData(headers, rows) {
  if (!headers || headers.length === 0) return false;

  const printableArea = document.getElementById('printableDocument');
  if (!printableArea) return false;
  const table = printableArea.querySelector('.doc-table');
  if (!table) return false;

  let thead = table.querySelector('thead');
  if (!thead) {
    thead = document.createElement('thead');
    table.insertBefore(thead, table.firstChild);
  }
  let tbody = table.querySelector('tbody');
  if (!tbody) {
    tbody = document.createElement('tbody');
    table.appendChild(tbody);
  }

  const colCount = headers.length;
  const firstHeaderTitle = typeof headers[0] === 'object' ? headers[0].text : String(headers[0] || '');
  const isCol0Index = Boolean(firstHeaderTitle && /^(ល\.?រ|no|n°|លរ|#)/i.test(firstHeaderTitle.trim()));

  // ១. កសាងក្បាលតារាងថ្មី (New thead)
  let headHtml = '<tr>';
  headers.forEach((h, idx) => {
    const title = (typeof h === 'object' ? h.text : String(h)) || `កូឡោនទី ${khmerNumber(idx + 1)}`;
    const colSpan = (typeof h === 'object' && h.colSpan > 1) ? ` colspan="${h.colSpan}"` : '';
    const rowSpan = (typeof h === 'object' && h.rowSpan > 1) ? ` rowspan="${h.rowSpan}"` : '';
    headHtml += `<th class="editable-header" contenteditable="true"${colSpan}${rowSpan}>${escapeHtml(title)}</th>`;
  });
  headHtml += `<th class="row-action-col no-print" style="width: 58px; border: none; background: transparent;"></th>`;
  headHtml += '</tr>';
  thead.innerHTML = headHtml;

  // ២. កសាងជួរដេកទិន្នន័យថ្មី (New tbody)
  let bodyHtml = '';
  rows.forEach((row, rIdx) => {
    bodyHtml += '<tr>';
    for (let cIdx = 0; cIdx < colCount; cIdx++) {
      const cell = row[cIdx];
      let cellVal = typeof cell === 'object' ? (cell.text || '') : (cell !== undefined ? String(cell).trim() : '');
      
      // Auto convert index if column 0 is No/ល.រ
      if (cIdx === 0 && isCol0Index) {
        if (!cellVal || !isNaN(cellVal) || /^[០-៩\d]+$/.test(cellVal)) {
          cellVal = khmerNumber(cellVal || (rIdx + 1));
        }
      }
      const isCenter = (cIdx === 0 && isCol0Index) || isNumericData(cellVal);
      const colSpan = (typeof cell === 'object' && cell.colSpan > 1) ? ` colspan="${cell.colSpan}"` : '';
      const rowSpan = (typeof cell === 'object' && cell.rowSpan > 1) ? ` rowspan="${cell.rowSpan}"` : '';
      bodyHtml += `<td class="${isCenter ? 'text-center ' : ''}editable-cell" contenteditable="true"${colSpan}${rowSpan}>${escapeHtml(cellVal)}</td>`;
    }
    bodyHtml += `<td class="row-action-col no-print">${getRowActionButtonsHtml()}</td>`;
    bodyHtml += '</tr>';
  });

  // បើគ្មានជួរដេកទិន្នន័យ ផ្តល់ជួរដេកទទេគំរូមួយ
  if (rows.length === 0) {
    bodyHtml += '<tr>';
    for (let cIdx = 0; cIdx < colCount; cIdx++) {
      const isCenter = (cIdx === 0 && isCol0Index);
      const val = isCenter ? khmerNumber(1) : '';
      bodyHtml += `<td class="${isCenter ? 'text-center ' : ''}editable-cell" contenteditable="true">${val}</td>`;
    }
    bodyHtml += `<td class="row-action-col no-print">${getRowActionButtonsHtml()}</td>`;
    bodyHtml += '</tr>';
  }

  tbody.innerHTML = bodyHtml;

  // ៣. កំណត់ទំហំទទឹង និងទិសដៅក្រដាស (Orientation & Wide styling)
  if (colCount >= 10) {
    table.classList.add('table-ultra-wide');
    table.classList.remove('table-wide');
    setPrintOrientation('landscape');
  } else if (colCount >= 7) {
    table.classList.add('table-wide');
    table.classList.remove('table-ultra-wide');
    setPrintOrientation('landscape');
  } else {
    table.classList.remove('table-wide', 'table-ultra-wide');
  }

  // ៤. បំពាក់ Resizers និង Action buttons ឡើងវិញ
  ensureColumnActionButtonsOnHeaders(table);
  initTableColumnResizers();
  markDocumentModified();

  showToast(`⚡ បានផ្លាស់ប្តូរតារាងទិន្នន័យទាំងស្រុងតាម Excel/Sheet (${khmerNumber(colCount)} ជួរឈរ, ${khmerNumber(rows.length)} ជួរដេក) ជោគជ័យ!`);
  return true;
}

/**
 * បំពេញទិន្នន័យដែលបាន Paste ចូលក្នុងក្រឡានៃតារាងដែលមានស្រាប់ (ដោយរក្សាក្បាលតារាងដដែល)
 */
function fillPastedMatrixIntoExistingTable(matrix, targetCell) {
  if (!matrix || matrix.length === 0) return;

  const printableArea = document.getElementById('printableDocument');
  if (!printableArea) return;
  const table = printableArea.querySelector('.doc-table');
  if (!table) return;
  const tbody = table.querySelector('tbody');
  if (!tbody) return;

  if (!targetCell) {
    targetCell = window._lastActiveCell || tbody.querySelector('tr:first-child td:nth-child(2)');
  }

  if (targetCell && targetCell.closest('thead')) {
    const theadTr = targetCell.closest('tr');
    const thCells = Array.from(theadTr.querySelectorAll('th:not(.row-action-col)'));
    let colIdx = thCells.indexOf(targetCell);
    if (colIdx === -1) colIdx = 1;

    let firstTr = tbody.querySelector('tr');
    if (!firstTr) {
      addNewTableRow();
      firstTr = tbody.querySelector('tr');
    }
    const dataCells = Array.from(firstTr.querySelectorAll('td:not(.row-action-col)'));
    targetCell = dataCells[colIdx] || dataCells[1] || dataCells[0];
  }

  if (!targetCell || !targetCell.closest('tbody')) {
    targetCell = tbody.querySelector('tr:first-child td:nth-child(2)') || tbody.querySelector('td.editable-cell');
  }
  if (!targetCell) return;

  const targetTr = targetCell.closest('tr');
  const allRows = Array.from(tbody.querySelectorAll('tr'));
  const startRowIdx = allRows.indexOf(targetTr);
  if (startRowIdx === -1) return;

  const startRowCells = Array.from(targetTr.querySelectorAll('td:not(.row-action-col)'));
  const startColIdx = startRowCells.indexOf(targetCell);
  if (startColIdx === -1) return;

  const colCount = table.querySelectorAll('thead th:not(.row-action-col)').length;

  matrix.forEach((rowVals, lineOffset) => {
    const rIdx = startRowIdx + lineOffset;
    let tr = tbody.querySelectorAll('tr')[rIdx];

    // បើចំនួនជួរដេកដែល Paste លើសពីជួរដេកដែលមានស្រាប់ សូមបង្កើតជួរដេកថ្មីស្វ័យប្រវត្តិ
    if (!tr) {
      tr = document.createElement('tr');
      const curTotalRows = tbody.querySelectorAll('tr').length;
      for (let c = 0; c < colCount; c++) {
        const td = document.createElement('td');
        const isCenter = (c === 0);
        td.className = (isCenter ? 'text-center ' : '') + 'editable-cell';
        td.setAttribute('contenteditable', 'true');
        td.textContent = (c === 0) ? khmerNumber(curTotalRows + 1) : '';
        tr.appendChild(td);
      }
      const actionTd = document.createElement('td');
      actionTd.className = 'row-action-col no-print';
      actionTd.innerHTML = getRowActionButtonsHtml();
      tr.appendChild(actionTd);
      tbody.appendChild(tr);
    }

    const dataCells = Array.from(tr.querySelectorAll('td:not(.row-action-col)'));
    rowVals.forEach((val, cOffset) => {
      const cIdx = startColIdx + cOffset;
      if (cIdx < dataCells.length) {
        const rawText = typeof val === 'object' ? (val.text || '') : String(val || '');
        const cleanVal = rawText.trim().replace(/^"|"$/g, '').replace(/""/g, '"');
        dataCells[cIdx].textContent = cleanVal;
        if (cIdx === 0 || isNumericData(cleanVal)) {
          dataCells[cIdx].classList.add('text-center');
        } else {
          dataCells[cIdx].classList.remove('text-center');
        }
      }
    });
  });

  // ធានាលេខរៀង ល.រ ក្នុងជួរឈរទី១ រត់ត្រឹមត្រូវជាលេខខ្មែរ
  tbody.querySelectorAll('tr').forEach((row, idx) => {
    const firstTd = row.querySelector('td:first-child');
    if (firstTd && (!firstTd.textContent.trim() || !isNaN(firstTd.textContent.trim()) || /^[០-៩\d]+$/.test(firstTd.textContent.trim()))) {
      firstTd.textContent = khmerNumber(idx + 1);
      firstTd.classList.add('text-center');
    }
  });

  markDocumentModified();
  showToast(`⚡ បានបិទភ្ជាប់ទិន្នន័យ ${khmerNumber(matrix.length)} ជួរ ចូលក្នុងតារាងជោគជ័យ!`);
}

/**
 * វិភាគទិន្នន័យ Clipboard និងអនុវត្តការបិទភ្ជាប់ (ឆ្លាតវៃ៖ សួរអ្នកប្រើប្រាស់ដើម្បីជំនួសតារាងទាំងស្រុង ឬបំពេញតែក្រឡា)
 */
function applyPastedDataSmart(matrix, htmlParsed, targetCell) {
  let headers = [];
  let rows = [];

  if (htmlParsed && htmlParsed.headers && htmlParsed.headers.length > 0) {
    headers = htmlParsed.headers;
    rows = htmlParsed.rows || [];
  } else if (matrix && matrix.length > 0) {
    if (matrix.length > 1) {
      headers = matrix[0].map(c => typeof c === 'object' ? c : { text: String(c || '').trim(), colSpan: 1, rowSpan: 1 });
      rows = matrix.slice(1).map(r => r.map(c => typeof c === 'object' ? c : { text: String(c || '').trim(), colSpan: 1, rowSpan: 1 }));
    } else {
      headers = matrix[0].map(c => typeof c === 'object' ? c : { text: String(c || '').trim(), colSpan: 1, rowSpan: 1 });
      rows = [];
    }
  }

  // ពិនិត្យមើលថាតើទិន្នន័យដែលបាន Copy មកជាតារាងពេញលេញ (ច្រើនជួរឈរ ឬច្រើនជួរដេក) ឬទេ
  const isFullTableData = (headers.length > 1 && rows.length > 0) || (rows.length >= 2);

  if (isFullTableData) {
    const shouldReplaceCompletely = confirm(
      `📋 បានរកឃើញទិន្នន័យតារាងចម្លងពី Excel/Google Sheet (${khmerNumber(headers.length)} ជួរឈរ, ${khmerNumber(rows.length)} ជួរដេក)!\n\n` +
      `តើលោកគ្រូ-អ្នកគ្រូចង់ផ្លាស់ប្តូរតារាងទិន្នន័យ «ទាំងស្រុង» (ទាំងក្បាលជួរឈរ និងជួរដេកទាំងអស់) តាមទិន្នន័យដែលបាន Copy នេះដែរឬទេ?\n\n` +
      `• ចុច [OK] ៖ ផ្លាស់ប្តូរតារាងទាំងស្រុងតាម Excel/Sheet\n` +
      `• ចុច [Cancel] ៖ បិទភ្ជាប់បំពេញតែក្នុងក្រឡាទិន្នន័យធម្មតា (រក្សាក្បាលតារាងគំរូដើម)`
    );

    if (shouldReplaceCompletely) {
      replaceDocumentTableWithData(headers, rows);
      return;
    }
  }

  // Fallback: បំពេញទិន្នន័យចូលក្នុងក្រឡាតារាងបច្ចុប្បន្ន
  const fallbackMatrix = matrix || (htmlParsed ? [htmlParsed.headers.map(h => h.text), ...htmlParsed.rows.map(r => r.map(c => c.text))] : []);
  fillPastedMatrixIntoExistingTable(fallbackMatrix, targetCell);
}

/**
 * អនុវត្តការបិទភ្ជាប់ Matrix នៃទិន្នន័យពី Excel ឬ Google Sheet ចូលក្នុងតារាងឯកសារ
 * (រក្សាទុកសម្រាប់ភាពត្រូវគ្នាក្នុងប្រព័ន្ធ)
 */
function applyPastedMatrixToTable(matrix, targetCell) {
  applyPastedDataSmart(matrix, null, targetCell);
}

/**
 * ចាប់ព្រឹត្តិការណ៍បិទភ្ជាប់ (Paste) ទិន្នន័យពី Excel ឬ Google Sheet ចូលក្នុងតារាងឯកសារផ្ទាល់
 * នៅពេលលោកគ្រូ-អ្នកគ្រូចុចលើក្រឡាណាមួយក្នុងតារាង ហើយចុច Ctrl+V
 */
function handleTableClipboardPaste(e) {
  const activeEl = document.activeElement;
  const targetCell = (e.target && e.target.closest) ? e.target.closest('.editable-cell, .editable-header') : (activeEl && activeEl.closest ? activeEl.closest('.editable-cell, .editable-header') : window._lastActiveCell);
  if (!targetCell || !targetCell.closest('.doc-table')) return;

  const clipboardData = e.clipboardData || window.clipboardData;
  if (!clipboardData) return;

  // ១. សាកល្បងអានពី HTML Table (ប្រសិនបើ Copy ពី Google Sheet ឬ Excel លើ Browser)
  const htmlData = clipboardData.getData('text/html');
  let htmlParsed = null;
  if (htmlData && (htmlData.includes('<tr') || htmlData.includes('<table'))) {
    htmlParsed = parseHtmlTableToMatrix(htmlData);
  }

  // ២. អានពី Text Plain (TSV)
  const pastedText = clipboardData.getData('text/plain') || '';
  let matrix = [];
  if (pastedText) {
    const lines = pastedText.split(/\r?\n/).filter(line => line.trim().length > 0);
    if (lines.length > 0) {
      matrix = lines.map(line => line.split('\t'));
    }
  }

  // បើជាអត្ថបទខ្លីធម្មតាដែលគ្មាន \t ឬ \n អនុញ្ញាតឱ្យ Browser paste ធម្មតាចូលក្រឡានេះ
  if (!htmlParsed && (!pastedText.includes('\t') && !pastedText.includes('\n') && !pastedText.includes('\r'))) {
    setTimeout(() => {
      if (targetCell) {
        const val = targetCell.innerText.trim();
        if (isNumericData(val)) {
          targetCell.classList.add('text-center');
        }
      }
    }, 10);
    return;
  }

  if (!htmlParsed && matrix.length === 0) return;

  e.preventDefault();
  applyPastedDataSmart(matrix, htmlParsed, targetCell);
}

/**
 * បិទភ្ជាប់ទិន្នន័យពី Clipboard ចូលក្នុងតារាងតាមប៊ូតុងបញ្ជា
 */
async function pasteClipboardToTable() {
  const printableArea = document.getElementById('printableDocument');
  if (!printableArea) return;

  const targetCell = window._lastActiveCell || printableArea.querySelector('.doc-table tbody tr:first-child td:nth-child(2)');

  if (navigator.clipboard && navigator.clipboard.read) {
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes('text/html')) {
          const blob = await item.getType('text/html');
          const html = await blob.text();
          const parsed = parseHtmlTableToMatrix(html);
          if (parsed && parsed.headers && parsed.headers.length > 0) {
            applyPastedDataSmart(null, parsed, targetCell);
            return;
          }
        }
        if (item.types.includes('text/plain')) {
          const blob = await item.getType('text/plain');
          const text = await blob.text();
          if (text && text.trim()) {
            const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
            const matrix = lines.map(l => l.split('\t'));
            applyPastedDataSmart(matrix, null, targetCell);
            return;
          }
        }
      }
    } catch (err) {
      console.warn('navigator.clipboard.read failed or permission denied:', err);
    }
  }

  if (navigator.clipboard && navigator.clipboard.readText) {
    try {
      const text = await navigator.clipboard.readText();
      if (text && text.trim()) {
        const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
        const matrix = lines.map(l => l.split('\t'));
        applyPastedDataSmart(matrix, null, targetCell);
        return;
      }
    } catch (err) {
      console.warn('navigator.clipboard.readText failed:', err);
    }
  }

  // បើ Browser មិនអនុញ្ញាតឱ្យអាន Clipboard ផ្ទាល់ បើកផ្ទាំង Paste Modal
  openExcelImportModal(currentActiveDoc ? currentActiveDoc.code : null, 'paste');
}

/* ==========================================================================
   PRINT ORIENTATION & ANTI-CUTOFF MANAGEMENT
   ========================================================================== */

/**
 * អនុវត្ត Style ទិសដៅក្រដាសបោះពុម្ព (Portrait ឬ Landscape) ចូលក្នុង <head>
 * កំណត់ Margins ខាងឆ្វេង និងខាងស្តាំ 0.5cm ដូចគ្នា ដើម្បីកុំឱ្យដាច់ទិន្នន័យ
 */
function applyPrintOrientationStyles(orientation) {
  let styleEl = document.getElementById('dynamicPrintOrientationStyle');
  if (!styleEl) {
    styleEl = document.createElement('style');
    styleEl.id = 'dynamicPrintOrientationStyle';
    document.head.appendChild(styleEl);
  }

  const isLand = (orientation === 'landscape');
  styleEl.textContent = `
    @media print {
      @page {
        size: A4 ${isLand ? 'landscape' : 'portrait'} !important;
        margin: 0.5cm 0.5cm 0.5cm 0.5cm !important;
      }
    }
  `;
}

/**
 * កំណត់ទិសដៅក្រដាសបោះពុម្ព (ក្រដាសផ្តេក ឬ ក្រដាសបញ្ឈរ)
 */
function setPrintOrientation(orientation, showToastMsg = false) {
  currentPrintOrientation = (orientation === 'landscape') ? 'landscape' : 'portrait';
  applyPrintOrientationStyles(currentPrintOrientation);

  // ធ្វើបច្ចុប្បន្នកម្មប៊ូតុងជ្រើសរើសទម្រង់ក្រដាសទាំងពីរ
  const btnLand = document.getElementById('btnOrientLandscape');
  const btnPort = document.getElementById('btnOrientPortrait');

  if (currentPrintOrientation === 'landscape') {
    if (btnLand) {
      btnLand.classList.add('active-orientation');
      btnLand.title = 'បច្ចុប្បន្នកំពុងជ្រើស៖ ក្រដាសផ្តេក (A4 Landscape)';
    }
    if (btnPort) {
      btnPort.classList.remove('active-orientation');
      btnPort.title = 'ចុចដើម្បីប្តូរទៅជា៖ ក្រដាសបញ្ឈរ (A4 Portrait)';
    }
  } else {
    if (btnPort) {
      btnPort.classList.add('active-orientation');
      btnPort.title = 'បច្ចុប្បន្នកំពុងជ្រើស៖ ក្រដាសបញ្ឈរ (A4 Portrait)';
    }
    if (btnLand) {
      btnLand.classList.remove('active-orientation');
      btnLand.title = 'ចុចដើម្បីប្តូរទៅជា៖ ក្រដាសផ្តេក (A4 Landscape)';
    }
  }

  // គាំទ្រប៊ូតុង Toggle ចាស់ប្រសិនបើមាន
  const legacyBtn = document.getElementById('btnPrintOrientation');
  if (legacyBtn) {
    if (currentPrintOrientation === 'landscape') {
      legacyBtn.classList.add('is-landscape');
      legacyBtn.innerHTML = '<i class="fas fa-arrows-left-right"></i> <span id="lblPrintOrientation">ក្រដាសផ្តេក</span>';
    } else {
      legacyBtn.classList.remove('is-landscape');
      legacyBtn.innerHTML = '<i class="fas fa-arrows-up-down"></i> <span id="lblPrintOrientation">ក្រដាសបញ្ឈរ</span>';
    }
  }

  const docSheet = document.getElementById('printableDocument');
  if (docSheet) {
    if (currentPrintOrientation === 'landscape') {
      docSheet.classList.add('sheet-landscape');
    } else {
      docSheet.classList.remove('sheet-landscape');
    }
  }

  if (showToastMsg) {
    showToast(currentPrintOrientation === 'landscape' ? 'បានជ្រើសរើស៖ ក្រដាសផ្តេក (A4 Landscape)' : 'បានជ្រើសរើស៖ ក្រដាសបញ្ឈរ (A4 Portrait)');
  }
}

/**
 * ចុចប្តូរទិសដៅក្រដាសរវាង Portrait និង Landscape
 */
function togglePrintOrientation() {
  const nextOrientation = (currentPrintOrientation === 'landscape') ? 'portrait' : 'landscape';
  setPrintOrientation(nextOrientation, true);
}

/* ==========================================================================
   TABLE COLUMN RESIZING & WIDTH ADJUSTMENT ENGINE
   ========================================================================== */

/**
 * បំពាក់ឧបករណ៍អូសពង្រីក/បង្រួមជួរឈរ (Column Drag Resizer) លើគ្រប់ក្បាលជួរឈរ
 */
function initTableColumnResizers() {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const theadRow = table.querySelector('thead tr');
  if (!theadRow) return;

  const headers = Array.from(theadRow.querySelectorAll('th:not(.row-action-col)'));
  if (!headers.length) return;

  headers.forEach((th) => {
    // លុប resizer ចាស់ចេញប្រសិនបើមាន ដើម្បីកុំឱ្យស្ទួន
    const existing = th.querySelector('.col-resizer');
    if (existing) existing.remove();

    const resizer = document.createElement('div');
    resizer.className = 'col-resizer no-print';
    resizer.setAttribute('contenteditable', 'false');
    resizer.title = 'ចុចអូសទៅឆ្វេង ឬស្តាំ ដើម្បីបង្រួម ឬពង្រីកទទឹងជួរឈរនេះ';

    let startX = 0;
    let startWidth = 0;
    let isResizing = false;

    resizer.addEventListener('mousedown', (e) => {
      e.preventDefault();
      e.stopPropagation();

      startX = e.pageX;
      startWidth = th.offsetWidth;
      isResizing = true;
      resizer.classList.add('resizing');
      document.body.style.cursor = 'col-resize';
      document.body.style.userSelect = 'none';

      // កំណត់ទំហំទទឹងបច្ចុប្បន្នជា inline style លើគ្រប់ th ដើម្បីការពារការរង្គោះរង្គើ
      headers.forEach(h => {
        if (!h.style.width) {
          h.style.width = h.offsetWidth + 'px';
        }
      });

      const onMouseMove = (moveEvent) => {
        if (!isResizing) return;
        const diffX = moveEvent.pageX - startX;
        const newWidth = Math.max(40, startWidth + diffX);
        th.style.width = newWidth + 'px';
        th.style.minWidth = newWidth + 'px';
      };

      const onMouseUp = () => {
        if (!isResizing) return;
        isResizing = false;
        resizer.classList.remove('resizing');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);

        markDocumentModified();
      };

      document.addEventListener('mousemove', onMouseMove);
      document.addEventListener('mouseup', onMouseUp);
    });

    th.appendChild(resizer);
  });
}

/**
 * ស្វែងរក Index នៃជួរឈរដែលកំពុងជ្រើសរើស (Active Column)
 */
function getActiveTableColumnIndex() {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return 0;

  const cell = window._lastActiveCell;
  if (cell && table.contains(cell)) {
    const tr = cell.closest('tr');
    if (tr) {
      const dataCells = Array.from(tr.children).filter(c => !c.classList.contains('row-action-col'));
      const idx = dataCells.indexOf(cell);
      if (idx !== -1) return idx;
    }
  }
  return 1; // Default to Column 1 (usually Name / Main data) if not set
}

/**
 * ពង្រីក ឬបង្រួមទទឹងជួរឈរដែលកំពុងជ្រើសរើស
 * @param {number} delta (+25 ឬ -25 pixels)
 */
function adjustActiveColumnWidth(delta) {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const theadRow = table.querySelector('thead tr');
  if (!theadRow) return;

  const headers = Array.from(theadRow.querySelectorAll('th:not(.row-action-col)'));
  if (!headers.length) return;

  const colIdx = getActiveTableColumnIndex();
  const targetTh = headers[colIdx] || headers[0];
  if (!targetTh) return;

  const currentWidth = targetTh.offsetWidth || 100;
  const newWidth = Math.max(40, currentWidth + delta);

  targetTh.style.width = newWidth + 'px';
  targetTh.style.minWidth = newWidth + 'px';

  // ធានាថាក្បាលជួរឈរផ្សេងទៀតរក្សាទទឹងជាក់ស្តែង
  headers.forEach((h, i) => {
    if (i !== colIdx && !h.style.width) {
      h.style.width = h.offsetWidth + 'px';
    }
  });

  markDocumentModified();
  const colName = targetTh.innerText.replace(/[\r\n\t]/g, ' ').trim() || `កូឡោនទី ${khmerNumber(colIdx + 1)}`;
  showToast(`${delta > 0 ? 'ពង្រីក' : 'បង្រួម'} «${colName}» មកត្រឹម ${newWidth}px`);
}

/**
 * លៃតម្រូវទទឹងគ្រប់ជួរឈរដោយស្វ័យប្រវត្តិតាមទិន្នន័យ (Auto-fit)
 */
function autoFitTableColumns() {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const headers = Array.from(table.querySelectorAll('thead th:not(.row-action-col)'));
  headers.forEach(th => {
    th.style.width = '';
    th.style.minWidth = '';
    th.style.maxWidth = '';
  });

  markDocumentModified();
  showToast('បានលៃតម្រូវទទឹងជួរឈរទាំងអស់តាមទិន្នន័យជាក់ស្តែង (Auto-fit)');
}

/**
 * ចែករំលែកទទឹងជួរឈរទាំងអស់ឱ្យស្មើគ្នា (Equalize Column Widths)
 */
function equalizeTableColumns() {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const headers = Array.from(table.querySelectorAll('thead th:not(.row-action-col)'));
  if (!headers.length) return;

  const equalPct = (100 / headers.length).toFixed(2);
  headers.forEach(th => {
    th.style.width = equalPct + '%';
    th.style.minWidth = '';
    th.style.maxWidth = '';
  });

  markDocumentModified();
  showToast(`បានកំណត់ទទឹងជួរឈរទាំងអស់ស្មើគ្នា (${equalPct}%)`);
}

/* ==========================================================================
   TABLE CELL MERGING & MULTI-CELL SELECTION ENGINE
   ========================================================================== */

let isSelectingTableCells = false;
let selectionAnchorCell = null;
let currentSelectedCells = new Set();
let isMouseDownOnTable = false;
let contextTargetCell = null;

/**
 * គណនា 2D Coordinate Grid Matrix នៃតារាង (គាំទ្រ rowspan & colspan គ្រប់ទម្រង់)
 */
function getTableGridMatrix(table) {
  if (!table) return { grid: [], cellMeta: new Map(), rows: [], maxCols: 0 };
  const rows = Array.from(table.querySelectorAll('tr'));
  const grid = [];
  const cellMeta = new Map();
  let maxCols = 0;

  rows.forEach((tr, r) => {
    if (!grid[r]) grid[r] = [];
    let c = 0;
    const cells = Array.from(tr.children).filter(cell => !cell.classList.contains('row-action-col'));

    cells.forEach(cell => {
      while (grid[r][c]) {
        c++;
      }
      const rowSpan = parseInt(cell.getAttribute('rowspan') || '1', 10);
      const colSpan = parseInt(cell.getAttribute('colspan') || '1', 10);

      const meta = { cell, r, c, rowSpan, colSpan, tr };
      cellMeta.set(cell, meta);

      for (let i = 0; i < rowSpan; i++) {
        const targetR = r + i;
        if (!grid[targetR]) grid[targetR] = [];
        for (let j = 0; j < colSpan; j++) {
          const targetC = c + j;
          grid[targetR][targetC] = {
            cell,
            isOrigin: (i === 0 && j === 0),
            originR: r,
            originC: c,
            rowSpan,
            colSpan,
            tr
          };
          if (targetC + 1 > maxCols) maxCols = targetC + 1;
        }
      }
      c += colSpan;
    });
  });

  return { grid, cellMeta, rows, maxCols };
}

/**
 * សម្អាតរាល់ការជ្រើសរើសក្រឡា (Clear Cell Selection)
 */
function clearCellSelection() {
  currentSelectedCells.forEach(cell => {
    if (cell) cell.classList.remove('cell-selected');
  });
  currentSelectedCells.clear();
  selectionAnchorCell = null;
  removeFloatingMergeToolbar();
  closeMergeDropdown();
  closeTableContextMenu();
}

/**
 * ជ្រើសរើសចតុកោណកែងនៃក្រឡាចន្លោះ cellA និង cellB
 */
function selectCellsBetween(cellA, cellB, table) {
  if (!table) table = cellA.closest('table');
  if (!table) return;

  const { grid, cellMeta } = getTableGridMatrix(table);
  const metaA = cellMeta.get(cellA);
  const metaB = cellMeta.get(cellB);
  if (!metaA || !metaB) return;

  let minR = Math.min(metaA.r, metaB.r);
  let maxR = Math.max(metaA.r + metaA.rowSpan - 1, metaB.r + metaB.rowSpan - 1);
  let minC = Math.min(metaA.c, metaB.c);
  let maxC = Math.max(metaA.c + metaA.colSpan - 1, metaB.c + metaB.colSpan - 1);

  // Expand boundaries if any spanned cell extends outside
  let changed = true;
  while (changed) {
    changed = false;
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const item = grid[r] && grid[r][c];
        if (item) {
          const oR = item.originR;
          const eR = item.originR + item.rowSpan - 1;
          const oC = item.originC;
          const eC = item.originC + item.colSpan - 1;
          if (oR < minR) { minR = oR; changed = true; }
          if (eR > maxR) { maxR = eR; changed = true; }
          if (oC < minC) { minC = oC; changed = true; }
          if (eC > maxC) { maxC = eC; changed = true; }
        }
      }
    }
  }

  // Clear existing highlight
  currentSelectedCells.forEach(c => c.classList.remove('cell-selected'));
  currentSelectedCells.clear();

  for (let r = minR; r <= maxR; r++) {
    for (let c = minC; c <= maxC; c++) {
      const item = grid[r] && grid[r][c];
      if (item && item.cell) {
        currentSelectedCells.add(item.cell);
        item.cell.classList.add('cell-selected');
      }
    }
  }

  if (currentSelectedCells.size > 1) {
    updateFloatingMergeToolbar();
  } else {
    removeFloatingMergeToolbar();
  }
}

/**
 * បង្ហាញ ឬធ្វើបច្ចុប្បន្នកម្ម Floating Toolbar លើតំបន់ក្រឡាដែលបានជ្រើស
 */
function updateFloatingMergeToolbar() {
  removeFloatingMergeToolbar();
  if (currentSelectedCells.size <= 1) return;

  const cells = Array.from(currentSelectedCells);
  const firstCell = cells[0];
  if (!firstCell) return;

  const rect = firstCell.getBoundingClientRect();
  const toolbar = document.createElement('div');
  toolbar.className = 'floating-merge-toolbar no-print';
  toolbar.id = 'floatingMergeToolbar';
  toolbar.style.top = Math.max(10, window.scrollY + rect.top - 38) + 'px';
  toolbar.style.left = (window.scrollX + rect.left) + 'px';

  toolbar.innerHTML = `
    <span><i class="fas fa-check-square"></i> បានជ្រើស ${khmerNumber(currentSelectedCells.size)} ក្រឡា</span>
    <button type="button" class="btn-floating-merge" onclick="mergeSelectedTableCells()">
      <i class="fas fa-object-group"></i> បញ្ចូលក្រឡា (Merge)
    </button>
    <button type="button" class="btn-floating-close" onclick="clearCellSelection()" title="បោះបង់">
      <i class="fas fa-times"></i>
    </button>
  `;

  document.body.appendChild(toolbar);
}

function removeFloatingMergeToolbar() {
  const el = document.getElementById('floatingMergeToolbar');
  if (el) el.remove();
}

/**
 * បញ្ចូលក្រឡាដែលបានជ្រើសរើស (Merge Selected Cells)
 */
function mergeSelectedTableCells() {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  if (currentSelectedCells.size <= 1) {
    const active = window._lastActiveCell;
    if (active && table.contains(active)) {
      toggleMergeDropdown();
      return;
    }
    showToast('សូមជ្រើសរើសក្រឡាយ៉ាងហោចណាស់ ២ (ដោយអូស Mouse ឬសង្កត់ Shift) ដើម្បីបញ្ចូលគ្នា!');
    return;
  }

  const { grid, cellMeta } = getTableGridMatrix(table);
  const selectedArr = Array.from(currentSelectedCells);

  let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
  selectedArr.forEach(cell => {
    const meta = cellMeta.get(cell);
    if (meta) {
      minR = Math.min(minR, meta.r);
      maxR = Math.max(maxR, meta.r + meta.rowSpan - 1);
      minC = Math.min(minC, meta.c);
      maxC = Math.max(maxC, meta.c + meta.colSpan - 1);
    }
  });

  const targetRowSpan = maxR - minR + 1;
  const targetColSpan = maxC - minC + 1;

  if (targetRowSpan === 1 && targetColSpan === 1) {
    showToast('សូមជ្រើសរើសក្រឡាលើសពី ១ ដើម្បីបញ្ចូលគ្នា!');
    return;
  }

  const originItem = grid[minR] && grid[minR][minC];
  if (!originItem || !originItem.cell) {
    showToast('មិនអាចបញ្ចូលក្រឡានេះបានឡើយ!');
    return;
  }

  const primaryCell = originItem.cell;

  // Collect text contents from all cells in the merge region
  let texts = [];
  const primaryTxt = primaryCell.innerText.trim();
  if (primaryTxt) texts.push(primaryTxt);

  selectedArr.forEach(cell => {
    if (cell !== primaryCell) {
      const t = cell.innerText.trim();
      if (t && !texts.includes(t)) {
        texts.push(t);
      }
      cell.remove();
    }
  });

  if (targetColSpan > 1) {
    primaryCell.setAttribute('colspan', targetColSpan);
  } else {
    primaryCell.removeAttribute('colspan');
  }

  if (targetRowSpan > 1) {
    primaryCell.setAttribute('rowspan', targetRowSpan);
  } else {
    primaryCell.removeAttribute('rowspan');
  }

  if (texts.length > 0) {
    primaryCell.innerText = texts.join(' ');
  }

  clearCellSelection();
  window._lastActiveCell = primaryCell;
  primaryCell.focus();

  markDocumentModified();
  initTableColumnResizers();
  ensureColumnActionButtonsOnHeaders(table);

  showToast(`បានបញ្ចូលក្រឡា ${khmerNumber(targetColSpan)}x${khmerNumber(targetRowSpan)} ជោគជ័យ!`);
}

/**
 * បញ្ចូលក្រឡាបច្ចុប្បន្នជាមួយក្រឡាខាងស្តាំ (Merge with right neighbor)
 */
function mergeActiveCellRight(targetCell) {
  const cell = targetCell || window._lastActiveCell;
  if (!cell) {
    showToast('សូមចុចជ្រើសរើសក្រឡាមួយជាមុនសិន!');
    return;
  }
  const table = cell.closest('table');
  if (!table) return;

  const { grid, cellMeta } = getTableGridMatrix(table);
  const meta = cellMeta.get(cell);
  if (!meta) return;

  const rightC = meta.c + meta.colSpan;
  const rightItem = grid[meta.r] && grid[meta.r][rightC];
  if (!rightItem || !rightItem.cell || rightItem.cell === cell) {
    showToast('គ្មានក្រឡានៅខាងស្តាំដើម្បីបញ្ចូលគ្នាទៀតទេ!');
    return;
  }

  clearCellSelection();
  currentSelectedCells.add(cell);
  currentSelectedCells.add(rightItem.cell);
  mergeSelectedTableCells();
}

/**
 * បញ្ចូលក្រឡាបច្ចុប្បន្នជាមួយក្រឡាខាងក្រោម (Merge with bottom neighbor)
 */
function mergeActiveCellDown(targetCell) {
  const cell = targetCell || window._lastActiveCell;
  if (!cell) {
    showToast('សូមចុចជ្រើសរើសក្រឡាមួយជាមុនសិន!');
    return;
  }
  const table = cell.closest('table');
  if (!table) return;

  const { grid, cellMeta } = getTableGridMatrix(table);
  const meta = cellMeta.get(cell);
  if (!meta) return;

  const downR = meta.r + meta.rowSpan;
  const downItem = grid[downR] && grid[downR][meta.c];
  if (!downItem || !downItem.cell || downItem.cell === cell) {
    showToast('គ្មានក្រឡានៅខាងក្រោមដើម្បីបញ្ចូលគ្នាទៀតទេ!');
    return;
  }

  clearCellSelection();
  currentSelectedCells.add(cell);
  currentSelectedCells.add(downItem.cell);
  mergeSelectedTableCells();
}

/**
 * បំបែកក្រឡា (Unmerge Selected Cells)
 */
function unmergeSelectedTableCells(targetCell) {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  let cellsToUnmerge = [];
  if (targetCell) {
    cellsToUnmerge.push(targetCell);
  } else if (currentSelectedCells.size > 0) {
    cellsToUnmerge = Array.from(currentSelectedCells);
  } else if (window._lastActiveCell && table.contains(window._lastActiveCell)) {
    cellsToUnmerge.push(window._lastActiveCell);
  }

  if (cellsToUnmerge.length === 0) {
    showToast('សូមជ្រើសរើសក្រឡាដែលបានបញ្ចូល (Merged Cell) ដើម្បីបំបែក!');
    return;
  }

  let unmergedCount = 0;

  cellsToUnmerge.forEach(cell => {
    const colSpan = parseInt(cell.getAttribute('colspan') || '1', 10);
    const rowSpan = parseInt(cell.getAttribute('rowspan') || '1', 10);
    if (colSpan === 1 && rowSpan === 1) return;

    const { grid, cellMeta, rows, maxCols } = getTableGridMatrix(table);
    const meta = cellMeta.get(cell);
    if (!meta) return;

    const isTh = cell.tagName === 'TH';
    const cellTag = isTh ? 'th' : 'td';
    const cellClass = isTh ? 'editable-header' : 'editable-cell';

    // 1. In current row meta.r: insert colSpan - 1 cells after cell
    let prevCell = cell;
    for (let j = 1; j < colSpan; j++) {
      const newCell = document.createElement(cellTag);
      newCell.className = cellClass;
      newCell.setAttribute('contenteditable', 'true');
      newCell.textContent = '';
      if (meta.c + j === 0) newCell.classList.add('text-center');
      if (prevCell.nextSibling) {
        prevCell.parentNode.insertBefore(newCell, prevCell.nextSibling);
      } else {
        prevCell.parentNode.appendChild(newCell);
      }
      prevCell = newCell;
    }

    // 2. In subsequent rows meta.r + i: insert colSpan cells
    for (let i = 1; i < rowSpan; i++) {
      const targetR = meta.r + i;
      const targetTr = rows[targetR];
      if (!targetTr) continue;

      let anchorCell = null;
      for (let c = meta.c + colSpan; c < maxCols; c++) {
        const item = grid[targetR] && grid[targetR][c];
        if (item && item.isOrigin && item.cell && item.cell.parentNode === targetTr) {
          anchorCell = item.cell;
          break;
        }
      }
      if (!anchorCell) {
        anchorCell = targetTr.querySelector('.row-action-col');
      }

      for (let j = 0; j < colSpan; j++) {
        const newCell = document.createElement(cellTag);
        newCell.className = cellClass;
        newCell.setAttribute('contenteditable', 'true');
        newCell.textContent = '';
        if (meta.c + j === 0) newCell.classList.add('text-center');

        if (anchorCell) {
          targetTr.insertBefore(newCell, anchorCell);
        } else {
          targetTr.appendChild(newCell);
        }
      }
    }

    cell.removeAttribute('colspan');
    cell.removeAttribute('rowspan');
    unmergedCount++;
  });

  if (unmergedCount > 0) {
    clearCellSelection();
    markDocumentModified();
    initTableColumnResizers();
    ensureColumnActionButtonsOnHeaders(table);
    showToast(`បានបំបែកក្រឡា ${khmerNumber(unmergedCount)} ទៅជាទម្រង់ធម្មតាវិញជោគជ័យ!`);
  } else {
    showToast('ក្រឡានេះមិនមែនជាក្រឡាដែលបានបញ្ចូលគ្នា (Merged Cell) ទេ!');
  }
}

/**
 * បើក ឬបិទម៉ឺនុយទម្លាក់ចុះនៃប៊ូតុង Merge (Toggle Merge Dropdown)
 */
function toggleMergeDropdown(e) {
  if (e) {
    e.preventDefault();
    e.stopPropagation();
  }
  // If multiple cells are already selected, clicking Merge directly merges them!
  if (currentSelectedCells.size > 1) {
    mergeSelectedTableCells();
    closeMergeDropdown();
    return;
  }

  const dropdown = document.getElementById('mergeDropdownMenu');
  if (!dropdown) return;
  const isHidden = (dropdown.style.display === 'none' || !dropdown.style.display);
  dropdown.style.display = isHidden ? 'flex' : 'none';
}

function closeMergeDropdown() {
  const dropdown = document.getElementById('mergeDropdownMenu');
  if (dropdown) dropdown.style.display = 'none';
}

/**
 * Context Menu (Right Click) Handling
 */
function openTableContextMenu(e, cell) {
  e.preventDefault();
  e.stopPropagation();

  contextTargetCell = cell;
  window._lastActiveCell = cell;

  if (!currentSelectedCells.has(cell)) {
    clearCellSelection();
    currentSelectedCells.add(cell);
    cell.classList.add('cell-selected');
  }

  const menu = document.getElementById('tableContextMenu');
  if (!menu) return;

  menu.style.display = 'block';
  const menuWidth = 250;
  const menuHeight = 320;
  let posX = e.clientX;
  let posY = e.clientY;

  if (posX + menuWidth > window.innerWidth) {
    posX = window.innerWidth - menuWidth - 10;
  }
  if (posY + menuHeight > window.innerHeight) {
    posY = window.innerHeight - menuHeight - 10;
  }

  menu.style.left = posX + 'px';
  menu.style.top = posY + 'px';
}

function closeTableContextMenu() {
  const menu = document.getElementById('tableContextMenu');
  if (menu) menu.style.display = 'none';
}

function handleContextMenuAction(action) {
  closeTableContextMenu();
  const target = contextTargetCell || window._lastActiveCell;

  switch (action) {
    case 'merge':
      mergeSelectedTableCells();
      break;
    case 'merge-right':
      mergeActiveCellRight(target);
      break;
    case 'merge-down':
      mergeActiveCellDown(target);
      break;
    case 'unmerge':
      unmergeSelectedTableCells(target);
      break;
    case 'insert-row':
      if (target) {
        const tr = target.closest('tr');
        if (tr) {
          const insertBtn = tr.querySelector('.row-insert-btn') || { closest: () => tr };
          insertTableRowAfter(insertBtn);
        }
      } else {
        addNewTableRow();
      }
      break;
    case 'delete-row':
      if (target) {
        const tr = target.closest('tr');
        if (tr) {
          deleteRowAt(tr, false);
        }
      } else {
        deleteTableRowSmart();
      }
      break;
    case 'insert-col':
      if (target) {
        const table = target.closest('table');
        const ths = table ? Array.from(table.querySelectorAll('thead th:not(.row-action-col)')) : [];
        const { cellMeta } = getTableGridMatrix(table);
        const meta = cellMeta.get(target);
        const colIdx = meta ? meta.c : 0;
        const targetTh = ths[colIdx] || ths[ths.length - 1];
        if (targetTh) insertTableColumnAfter(targetTh);
        else addNewTableColumn();
      } else {
        addNewTableColumn();
      }
      break;
    case 'delete-col':
      if (target) {
        const table = target.closest('table') || document.querySelector('#printableDocument .doc-table');
        if (table) {
          const { cellMeta } = getTableGridMatrix(table);
          const meta = cellMeta.get(target);
          const colIdx = meta ? meta.c : getActiveTableColumnIndex();
          deleteColumnByIndex(table, colIdx, true);
        }
      } else {
        deleteLastTableColumn();
      }
      break;
    case 'paste':
      pasteClipboardToTable();
      break;
  }
}

/**
 * កំណត់ Events សម្រាប់ការកែប្រែ និងផ្លាស់ទីក្នុងតារាង
 */
function initTableEditingFeatures() {
  const printableArea = document.getElementById('printableDocument');
  if (!printableArea) return;

  // ដំណើរការបំពាក់ Dropdown List ឆ្នាំសិក្សា និងប្រតិទិន Calendar
  enhanceDocumentPickers();

  // ធានាវត្តមានប៊ូតុងបន្ថែម/លុបជួរឈរ និងប៊ូតុងបិទភ្ជាប់
  ensureColumnActionButtons();

  // ធានាវត្តមានប៊ូតុងសកម្មភាព (+) និង (×) លើគ្រប់ជួរដេកនៃតារាង
  ensureRowActionButtons();

  // បំពាក់ឧបករណ៍អូសពង្រីក/បង្រួមជួរឈរ (Column Drag Resizers)
  initTableColumnResizers();

  // បំពាក់ប៊ូតុងសកម្មភាព (+) និង (×) លើគ្រប់ក្បាលជួរឈរនៃតារាង
  ensureColumnActionButtonsOnHeaders();

  // តាមដានក្រឡាដែលបានចុចចុងក្រោយ (Track last active cell)
  printableArea.addEventListener('focusin', (e) => {
    const td = e.target.closest('.editable-cell, .editable-header');
    if (td) window._lastActiveCell = td;
  });
  printableArea.addEventListener('click', (e) => {
    const td = e.target.closest('.editable-cell, .editable-header');
    if (td) window._lastActiveCell = td;
  });

  // ចាប់ព្រឹត្តិការណ៍វាយអក្សរកែប្រែទិន្នន័យ (Input change detection)
  printableArea.addEventListener('input', (e) => {
    markDocumentModified();
    if (e.target && e.target.id === 'docUpdateDateText') {
      syncTeacherSignatureDate(e.target.innerText);
    }
  });

  // ពេលចាកចេញពីក្រឡា (blur/focusout) ពិនិត្យតម្រឹមកណ្តាលលេខស្វ័យប្រវត្តិ
  printableArea.addEventListener('focusout', (e) => {
    const td = e.target.closest('.editable-cell');
    if (td) {
      const tr = td.closest('tr');
      if (tr) {
        const dataCells = Array.from(tr.querySelectorAll('td:not(.row-action-col)'));
        const colIdx = dataCells.indexOf(td);
        if (colIdx !== 0) {
          const val = td.innerText.trim();
          if (isNumericData(val)) {
            td.classList.add('text-center');
          } else {
            td.classList.remove('text-center');
          }
        }
      }
    }
  });

  // ធានាថារាល់ទិន្នន័យជាលេខដែលមានស្រាប់ ត្រូវបានតម្រឹមកណ្តាល (Center align)
  printableArea.querySelectorAll('.doc-table tbody tr').forEach(tr => {
    const dataCells = Array.from(tr.querySelectorAll('td:not(.row-action-col)'));
    dataCells.forEach((td, cIdx) => {
      const val = td.innerText.trim();
      if (cIdx === 0 || isNumericData(val)) {
        td.classList.add('text-center');
      }
    });
  });

  // ចាប់ព្រឹត្តិការណ៍ចុចលើប៊ូតុងបន្ថែម (+) ឬលុប (×) ជួរឈរ និងជួរដេក (Event delegation)
  printableArea.addEventListener('click', (e) => {
    const colInsertBtn = e.target.closest('.col-insert-btn');
    if (colInsertBtn) {
      e.preventDefault();
      e.stopPropagation();
      const targetTh = colInsertBtn.closest('th');
      if (targetTh) insertTableColumnAfter(targetTh);
      return;
    }

    const colDeleteBtn = e.target.closest('.col-delete-btn');
    if (colDeleteBtn) {
      e.preventDefault();
      e.stopPropagation();
      const targetTh = colDeleteBtn.closest('th');
      if (targetTh) deleteTableColumnAt(targetTh);
      return;
    }

    const insertBtn = e.target.closest('.row-insert-btn');
    if (insertBtn) {
      insertTableRowAfter(insertBtn);
      return;
    }

    const delBtn = e.target.closest('.row-delete-btn');
    if (delBtn) {
      deleteCurrentTableRow(delBtn);
      return;
    }
  });

  // បំពាក់ការជ្រើសរើសក្រឡាច្រើន (Cell Selection) និង Right-Click Context Menu លើតារាង
  const docTable = printableArea.querySelector('.doc-table');
  if (docTable) {
    // Mousedown on cells
    docTable.addEventListener('mousedown', (e) => {
      const cell = e.target.closest('.editable-cell, .editable-header');
      if (!cell || cell.classList.contains('row-action-col') || e.target.closest('.col-resizer, .col-action-btns, .row-action-btns')) return;

      if (e.button === 0) { // Left Click
        if (e.shiftKey && selectionAnchorCell) {
          e.preventDefault();
          selectCellsBetween(selectionAnchorCell, cell, docTable);
          return;
        }
        if (e.ctrlKey || e.metaKey) {
          e.preventDefault();
          if (currentSelectedCells.has(cell)) {
            currentSelectedCells.delete(cell);
            cell.classList.remove('cell-selected');
          } else {
            currentSelectedCells.add(cell);
            cell.classList.add('cell-selected');
          }
          if (currentSelectedCells.size > 1) updateFloatingMergeToolbar();
          else removeFloatingMergeToolbar();
          return;
        }

        isMouseDownOnTable = true;
        selectionAnchorCell = cell;
        window._lastActiveCell = cell;
      }
    });

    // Mouseover while dragging to select
    docTable.addEventListener('mouseover', (e) => {
      if (!isMouseDownOnTable || !selectionAnchorCell) return;
      const cell = e.target.closest('.editable-cell, .editable-header');
      if (!cell || cell.classList.contains('row-action-col')) return;

      if (cell !== selectionAnchorCell) {
        isSelectingTableCells = true;
        docTable.classList.add('table-selecting-active');
        selectCellsBetween(selectionAnchorCell, cell, docTable);
      }
    });

    // Context Menu (Right Click)
    docTable.addEventListener('contextmenu', (e) => {
      const cell = e.target.closest('.editable-cell, .editable-header');
      if (!cell || cell.classList.contains('row-action-col')) return;
      openTableContextMenu(e, cell);
    });
  }

  // Mouseup on document
  const onDocMouseUp = () => {
    if (isMouseDownOnTable) {
      isMouseDownOnTable = false;
      if (docTable) docTable.classList.remove('table-selecting-active');
    }
  };
  document.removeEventListener('mouseup', window._tableMouseUpHandler);
  window._tableMouseUpHandler = onDocMouseUp;
  document.addEventListener('mouseup', onDocMouseUp);

  // Click on document to close menus and deselect if clicked outside
  const onDocClick = (e) => {
    if (!e.target.closest('#tableContextMenu')) {
      closeTableContextMenu();
    }
    if (!e.target.closest('.merge-dropdown-wrap')) {
      closeMergeDropdown();
    }
    if (!e.target.closest('.doc-table') && !e.target.closest('.table-actions-bar') && !e.target.closest('#floatingMergeToolbar') && !e.target.closest('#tableContextMenu')) {
      clearCellSelection();
    }
  };
  document.removeEventListener('click', window._tableClickHandler);
  window._tableClickHandler = onDocClick;
  document.addEventListener('click', onDocClick);

  // Keyboard Shortcuts (Alt+M to merge, Alt+Shift+M to unmerge, Escape to clear)
  const onDocKeyDown = (e) => {
    if (e.key === 'Escape') {
      clearCellSelection();
      closeTableContextMenu();
      closeMergeDropdown();
      return;
    }
    if (e.altKey && !e.shiftKey && (e.key === 'm' || e.key === 'M' || e.code === 'KeyM')) {
      e.preventDefault();
      mergeSelectedTableCells();
      return;
    }
    if (e.altKey && e.shiftKey && (e.key === 'm' || e.key === 'M' || e.code === 'KeyM')) {
      e.preventDefault();
      unmergeSelectedTableCells();
      return;
    }
  };
  document.removeEventListener('keydown', window._tableKeyDownHandler);
  window._tableKeyDownHandler = onDocKeyDown;
  document.addEventListener('keydown', onDocKeyDown);

  // ចាប់ព្រឹត្តិការណ៍បិទភ្ជាប់ (Paste) ទិន្នន័យពី Excel ចូលក្នុងតារាងផ្ទាល់
  printableArea.addEventListener('paste', handleTableClipboardPaste);

  // សម្រួលការចុច Tab / Shift+Tab ដើម្បីផ្លាស់ទីចន្លោះក្រឡាតារាង
  printableArea.addEventListener('keydown', (e) => {
    if (e.key === 'Tab') {
      const activeEl = document.activeElement;
      if (activeEl && activeEl.classList.contains('editable-cell')) {
        const cells = Array.from(printableArea.querySelectorAll('.doc-table tbody .editable-cell'));
        const currentIndex = cells.indexOf(activeEl);
        if (currentIndex !== -1) {
          e.preventDefault();
          if (!e.shiftKey) {
            if (currentIndex < cells.length - 1) {
              cells[currentIndex + 1].focus();
            } else {
              // ដល់ក្រឡាចុងក្រោយហើយ -> បន្ថែមជួរដេកថ្មីស្វ័យប្រវត្តិ
              addNewTableRow();
            }
          } else {
            if (currentIndex > 0) {
              cells[currentIndex - 1].focus();
            }
          }
        }
      }
    }
  });
}

/**
 * បង្ហាញសញ្ញាថាមានការកែប្រែមិនទាន់រក្សាទុក
 */
function markDocumentModified() {
  hasUnsavedChanges = true;
  const saveBtn = document.getElementById('btnSaveDocument');
  if (saveBtn) {
    saveBtn.classList.add('has-unsaved');
    saveBtn.innerHTML = '<i class="fas fa-save"></i> <span>រក្សាទុក *</span>';
  }

  const statusBadge = document.getElementById('modalSaveStatus');
  if (statusBadge) {
    statusBadge.className = 'doc-edit-status-badge status-unsaved';
    statusBadge.innerHTML = '<i class="fas fa-pen"></i> <span>មានការកែប្រែមិនទាន់រក្សាទុក</span>';
  }
}

/**
 * សម្អាតខ្សែអក្សររូបភាព Base64 ធំៗចេញពី LocalStorage ប្រសិនបើទំហំផ្ទុកជិតពេញ
 * (រូបភាពទាំងអស់មានរក្សាទុកក្នុង IndexedDB និង Cache រួចជាស្រេច ដូច្នេះការសម្អាតនេះមិនបាត់បង់ទិន្នន័យឡើយ)
 */
function cleanupLocalStorageEvidenceImages() {
  let freedCount = 0;
  try {
    const keys = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith('evidence_')) keys.push(k);
    }
    keys.forEach(k => {
      try {
        const val = localStorage.getItem(k);
        if (val && val.includes('data:image')) {
          const parsed = JSON.parse(val);
          if (Array.isArray(parsed)) {
            let modified = false;
            parsed.forEach(item => {
              if (item && item.imageData && item.imageData.length > 500) {
                delete item.imageData;
                item.hasIndexedDbMedia = true;
                modified = true;
                freedCount++;
              }
            });
            if (modified) {
              localStorage.setItem(k, JSON.stringify(parsed));
            }
          }
        }
      } catch (e) {}
    });
  } catch (e) {}
  return freedCount;
}

/**
 * មុខងាររក្សាទុកទិន្នន័យឯកសារ (Save to LocalStorage & IndexedDB)
 */
function saveCurrentDocument() {
  if (!currentActiveDoc) return;

  // បញ្ចប់ការ Edit នៅលើ Cell បច្ចុប្បន្នជាមុនសិន
  if (document.activeElement && typeof document.activeElement.blur === 'function') {
    try {
      document.activeElement.blur();
    } catch (e) {}
  }

  const pagesContainer = document.getElementById('docPagesContainer');
  const firstPage = document.getElementById('docFirstPage');
  const printableArea = document.getElementById('printableDocument');
  if (!printableArea) return;

  // ធ្វើបច្ចុប្បន្នកម្ម attributes លើ Dropdown និង Date Input មុនពេល serialize innerHTML
  const yearSelects = printableArea.querySelectorAll('.doc-year-select');
  yearSelects.forEach(yearSelect => {
    Array.from(yearSelect.options).forEach(opt => {
      if (opt.value === yearSelect.value) {
        opt.setAttribute('selected', 'selected');
      } else {
        opt.removeAttribute('selected');
      }
    });
  });
  const nativeDateInputs = printableArea.querySelectorAll('.native-date-input-hidden');
  nativeDateInputs.forEach(nativeDateInput => {
    if (nativeDateInput.value) {
      nativeDateInput.setAttribute('value', nativeDateInput.value);
    }
  });

  // រក្សាទុកគ្រប់ទំព័រនៃសន្លឹករបាយការណ៍ (docPagesContainer) ក្នុង LocalStorage ដោយមិនលាយជាមួយភស្តុតាង
  const htmlToSave = pagesContainer ? pagesContainer.innerHTML : (firstPage ? firstPage.innerHTML : printableArea.innerHTML);
  const timeFormatted = formatKhmerDateTime(new Date());

  let saveSuccess = false;

  try {
    localStorage.setItem('saved_doc_' + currentActiveDoc.code, htmlToSave);
    localStorage.setItem('saved_doc_time_' + currentActiveDoc.code, timeFormatted);
    saveSuccess = true;
  } catch (err) {
    console.warn('LocalStorage quota exceeded. Performing auto-cleanup of heavy evidence...', err);
    try {
      cleanupLocalStorageEvidenceImages();
      localStorage.setItem('saved_doc_' + currentActiveDoc.code, htmlToSave);
      localStorage.setItem('saved_doc_time_' + currentActiveDoc.code, timeFormatted);
      saveSuccess = true;
    } catch (retryErr) {
      console.warn('LocalStorage still full, storing document in IndexedDB cache:', retryErr);
      if (typeof saveEvidenceToDB === 'function') {
        saveEvidenceToDB('doc_' + currentActiveDoc.code, [{ html: htmlToSave, time: timeFormatted }]);
      }
      saveSuccess = true;
    }
  }

  if (saveSuccess) {
    hasUnsavedChanges = false;

    // Update Save button UI
    const saveBtn = document.getElementById('btnSaveDocument');
    if (saveBtn) {
      saveBtn.classList.remove('has-unsaved');
      saveBtn.innerHTML = '<i class="fas fa-check"></i> <span>បានរក្សាទុក!</span>';
      setTimeout(() => {
        if (saveBtn && !hasUnsavedChanges) {
          saveBtn.innerHTML = '<i class="fas fa-save"></i> <span>រក្សាទុក</span>';
        }
      }, 2400);
    }

    // Update Status Badge UI
    const statusBadge = document.getElementById('modalSaveStatus');
    if (statusBadge) {
      statusBadge.className = 'doc-edit-status-badge status-saved';
      statusBadge.innerHTML = `<i class="fas fa-check-circle"></i> <span>បានរក្សាទុក (${timeFormatted})</span>`;
    }

    // Update Card in main view
    updateCardBadge(currentActiveDoc.code, true);

    showToast(`បានរក្សាទុកទិន្នន័យឯកសារ [${currentActiveDoc.code}] ជោគជ័យ!`);
  } else {
    alert('មិនអាចរក្សាទុកឯកសារបានឡើយ ដោយសារទំហំផ្ទុករបស់កម្មវិធីរុករកពេញ។');
  }
}

/**
 * មុខងារកំណត់ទម្រង់ឯកសារទៅជាលំនាំដើមវិញ (Reset Document to Default)
 */
function resetCurrentDocument() {
  if (!currentActiveDoc) return;

  const hasSaved = localStorage.getItem('saved_doc_' + currentActiveDoc.code);
  if (!hasSaved && !hasUnsavedChanges) {
    showToast('ឯកសារនេះស្ថិតក្នុងទម្រង់គំរូលំនាំដើមស្រាប់ហើយ');
    return;
  }

  const confirmReset = confirm(`តើលោកអ្នកពិតជាចង់កំណត់ទម្រង់ឯកសារ [${currentActiveDoc.code}] នេះត្រឡប់ទៅជាទម្រង់គំរូលំនាំដើមវិញមែនទេ? រាល់ការកែប្រែទាំងអស់នឹងត្រូវសម្អាត។`);
  if (!confirmReset) return;

  localStorage.removeItem('saved_doc_' + currentActiveDoc.code);
  localStorage.removeItem('saved_doc_time_' + currentActiveDoc.code);
  hasUnsavedChanges = false;

  // Render fresh default template
  const sub = currentActiveDoc;
  const evList = getEvidenceList(sub.code).filter(item => item && (item.imageData || hasLinkUrl(item) || (item.caption && item.caption.trim())));
  const hasEvidence = evList.length > 0;

  const modalContainer = document.getElementById('documentSheetContainer');
  const bannerHtml = `
    <div class="doc-edit-banner no-print">
      <div class="banner-text">
        <i class="fas fa-edit"></i>
        <span>ទម្រង់កែសម្រួលផ្ទាល់៖ លោកអ្នកអាចចុចលើទិន្នន័យក្នុងតារាង ឬអត្ថបទនានាដើម្បីកែប្រែ រួចចុចប៊ូតុង <strong>«រក្សាទុក»</strong> នៅខាងលើ។</span>
      </div>
      <div id="modalSaveStatus" class="doc-edit-status-badge status-default">
        <i class="fas fa-circle-info"></i>
        <span>ទម្រង់គំរូលំនាំដើម</span>
      </div>
    </div>
  `;

  const fullSheetHtml = `
    <div class="doc-pages-container" id="docPagesContainer">
      <div class="doc-report-page" id="docPage_1" data-page-index="1">
        ${buildDefaultDocumentHtml(sub)}
      </div>
    </div>
    <div class="add-sheet-action-wrap no-print">
      <button type="button" class="btn-add-sheet-primary" onclick="addNewDocumentSheetPage()" title="បន្ថែមសន្លឹករបាយការណ៍ថ្មី (ទំព័រទី...)">
        <i class="fas fa-file-circle-plus"></i> <span>+ បន្ថែមសន្លឹករបាយការណ៍ថ្មី (ទំព័រទី...)</span>
      </button>
    </div>
    <div class="evidence-page-separator no-print">
      <div class="separator-line"></div>
      <div class="separator-pill">
        <i class="fas fa-images"></i> ឧបសម្ព័ន្ធភស្តុតាងបញ្ជាក់បន្ថែម (ស្ថិតនៅបន្តបន្ទាប់ពីទំព័ររបាយការណ៍) ${hasEvidence ? `(${khmerNumber(evList.length)})` : ''}
      </div>
      <div class="separator-line"></div>
    </div>
    <div class="document-page-2 ${hasEvidence ? '' : 'has-no-evidence'}" id="docEvidenceAnnex">
      ${renderEvidenceAnnexHtml(sub)}
    </div>
  `;

  modalContainer.innerHTML = bannerHtml + `<div class="document-sheet" id="printableDocument">${fullSheetHtml}</div>`;

  // ពិនិត្យចំនួនជួរឈរដើម្បីកំណត់ទិសដៅក្រដាស និងទំហំសមាមាត្រកុំឱ្យដាច់ពេលបោះពុម្ព
  const docTable = modalContainer.querySelector('.doc-table');
  if (docTable) {
    const dataHeaders = docTable.querySelectorAll('thead th:not(.row-action-col)');
    const colCount = dataHeaders.length;
    if (colCount >= 10) {
      docTable.classList.add('table-ultra-wide');
      docTable.classList.remove('table-wide');
      setPrintOrientation('landscape');
    } else if (colCount >= 7) {
      docTable.classList.add('table-wide');
      docTable.classList.remove('table-ultra-wide');
      setPrintOrientation('landscape');
    } else {
      docTable.classList.remove('table-wide', 'table-ultra-wide');
      setPrintOrientation('portrait');
    }
  }

  initTableEditingFeatures();

  // Reset Save button UI
  const saveBtn = document.getElementById('btnSaveDocument');
  if (saveBtn) {
    saveBtn.classList.remove('has-unsaved');
    saveBtn.innerHTML = '<i class="fas fa-save"></i> <span>រក្សាទុក</span>';
  }

  // Update Card in main view
  updateCardBadge(currentActiveDoc.code, false);

  showToast(`បានកំណត់ទម្រង់ឯកសារ ${currentActiveDoc.code} ទៅជាលំនាំដើមវិញជោគជ័យ!`);
}

/**
 * បង្កើតសន្លឹករបាយការណ៍ថ្មីបន្ថែម (Multi-Page Report Sheets)
 * - រក្សាទុកក្បាលសំបុត្រ និងព័ត៌មានគោល (អត្ថន័យក្នុងស៊ុំពណ៌ខៀវផ្នែកខាងលើ)
 * - រក្សាទុកក្បាលតារាង (<thead>) ឱ្យនៅដដែល ១០០% ទាំងឈ្មោះកូឡោន និងទទឹង
 * - បង្កើតជួរដេកទិន្នន័យថ្មីស្អាតក្នុង <tbody> ដោយរត់លេខរៀង (ល.រ) បន្តពីសន្លឹកមុន
 * - រក្សាទុកផ្នែកហត្ថលេខា (អត្ថន័យក្នុងស៊ុំពណ៌ខៀវផ្នែកខាងក្រោម) ឱ្យនៅខាងក្រោមនៃសន្លឹក
 * - ឧបសម្ព័ន្ធភស្តុតាងនៅតែបន្តស្ថិតនៅខាងក្រោមបង្អស់នៃសន្លឹករបាយការណ៍ទាំងអស់ជានិច្ច
 */
function addNewDocumentSheetPage() {
  const printableArea = document.getElementById('printableDocument');
  if (!printableArea) return;

  const pagesContainer = document.getElementById('docPagesContainer') || printableArea.querySelector('.doc-first-page') || printableArea;
  if (!pagesContainer) return;

  // ស្វែងរកសន្លឹកទី១ ដើម្បីចម្លងក្បាលសំបុត្រ ក្បាលតារាង និងហត្ថលេខា
  let page1 = pagesContainer.querySelector('.doc-report-page');

  // ប្រសិនបើសន្លឹកទី១ មិនទាន់មាន class .doc-report-page សូមរុំវា
  if (!page1) {
    const currentInner = pagesContainer.innerHTML;
    pagesContainer.innerHTML = `<div class="doc-report-page" id="docPage_1" data-page-index="1">${currentInner}</div>`;
    page1 = pagesContainer.querySelector('#docPage_1');
  }

  const allCurrentPages = Array.from(pagesContainer.querySelectorAll('.doc-report-page'));
  const newPageNum = allCurrentPages.length + 1;

  // ១. ចម្លងក្បាលសំបុត្រ និងព័ត៌មានគោលពីសន្លឹកទី១ (Header & Meta - ស៊ុំខៀវខាងលើ)
  const badgeEl = page1.querySelector('.indicator-badge-top-right');
  const badgeHtml = badgeEl ? badgeEl.outerHTML : '';

  const headerEl = page1.querySelector('.doc-header');
  const headerHtml = headerEl ? headerEl.outerHTML : '';

  const titleEl = page1.querySelector('.doc-title-container');
  const titleHtml = titleEl ? titleEl.outerHTML : '';

  const metaEl = page1.querySelector('.doc-meta-info');
  const metaHtml = metaEl ? metaEl.outerHTML : '';

  // ២. ចម្លងក្បាលតារាង (<thead>) ពីសន្លឹកទី១
  const table1 = page1.querySelector('.doc-table');
  let theadHtml = '';
  let colCount = 5;
  let tableClasses = 'doc-table';

  if (table1) {
    const thead1 = table1.querySelector('thead');
    if (thead1) {
      const clonedThead = thead1.cloneNode(true);
      clonedThead.querySelectorAll('.col-resizer, .col-action-btns').forEach(el => el.remove());
      theadHtml = clonedThead.innerHTML;
      colCount = clonedThead.querySelectorAll('th:not(.row-action-col)').length;
    }
    tableClasses = table1.className;
  }

  // គណនាផលបូកជួរដេកពីសន្លឹកមុនៗទាំងអស់ ដើម្បីរត់លេខរៀង (ល.រ) បន្ត
  let totalPrevRows = 0;
  allCurrentPages.forEach(p => {
    totalPrevRows += p.querySelectorAll('tbody tr').length;
  });

  // ពិនិត្យមើលថាតើជួរឈរទី១ ជាលេខរៀង (ល.រ / No) ឬទេ
  let isCol0Index = true;
  if (table1) {
    const firstTh = table1.querySelector('thead th:not(.row-action-col)');
    if (firstTh) {
      const txt = firstTh.textContent.trim();
      isCol0Index = Boolean(/^(ល\.?រ|no|n°|លរ|#)/i.test(txt));
    }
  }

  // បង្កើតជួរដេកទទេស្អាតចំនួន ៣ ជួរសម្រាប់សន្លឹកថ្មី
  let newTbodyHtml = '';
  for (let r = 1; r <= 3; r++) {
    newTbodyHtml += '<tr>';
    for (let c = 0; c < colCount; c++) {
      const isCenter = (c === 0 && isCol0Index);
      const val = isCenter ? khmerNumber(totalPrevRows + r) : '';
      newTbodyHtml += `<td class="${isCenter ? 'text-center ' : ''}editable-cell" contenteditable="true">${val}</td>`;
    }
    newTbodyHtml += `<td class="row-action-col no-print">${getRowActionButtonsHtml()}</td>`;
    newTbodyHtml += '</tr>';
  }

  // ៣. ចម្លងសេចក្តីណែនាំ និងហត្ថលេខា (Signatures - ស៊ុំខៀវខាងក្រោម)
  const guidelineEl = page1.querySelector('div[style*="border-left"]');
  const guidelineHtml = guidelineEl ? guidelineEl.outerHTML : `
    <div class="editable" contenteditable="true" style="margin: 1.25rem 0; font-size: 0.84rem; color: #475569; background: #f8fafc; padding: 0.75rem 1rem; border-left: 3px solid #0284c7; border-radius: 4px;">
      <strong>* សេចក្តីណែនាំក្នុងការបំពេញ៖</strong> ទម្រង់ឯកសារនេះត្រូវបានបង្កើតឡើងដោយផ្អែកលើកូឡោនទី២ (ទំព័រ១១-៣៥) នៃសៀវភៅស្តង់ដាសាលាបឋមសិក្សាគំរូ។ លោកគ្រូ-អ្នកគ្រូ និងគណៈគ្រប់គ្រងសាលាអាចចុចកែសម្រួលទិន្នន័យផ្ទាល់លើតារាងខាងលើនេះរួចចុចប៊ូតុង «រក្សាទុក» បោះពុម្ព ឬទាញយកជាឯកសារ Word ដើម្បីចុះហត្ថលេខា និងរក្សាទុកជាភស្តុតាង។
    </div>
  `;

  const sigEl = page1.querySelector('.doc-signatures');
  const sigHtml = sigEl ? sigEl.outerHTML : '';

  // ៤. បង្កើត Container នៃសន្លឹករបាយការណ៍ថ្មី
  const pageDiv = document.createElement('div');
  pageDiv.className = 'doc-report-page';
  pageDiv.id = `docPage_${newPageNum}`;
  pageDiv.setAttribute('data-page-index', String(newPageNum));

  pageDiv.innerHTML = `
    <div class="doc-page-divider no-print">
      <div class="divider-line"></div>
      <div class="divider-pill">
        <i class="fas fa-scroll"></i> ទំព័រទី ${khmerNumber(newPageNum)} នៃសន្លឹករបាយការណ៍
      </div>
      <div class="divider-line"></div>
    </div>

    <div class="doc-page-header-bar no-print">
      <div class="doc-page-badge">
        <i class="fas fa-file-alt" style="color: #2563eb;"></i> <span>សន្លឹករបាយការណ៍ - ទំព័រទី ${khmerNumber(newPageNum)}</span>
      </div>
      <button type="button" class="btn-delete-page" onclick="deleteDocumentSheetPage(this.closest('.doc-report-page'))" title="លុបសន្លឹកទី ${khmerNumber(newPageNum)} នេះចេញ">
        <i class="fas fa-trash-alt"></i> <span>លុបសន្លឹកនេះ (ទំព័រទី ${khmerNumber(newPageNum)})</span>
      </button>
    </div>

    ${badgeHtml}
    ${headerHtml}
    ${titleHtml}
    ${metaHtml}

    <table class="${tableClasses}">
      <thead>
        ${theadHtml}
      </thead>
      <tbody>
        ${newTbodyHtml}
      </tbody>
    </table>

    <div class="table-actions-bar no-print">
      ${getTableActionsBarHtml()}
    </div>

    ${guidelineHtml}
    ${sigHtml}
  `;

  pagesContainer.appendChild(pageDiv);

  // ៥. បំពាក់ Resizers, Action Buttons និង Listeners លើតារាងថ្មី
  const newTable = pageDiv.querySelector('.doc-table');
  if (newTable) {
    ensureColumnActionButtonsOnHeaders(newTable);
    initTableColumnResizers();
  }

  initTableEditingFeatures();
  markDocumentModified();

  // ៦. Scroll ទៅកាន់សន្លឹកថ្មី
  setTimeout(() => {
    pageDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 100);

  showToast(`⚡ បានបង្កើតសន្លឹករបាយការណ៍ថ្មី (ទំព័រទី ${khmerNumber(newPageNum)}) ជោគជ័យ!`);
}

/**
 * លុបសន្លឹករបាយការណ៍ណាមួយចេញ
 */
function deleteDocumentSheetPage(pageEl) {
  if (!pageEl) return;
  const pagesContainer = document.getElementById('docPagesContainer');
  if (!pagesContainer) return;

  const allPages = Array.from(pagesContainer.querySelectorAll('.doc-report-page'));
  if (allPages.length <= 1) {
    alert('មិនអាចលុបសន្លឹកទី១ (សន្លឹកមេ) បានឡើយ!');
    return;
  }

  const pageIdx = allPages.indexOf(pageEl) + 1;
  const shouldDelete = confirm(`តើលោកគ្រូ-អ្នកគ្រូពិតជាចង់លុបសន្លឹករបាយការណ៍ទី ${khmerNumber(pageIdx)} នេះមែនទេ? រាល់ទិន្នន័យក្នុងសន្លឹកនេះនឹងត្រូវលុបចេញ!`);
  if (!shouldDelete) return;

  pageEl.remove();
  refreshDocumentPageNumbers();
  markDocumentModified();
  showToast(`បានលុបសន្លឹករបាយការណ៍ទី ${khmerNumber(pageIdx)} ជោគជ័យ!`);
}

/**
 * ធ្វើបច្ចុប្បន្នកម្មលេខទំព័រ និងក្បាលសន្លឹកឡើងវិញបន្ទាប់ពីមានការលុបសន្លឹក
 */
function refreshDocumentPageNumbers() {
  const pagesContainer = document.getElementById('docPagesContainer');
  if (!pagesContainer) return;

  const allPages = Array.from(pagesContainer.querySelectorAll('.doc-report-page'));
  allPages.forEach((p, idx) => {
    const pageNum = idx + 1;
    p.id = `docPage_${pageNum}`;
    p.setAttribute('data-page-index', String(pageNum));

    const pill = p.querySelector('.doc-page-divider .divider-pill');
    if (pill) {
      pill.innerHTML = `<i class="fas fa-scroll"></i> ទំព័រទី ${khmerNumber(pageNum)} នៃសន្លឹករបាយការណ៍`;
    }

    const badge = p.querySelector('.doc-page-badge span');
    if (badge) {
      badge.textContent = `សន្លឹករបាយការណ៍ - ទំព័រទី ${khmerNumber(pageNum)}`;
    }

    const delBtn = p.querySelector('.btn-delete-page');
    if (delBtn) {
      delBtn.title = `លុបសន្លឹកទី ${khmerNumber(pageNum)} នេះចេញ`;
      delBtn.innerHTML = `<i class="fas fa-trash-alt"></i> <span>លុបសន្លឹកនេះ (ទំព័រទី ${khmerNumber(pageNum)})</span>`;
    }

    const table = p.querySelector('.doc-table');
    if (table) {
      reindexTableRows(table);
      ensureColumnActionButtonsOnHeaders(table);
    }
  });

  initTableColumnResizers();
}

/**
 * ពិនិត្យមើលថាតើជួរឈរទី១ នៃតារាង ជាជួរឈរលេខរៀង (ល.រ) ដែរឬទេ
 */
function isRowIndexColumn(table) {
  if (!table) return false;
  const docTable = table.closest('table') || table;
  const theadTh = docTable.querySelector('thead tr th:not(.row-action-col)');
  if (theadTh) {
    const headerText = theadTh.innerText.replace(/[\r\n\t\s]+/g, ' ').trim();
    if (/^(ល\s*\.?\s*រ|លេខ\s*រៀង|លេខ\s*សម្គាល់|no\.?|n°|№|#)$/i.test(headerText) || 
        headerText.includes('ល.រ') || 
        headerText.includes('លរ') ||
        headerText.includes('លេខរៀង')) {
      return true;
    }
  }

  // ពិនិត្យទិន្នន័យជាក់ស្តែងក្នុងក្រឡាជួរទី១ នៃជួរដេកទាំងអស់
  const rows = docTable.querySelectorAll('tbody tr');
  let numericCount = 0;
  let totalNonEmpty = 0;
  rows.forEach(r => {
    const firstCell = r.querySelector('td:not(.row-action-col)');
    if (firstCell) {
      const txt = firstCell.innerText.trim();
      if (txt) {
        totalNonEmpty++;
        if (/^[0-9០-៩]+$/.test(txt)) {
          numericCount++;
        }
      }
    }
  });

  if (totalNonEmpty > 0 && (numericCount / totalNonEmpty >= 0.5)) {
    return true;
  }
  return false;
}

/**
 * រៀបចំលេខរៀងឡើងវិញក្នុងជួរឈរទី១ ឱ្យរត់ជាលេខខ្មែរស្វ័យប្រវត្តិ (១, ២, ៣...)
 */
function reindexTableRows(table) {
  if (!table) return;
  const tbody = table.tagName === 'TBODY' ? table : table.querySelector('tbody');
  if (!tbody) return;

  if (!isRowIndexColumn(tbody)) return;

  const docTable = tbody.closest('table') || tbody;
  const { grid } = getTableGridMatrix(docTable);

  let startCounter = 1;
  const pagesContainer = document.getElementById('docPagesContainer');
  if (pagesContainer) {
    const allTables = Array.from(pagesContainer.querySelectorAll('.doc-table'));
    const currentTableIdx = allTables.indexOf(docTable);
    if (currentTableIdx > 0) {
      for (let i = 0; i < currentTableIdx; i++) {
        startCounter += allTables[i].querySelectorAll('tbody tr').length;
      }
    }
  }

  let counter = startCounter;
  for (let r = 0; r < grid.length; r++) {
    const item = grid[r] && grid[r][0];
    if (item && item.isOrigin && item.cell && item.cell.tagName === 'TD' && !item.cell.classList.contains('row-action-col')) {
      item.cell.textContent = khmerNumber(counter++);
      item.cell.classList.add('text-center');
    }
  }
}

/**
 * បន្ថែមជួរដេកថ្មីបន្ទាប់ពីជួរដេកដែលបានចុច (+)
 */
function insertTableRowAfter(btn) {
  const tr = btn.closest('tr');
  if (!tr) return;

  const tbody = tr.closest('tbody');
  if (!tbody) return;

  const table = tbody.closest('table');
  const { maxCols } = getTableGridMatrix(table);
  const colCount = maxCols || 7;

  const newTr = document.createElement('tr');

  for (let i = 0; i < colCount; i++) {
    const td = document.createElement('td');
    const isCenter = (i === 0);
    td.className = (isCenter ? 'text-center ' : '') + 'editable-cell';
    td.setAttribute('contenteditable', 'true');
    td.textContent = '';
    newTr.appendChild(td);
  }

  // Row action buttons
  const actionTd = document.createElement('td');
  actionTd.className = 'row-action-col no-print';
  actionTd.innerHTML = getRowActionButtonsHtml();
  newTr.appendChild(actionTd);

  // បញ្ចូលជួរដេកថ្មីនៅបន្ទាប់ពីជួរដេកបច្ចុប្បន្នភ្លាមៗ
  if (tr.nextSibling) {
    tbody.insertBefore(newTr, tr.nextSibling);
  } else {
    tbody.appendChild(newTr);
  }

  // ធ្វើបច្ចុប្បន្នកម្មលេខរៀងគ្រប់ជួរដេកឡើងវិញដោយស្វ័យប្រវត្តិ
  reindexTableRows(tbody);

  markDocumentModified();
  showToast('បានបន្ថែមជួរដេកថ្មីនៅចន្លោះនេះជោគជ័យ!');

  // កំណត់ Focus ទៅកាន់ក្រឡាទិន្នន័យទី២ នៃជួរដេកថ្មី (ដើម្បីចាប់ផ្តើមវាយអក្សរ)
  const cells = newTr.querySelectorAll('td.editable-cell');
  if (cells.length > 1) {
    cells[1].focus();
  } else if (cells.length > 0) {
    cells[0].focus();
  }
}

/**
 * បន្ថែមជួរដេកថ្មីខាងក្រោមបង្អស់នៃតារាង
 */
function addNewTableRow() {
  let tbody = null;
  if (window._lastActiveCell) {
    const table = window._lastActiveCell.closest('.doc-table');
    if (table) tbody = table.querySelector('tbody');
  }
  if (!tbody) {
    const allTbodies = document.querySelectorAll('#printableDocument .doc-table tbody');
    if (allTbodies.length > 0) {
      tbody = allTbodies[allTbodies.length - 1];
    }
  }
  if (!tbody) return;

  const docTable = tbody.closest('table') || tbody;
  const { maxCols } = getTableGridMatrix(docTable);
  const colCount = maxCols || 7;

  const tr = document.createElement('tr');
  
  for (let i = 0; i < colCount; i++) {
    const td = document.createElement('td');
    const isCenter = (i === 0);
    td.className = (isCenter ? 'text-center ' : '') + 'editable-cell';
    td.setAttribute('contenteditable', 'true');
    td.textContent = '';
    tr.appendChild(td);
  }

  // Row action buttons in its own column
  const actionTd = document.createElement('td');
  actionTd.className = 'row-action-col no-print';
  actionTd.innerHTML = getRowActionButtonsHtml();
  tr.appendChild(actionTd);

  tbody.appendChild(tr);
  reindexTableRows(tbody);
  markDocumentModified();

  // Focus the second cell of the newly added row
  const cells = tr.querySelectorAll('td.editable-cell');
  if (cells.length > 1) {
    cells[1].focus();
  } else if (cells.length > 0) {
    cells[0].focus();
  }
}

/**
 * លុបជួរដេកណាមួយចេញពីតារាងទាំងស្រុង ដោយដោះស្រាយ rowspan និងសម្រួលគ្រប់ជួរដេក
 * ធានាថាមិនលុបក្បាលតារាង thead ឡើយ និងរក្សាជួរដេកទិន្នន័យយ៉ាងហោចណាស់មួយ
 */
function deleteRowAt(targetTr, confirmPrompt = false) {
  if (!targetTr) return false;

  // ការពារមិនឱ្យលុបក្បាលតារាង thead ដាច់ខាត
  if (targetTr.closest('thead')) {
    showToast('មិនអាចលុបក្បាលតារាង (Header) បានឡើយ!');
    return false;
  }

  const tbody = targetTr.closest('tbody') || targetTr.parentElement;
  if (!tbody) return false;

  const rows = Array.from(tbody.querySelectorAll('tr'));
  if (rows.length <= 1) {
    showToast('មិនអាចលុបបានទេ ត្រូវទុកយ៉ាងហោចណាស់មួយជួរដេក។');
    return false;
  }

  const table = targetTr.closest('table');
  const rowNum = rows.indexOf(targetTr) + 1;

  if (confirmPrompt) {
    const confirmDel = confirm(`តើលោកអ្នកពិតជាចង់លុបជួរដេកទី ${khmerNumber(rowNum)} នេះចេញពីតារាងមែនទេ?\nរាល់ទិន្នន័យក្នុងជួរដេកនេះនឹងត្រូវបាត់បង់។`);
    if (!confirmDel) return false;
  }

  if (table) {
    const { grid, rows: allRows, maxCols } = getTableGridMatrix(table);
    const rIdx = allRows.indexOf(targetTr);

    if (rIdx >= 0) {
      const nextTr = allRows[rIdx + 1];
      const processedCells = new Set();

      for (let c = 0; c < maxCols; c++) {
        const item = grid[rIdx] && grid[rIdx][c];
        if (!item || !item.cell || processedCells.has(item.cell)) continue;
        processedCells.add(item.cell);

        const cell = item.cell;
        const rowSpan = item.rowSpan || 1;

        if (rowSpan > 1) {
          if (item.originR === rIdx) {
            // ក្រឡាផ្តើមចេញពីជួរដេកនេះ៖ ផ្ទេរទៅកាន់ជួរដេកបន្ទាប់ (nextTr)
            if (nextTr) {
              const newRowSpan = rowSpan - 1;
              if (newRowSpan <= 1) {
                cell.removeAttribute('rowspan');
              } else {
                cell.setAttribute('rowspan', newRowSpan);
              }

              // រកក្រឡា anchor បន្ទាប់ពី originC ក្នុង nextTr ដើម្បីដាក់ចូល
              let anchorCell = null;
              for (let nextC = item.originC + (item.colSpan || 1); nextC < maxCols; nextC++) {
                const nextItem = grid[rIdx + 1] && grid[rIdx + 1][nextC];
                if (nextItem && nextItem.isOrigin && nextItem.cell && nextItem.cell.parentNode === nextTr) {
                  anchorCell = nextItem.cell;
                  break;
                }
              }
              if (!anchorCell) {
                anchorCell = nextTr.querySelector('.row-action-col');
              }

              if (anchorCell) {
                nextTr.insertBefore(cell, anchorCell);
              } else {
                nextTr.appendChild(cell);
              }
            }
          } else {
            // ក្រឡាផ្តើមចេញពីជួរដេកខាងលើ៖ កាត់បន្ថយ rowspan ចុះ ១
            const newRowSpan = rowSpan - 1;
            if (newRowSpan <= 1) {
              cell.removeAttribute('rowspan');
            } else {
              cell.setAttribute('rowspan', newRowSpan);
            }
          }
        }
      }
    }
  }

  // សម្អាត Selection ប្រសិនបើមានក្រឡាក្នុងជួរដេកនេះ
  if (currentSelectedCells && currentSelectedCells.size > 0) {
    Array.from(currentSelectedCells).forEach(c => {
      if (targetTr.contains(c)) {
        currentSelectedCells.delete(c);
      }
    });
    if (currentSelectedCells.size <= 1) {
      removeFloatingMergeToolbar();
    }
  }

  if (window._lastActiveCell && targetTr.contains(window._lastActiveCell)) {
    window._lastActiveCell = null;
  }

  targetTr.remove();

  // រៀបចំលេខរៀងឡើងវិញក្នុងជួរទី១ ឱ្យរត់តាមលំដាប់ជាលេខខ្មែរស្វ័យប្រវត្តិ
  reindexTableRows(tbody);

  // ធានាប៊ូតុងសកម្មភាពនៅលើជួរដេកទាំងអស់
  ensureRowActionButtons(table);

  markDocumentModified();
  showToast(`បានលុបជួរដេកទី ${khmerNumber(rowNum)} ជោគជ័យ!`);
  return true;
}

/**
 * លុបជួរដេកដោយឆ្លាតវៃ៖ ប្រសិនបើមានក្រឡា/ជួរដេកកំពុងជ្រើស លុបជួរដេកនោះ!
 * ប្រសិនបើគ្មានជួរដេកជ្រើស លុបជួរដេកចុងក្រោយបង្អស់
 */
function deleteTableRowSmart() {
  let table = null;
  const activeCell = window._lastActiveCell;
  if (activeCell) {
    table = activeCell.closest('.doc-table');
  }
  if (!table && currentSelectedCells && currentSelectedCells.size > 0) {
    const firstSelected = Array.from(currentSelectedCells)[0];
    if (firstSelected) table = firstSelected.closest('.doc-table');
  }
  if (!table) {
    const allTables = document.querySelectorAll('#printableDocument .doc-table');
    if (allTables.length > 0) {
      table = allTables[allTables.length - 1];
    }
  }
  if (!table) return;

  const tbody = table.querySelector('tbody');
  if (!tbody) return;

  let targetTr = null;

  // 1. ពិនិត្យក្រឡាដែលបានចុចចុងក្រោយ (last active cell)
  if (activeCell && table.contains(activeCell)) {
    targetTr = activeCell.closest('tbody tr');
  }

  // 2. ពិនិត្យក្រឡាដែលបានជ្រើសក្នុង selection
  if (!targetTr && currentSelectedCells && currentSelectedCells.size > 0) {
    const firstSelected = Array.from(currentSelectedCells)[0];
    if (firstSelected && table.contains(firstSelected)) {
      targetTr = firstSelected.closest('tbody tr');
    }
  }

  if (targetTr) {
    deleteRowAt(targetTr, true);
  } else {
    const rows = Array.from(tbody.querySelectorAll('tr'));
    if (rows.length <= 1) {
      showToast('មិនអាចលុបបានទៀតទេ ត្រូវទុកយ៉ាងហោចណាស់មួយជួរដេក។');
      return;
    }
    const lastRow = rows[rows.length - 1];
    deleteRowAt(lastRow, true);
  }
}

/**
 * លុបជួរដេកចុងក្រោយនៃតារាង (ឬជួរដេកដែលកំពុង active)
 */
function deleteLastTableRow() {
  deleteTableRowSmart();
}

/**
 * លុបជួរដេកណាមួយដែលបានចុចប៊ូតុងលុប (×)
 */
function deleteCurrentTableRow(btn) {
  const tr = btn.closest('tr');
  if (!tr) return;
  deleteRowAt(tr, false);
}

/**
 * ធានាថារាល់ជួរដេកទាំងអស់នៃតារាងមានប៊ូតុងបន្ថែម (+) និងលុប (×) គ្រប់គ្រាន់
 */
function ensureRowActionButtons(table) {
  if (!table) table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const tbody = table.querySelector('tbody');
  if (tbody) {
    tbody.querySelectorAll('tr').forEach(tr => {
      let actionTd = tr.querySelector('.row-action-col');
      if (!actionTd) {
        actionTd = document.createElement('td');
        actionTd.className = 'row-action-col no-print';
        tr.appendChild(actionTd);
      }
      if (!actionTd.querySelector('.row-insert-btn')) {
        actionTd.innerHTML = getRowActionButtonsHtml();
      }
    });
  }

  // ធានាក្បាលតារាងមានជួរឈរសកម្មភាព
  const theadTr = table.querySelector('thead tr');
  if (theadTr) {
    let actionTh = theadTr.querySelector('.row-action-col');
    if (!actionTh) {
      actionTh = document.createElement('th');
      actionTh.className = 'row-action-col no-print';
      theadTr.appendChild(actionTh);
    }
    actionTh.style.width = '58px';
    actionTh.style.border = 'none';
    actionTh.style.background = 'transparent';
  }
}

/**
 * បន្ថែមជួរឈរថ្មីក្នុងតារាង
 */
function addNewTableColumn() {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const theadRow = table.querySelector('thead tr');
  if (!theadRow) return;

  const actionTh = theadRow.querySelector('.row-action-col');
  const currentDataHeaders = theadRow.querySelectorAll('th:not(.row-action-col)');
  const nextColNum = currentDataHeaders.length + 1;

  const colName = prompt(`សូមបញ្ចូលចំណងជើងជួរឈរថ្មី (កូឡោនទី ${khmerNumber(nextColNum)})៖`, `កូឡោនថ្មី ${khmerNumber(nextColNum)}`);
  if (colName === null) return; // បោះបង់
  const finalColName = colName.trim() || `កូឡោនថ្មី ${khmerNumber(nextColNum)}`;

  // បង្កើតក្បាលជួរឈរ <th> ថ្មី
  const newTh = document.createElement('th');
  newTh.className = 'editable-header';
  newTh.setAttribute('contenteditable', 'true');
  newTh.textContent = finalColName;

  if (actionTh) {
    theadRow.insertBefore(newTh, actionTh);
  } else {
    theadRow.appendChild(newTh);
  }

  // បន្ថែមក្រឡា <td> ថ្មីក្នុងគ្រប់ជួរដេកទាំងអស់នៃ <tbody>
  const tbodyRows = table.querySelectorAll('tbody tr');
  tbodyRows.forEach(row => {
    const actionTd = row.querySelector('.row-action-col');
    const newTd = document.createElement('td');
    newTd.className = 'editable-cell';
    newTd.setAttribute('contenteditable', 'true');
    newTd.textContent = '';

    if (actionTd) {
      row.insertBefore(newTd, actionTd);
    } else {
      row.appendChild(newTd);
    }
  });

  markDocumentModified();
  initTableColumnResizers();
  ensureColumnActionButtonsOnHeaders(table);

  // ពិនិត្យចំនួនជួរឈរដើម្បីកំណត់ table-wide និង landscape
  const allDataHeaders = theadRow.querySelectorAll('th:not(.row-action-col)');
  if (allDataHeaders.length >= 10) {
    table.classList.add('table-ultra-wide');
    table.classList.remove('table-wide');
    setPrintOrientation('landscape');
  } else if (allDataHeaders.length >= 7) {
    table.classList.add('table-wide');
    table.classList.remove('table-ultra-wide');
    setPrintOrientation('landscape');
  }

  showToast(`បានបន្ថែមជួរឈរ «${finalColName}» ជោគជ័យ!`);

  // Focus លើក្បាលជួរឈរថ្មីដើម្បីឱ្យអ្នកប្រើអាចចុចកែសម្រួលឈ្មោះបានភ្លាមៗ
  newTh.focus();
}

/**
 * ស្រង់យកឈ្មោះចំណងជើងជួរឈរដោយកាត់ចោលធាតុ UI (UI Buttons & Resizers)
 */
function getColumnHeaderTitle(th) {
  if (!th) return '';
  const clone = th.cloneNode(true);
  clone.querySelectorAll('.no-print, .col-resizer, .col-action-btns').forEach(el => el.remove());
  return clone.innerText.replace(/[\r\n\t\s]+/g, ' ').trim();
}

/**
 * បំពាក់ប៊ូតុងបន្ថែម (+) និងលុប (×) លើគ្រប់ក្បាលជួរឈរទាំងអស់ (Column Action Buttons)
 */
function ensureColumnActionButtonsOnHeaders(table) {
  if (!table) table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const theadRow = table.querySelector('thead tr');
  if (!theadRow) return;

  const headers = Array.from(theadRow.querySelectorAll('th:not(.row-action-col)'));
  const isFirstColIndex = isRowIndexColumn(table);

  headers.forEach((th, idx) => {
    // លុប col-action-btns ចាស់ចេញប្រសិនបើមាន ដើម្បីកុំឱ្យស្ទួន
    const existing = th.querySelector('.col-action-btns');
    if (existing) existing.remove();

    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'col-action-btns no-print';
    actionsDiv.setAttribute('contenteditable', 'false');

    // ប៊ូតុងបន្ថែមជួរឈរខាងស្តាំនេះ (+)
    const insertBtn = document.createElement('button');
    insertBtn.className = 'col-insert-btn';
    insertBtn.type = 'button';
    insertBtn.title = 'បន្ថែមជួរឈរថ្មីខាងស្តាំជួរឈរនេះ (+)';
    insertBtn.innerHTML = '<i class="fas fa-plus"></i>';
    actionsDiv.appendChild(insertBtn);

    // ប៊ូតុងលុបជួរឈរ (×) (លើកលែងជួរឈរទី១ ប្រសិនបើជា ល.រ គោល)
    if (!(idx === 0 && isFirstColIndex)) {
      const deleteBtn = document.createElement('button');
      deleteBtn.className = 'col-delete-btn';
      deleteBtn.type = 'button';
      deleteBtn.title = 'លុបជួរឈរនេះ (×)';
      deleteBtn.innerHTML = '<i class="fas fa-times"></i>';
      actionsDiv.appendChild(deleteBtn);
    }

    th.appendChild(actionsDiv);
  });
}

/**
 * បន្ថែមជួរឈរថ្មីនៅបន្ទាប់ពីជួរឈរដែលបានជ្រើស (Insert Column to the right)
 */
function insertTableColumnAfter(targetTh) {
  const table = targetTh.closest('table') || document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const { grid, cellMeta, rows: allRows, maxCols } = getTableGridMatrix(table);
  const meta = cellMeta.get(targetTh);
  const targetCol = meta ? (meta.c + meta.colSpan - 1) : 0;

  const prevColName = getColumnHeaderTitle(targetTh);
  const nextColNum = maxCols + 1;

  const colName = prompt(
    `សូមបញ្ចូលចំណងជើងជួរឈរថ្មី (បន្ទាប់ពី «${prevColName || 'កូឡោនទី ' + khmerNumber(targetCol + 1)}»)៖`,
    `កូឡោនថ្មី ${khmerNumber(nextColNum)}`
  );
  if (colName === null) return; // បោះបង់
  const finalColName = colName.trim() || `កូឡោនថ្មី ${khmerNumber(nextColNum)}`;

  // 1. បង្កើតក្បាលជួរឈរ <th> ថ្មី ក្នុង thead
  const theadRow = table.querySelector('thead tr');
  const newTh = document.createElement('th');
  newTh.className = 'editable-header';
  newTh.setAttribute('contenteditable', 'true');
  newTh.textContent = finalColName;

  if (targetTh.nextSibling) {
    theadRow.insertBefore(newTh, targetTh.nextSibling);
  } else {
    const actionTh = theadRow.querySelector('.row-action-col');
    if (actionTh) theadRow.insertBefore(newTh, actionTh);
    else theadRow.appendChild(newTh);
  }

  // 2. បញ្ចូល <td> ថ្មីក្នុងគ្រប់ជួរដេកទាំងអស់នៃ <tbody>
  const tbody = table.querySelector('tbody');
  if (tbody) {
    const tbodyRows = Array.from(tbody.querySelectorAll('tr'));
    tbodyRows.forEach(row => {
      const rIdx = allRows.indexOf(row);
      const newTd = document.createElement('td');
      newTd.className = 'editable-cell';
      newTd.setAttribute('contenteditable', 'true');
      newTd.textContent = '';

      let anchorCell = null;
      if (rIdx >= 0) {
        for (let c = targetCol + 1; c < maxCols; c++) {
          const item = grid[rIdx] && grid[rIdx][c];
          if (item && item.isOrigin && item.cell && item.cell.parentNode === row) {
            anchorCell = item.cell;
            break;
          }
        }
      }
      if (!anchorCell) {
        anchorCell = row.querySelector('.row-action-col');
      }

      if (anchorCell) {
        row.insertBefore(newTd, anchorCell);
      } else {
        row.appendChild(newTd);
      }
    });
  }

  initTableColumnResizers();
  ensureColumnActionButtonsOnHeaders(table);
  ensureRowActionButtons(table);
  reindexTableRows(table);
  markDocumentModified();

  const allDataHeaders = table.querySelectorAll('thead th:not(.row-action-col)');
  if (allDataHeaders.length >= 10) {
    table.classList.add('table-ultra-wide');
    table.classList.remove('table-wide');
    setPrintOrientation('landscape');
  } else if (allDataHeaders.length >= 7) {
    table.classList.add('table-wide');
    table.classList.remove('table-ultra-wide');
    setPrintOrientation('landscape');
  }

  showToast(`បានបន្ថែមជួរឈរ «${finalColName}» ជោគជ័យ!`);
  newTh.focus();
}

/**
 * លុបជួរឈរទាំងមូលចេញពីតារាងតាម Grid Column Index (colIdx)
 * ធានាលុបក្បាលតារាង (th) និងក្រឡាទាំងអស់ (td) ក្នុងគ្រប់ជួរដេក ព្រមទាំងសម្រួល colspan/rowspan ដោយគ្មានចន្លោះប្រហោង
 */
function deleteColumnByIndex(table, colIdx, confirmPrompt = true) {
  if (!table) table = document.querySelector('#printableDocument .doc-table');
  if (!table) return false;

  const { grid, maxCols } = getTableGridMatrix(table);
  if (maxCols <= 1) {
    showToast('មិនអាចលុបបានទេ ត្រូវទុកយ៉ាងហោចណាស់មួយជួរឈរទិន្នន័យ។');
    return false;
  }

  if (colIdx < 0 || colIdx >= maxCols) {
    showToast('មិនអាចរកឃើញជួរឈរនេះទេ!');
    return false;
  }

  // ប្រសិនបើជាជួរឈរទី១ (ល.រ) មិនអនុញ្ញាតឱ្យលុបឡើយ
  if (colIdx === 0 && isRowIndexColumn(table)) {
    showToast('មិនអាចលុបជួរឈរលេខរៀង (ល.រ) គោលបានទេ!');
    return false;
  }

  // ស្វែងរកឈ្មោះជួរឈរដើម្បីបង្ហាញក្នុង Confirm dialog
  let colTitle = `កូឡោនទី ${khmerNumber(colIdx + 1)}`;
  for (let r = 0; r < grid.length; r++) {
    const item = grid[r] && grid[r][colIdx];
    if (item && item.cell && (item.cell.tagName === 'TH' || r === 0)) {
      const title = getColumnHeaderTitle(item.cell);
      if (title) {
        colTitle = title;
        break;
      }
    }
  }

  if (confirmPrompt) {
    const confirmDel = confirm(`តើលោកអ្នកពិតជាចង់លុបជួរឈរ «${colTitle}» នេះចេញពីតារាងមែនទេ?\nរាល់ទិន្នន័យក្នុងជួរឈរនេះនឹងត្រូវលុបចេញពីគ្រប់ជួរដេកទាំងអស់។`);
    if (!confirmDel) return false;
  }

  // ប្រមូលក្រឡាដែលត្រូវកែសម្រួល ឬលុបចេញ (Set ដើម្បីមិនឱ្យលុបស្ទួនពេលមាន rowspan)
  const processedCells = new Set();

  for (let r = 0; r < grid.length; r++) {
    const item = grid[r] && grid[r][colIdx];
    if (!item || !item.cell || processedCells.has(item.cell)) continue;
    processedCells.add(item.cell);

    const cell = item.cell;
    const colSpan = item.colSpan || 1;

    if (colSpan > 1) {
      // ក្រឡាបញ្ចូលគ្នា (Merged cell): កាត់បន្ថយ colspan ចុះ ១
      const newColSpan = colSpan - 1;
      if (newColSpan <= 1) {
        cell.removeAttribute('colspan');
      } else {
        cell.setAttribute('colspan', newColSpan);
      }
    } else {
      // ក្រឡាធម្មតា: លុបក្រឡានេះចេញពី DOM ទាំងស្រុង
      cell.remove();
    }
  }

  // សម្អាត Selection
  clearCellSelection();
  window._lastActiveCell = null;

  // រៀបចំលេខរៀងជួរដេកឡើងវិញ
  reindexTableRows(table);

  // បំពាក់ Resizers និង Action Buttons សារជាថ្មី
  initTableColumnResizers();
  ensureColumnActionButtonsOnHeaders(table);
  ensureRowActionButtons(table);

  markDocumentModified();

  // ពិនិត្យចំនួនជួរឈរដែលនៅសល់ដើម្បីកំណត់ table-wide និង orientation
  const remainingHeaders = table.querySelectorAll('thead th:not(.row-action-col)');
  if (remainingHeaders.length < 7) {
    table.classList.remove('table-wide', 'table-ultra-wide');
    setPrintOrientation('portrait');
  } else if (remainingHeaders.length < 10) {
    table.classList.remove('table-ultra-wide');
    table.classList.add('table-wide');
    setPrintOrientation('landscape');
  }

  showToast(`បានលុបជួរឈរ «${colTitle}» ចេញពីគ្រប់ជួរដេកទាំងអស់ជោគជ័យ!`);
  return true;
}

/**
 * លុបជួរឈរដែលបានជ្រើសចេញពីតារាងតាមក្បាលជួរឈរ targetTh
 */
function deleteTableColumnAt(targetTh) {
  const table = targetTh.closest('table') || document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const { cellMeta } = getTableGridMatrix(table);
  const meta = cellMeta.get(targetTh);
  if (!meta) {
    const theadRow = table.querySelector('thead tr');
    const dataHeaders = Array.from(theadRow.querySelectorAll('th:not(.row-action-col)'));
    const idx = dataHeaders.indexOf(targetTh);
    if (idx !== -1) {
      deleteColumnByIndex(table, idx, true);
    }
    return;
  }

  deleteColumnByIndex(table, meta.c, true);
}

/**
 * បើកផ្ទាំងជ្រើសរើសជួរឈរដើម្បីលុប
 */
function openDeleteColumnModal() {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) {
    showToast('សូមបើកផ្ទាំងគំរូឯកសារជាមុនសិន!');
    return;
  }

  const { grid, maxCols } = getTableGridMatrix(table);
  if (maxCols <= 1) {
    showToast('មិនអាចលុបបានទេ ត្រូវទុកយ៉ាងហោចណាស់មួយជួរឈរ!');
    return;
  }

  const container = document.getElementById('deleteColumnListContainer');
  if (!container) return;

  const isFirstIndex = isRowIndexColumn(table);
  let listHtml = '';

  for (let c = 0; c < maxCols; c++) {
    let colName = '';
    for (let r = 0; r < grid.length; r++) {
      const item = grid[r] && grid[r][c];
      if (item && item.cell && (item.cell.tagName === 'TH' || r === 0)) {
        colName = getColumnHeaderTitle(item.cell);
        if (colName) break;
      }
    }
    if (!colName) colName = `កូឡោនទី ${khmerNumber(c + 1)}`;

    const isLocked = (c === 0 && isFirstIndex);

    if (isLocked) {
      listHtml += `
        <div class="col-delete-item disabled">
          <div class="col-delete-info">
            <span class="col-num-badge">កូឡោនទី ${khmerNumber(c + 1)}</span>
            <span class="col-delete-name">${colName} (ជួរឈរគោល)</span>
          </div>
          <span style="font-size: 0.78rem; color: #94a3b8; font-weight: 500;"><i class="fas fa-lock"></i> រក្សាទុក</span>
        </div>
      `;
    } else {
      listHtml += `
        <div class="col-delete-item">
          <div class="col-delete-info" title="${colName}">
            <span class="col-num-badge">កូឡោនទី ${khmerNumber(c + 1)}</span>
            <span class="col-delete-name">${colName}</span>
          </div>
          <button type="button" class="btn-del-col-action" onclick="executeDeleteColumn(${c})">
            <i class="fas fa-trash-alt"></i> លុបកូឡោននេះ
          </button>
        </div>
      `;
    }
  }

  container.innerHTML = listHtml;

  const modal = document.getElementById('deleteColumnModal');
  if (modal) {
    modal.classList.add('active');
  }
}

/**
 * បិទផ្ទាំងលុបជួរឈរ
 */
function closeDeleteColumnModal() {
  const modal = document.getElementById('deleteColumnModal');
  if (modal) {
    modal.classList.remove('active');
  }
}

/**
 * អនុវត្តការលុបជួរឈរជាក់ស្តែងតាម Index
 */
function executeDeleteColumn(colIndex) {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const success = deleteColumnByIndex(table, colIndex, true);
  if (!success) return;

  const remainingHeaders = table.querySelectorAll('thead th:not(.row-action-col)');
  if (remainingHeaders.length > 1) {
    openDeleteColumnModal();
  } else {
    closeDeleteColumnModal();
  }
}

/**
 * លុបជួរឈរចុងក្រោយបង្អស់នៃតារាង
 */
function deleteLastTableColumn() {
  const table = document.querySelector('#printableDocument .doc-table');
  if (!table) return;

  const { maxCols } = getTableGridMatrix(table);
  if (maxCols <= 1) {
    showToast('មិនអាចលុបបានទៀតទេ ត្រូវទុកយ៉ាងហោចណាស់មួយជួរឈរ!');
    return;
  }

  deleteColumnByIndex(table, maxCols - 1, true);
}

/**
 * ធានាថាប៊ូតុងបន្ថែម/លុបជួរឈរ និងឧបករណ៍តារាងទាំងអស់ មានវត្តមានជានិច្ចក្នុងទម្រង់ ២ ជួរទំហំស្មើគ្នា
 */
function ensureColumnActionButtons() {
  const tableActionsBar = document.querySelector('#printableDocument .table-actions-bar');
  if (!tableActionsBar) return;

  const grid = tableActionsBar.querySelector('.table-actions-grid');
  if (!grid || !grid.querySelector('.btn-merge-cells') || !grid.querySelector('.btn-unmerge-cells') || !grid.querySelector('.btn-widen-col')) {
    tableActionsBar.innerHTML = getTableActionsBarHtml();
  }
}

/**
 * បច្ចុប្បន្នកម្ម Badge លើ Card ក្នុងបញ្ជីសូចនាករ
 */
function updateCardBadge(subCode, isSaved) {
  const cardId = 'card-' + subCode.replace(/\./g, '-');
  const card = document.getElementById(cardId);
  if (!card) return;

  const headerDiv = card.querySelector('.card-header-row > div');
  if (!headerDiv) return;

  const existingBadge = headerDiv.querySelector('.card-edited-badge');
  if (isSaved) {
    if (!existingBadge) {
      const badge = document.createElement('span');
      badge.className = 'card-edited-badge';
      badge.title = 'មានទិន្នន័យបានកែប្រែ និងរក្សាទុក';
      badge.innerHTML = '<i class="fas fa-check-circle"></i> បានកែប្រែ';
      headerDiv.insertBefore(badge, headerDiv.firstChild);
    }
  } else {
    if (existingBadge) {
      existingBadge.remove();
    }
  }
}

/**
 * បំប្លែងកាលបរិច្ឆេទជាទម្រង់ខ្មែរ
 */
function formatKhmerDateTime(date = new Date()) {
  const months = ['មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា', 'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'];
  const day = khmerNumber(date.getDate());
  const month = months[date.getMonth()];
  const year = khmerNumber(date.getFullYear());
  const hours = khmerNumber(String(date.getHours()).padStart(2, '0'));
  const minutes = khmerNumber(String(date.getMinutes()).padStart(2, '0'));
  return `ថ្ងៃទី ${day} ខែ ${month} ឆ្នាំ ${year} ម៉ោង ${hours}:${minutes}`;
}

function closeDocumentModal() {
  if (hasUnsavedChanges) {
    const confirmClose = confirm("លោកអ្នកមានការកែប្រែដែលមិនទាន់បានរក្សាទុក! តើលោកអ្នកពិតជាចង់ចាកចេញដោយមិនរក្សាទុកមែនទេ?");
    if (!confirmClose) return;
  }
  const modal = document.getElementById('documentModal');
  if (modal) modal.classList.remove('active');
  document.body.style.overflow = '';
  hasUnsavedChanges = false;
}

function initModalActions() {
  // Close when clicking backdrop
  const modal = document.getElementById('documentModal');
  if (modal) {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeDocumentModal();
    });
  }

  const delColModal = document.getElementById('deleteColumnModal');
  if (delColModal) {
    delColModal.addEventListener('click', (e) => {
      if (e.target === delColModal) closeDeleteColumnModal();
    });
  }

  // ESC key to close
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      const delColModal = document.getElementById('deleteColumnModal');
      if (delColModal && delColModal.classList.contains('active')) {
        closeDeleteColumnModal();
        return;
      }
      closeDocumentModal();
    }
    // ផ្លូវកាត់ Ctrl+S ឬ Cmd+S ដើម្បីរក្សាទុកឯកសារភ្លាមៗ
    if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
      const modal = document.getElementById('documentModal');
      if (modal && modal.classList.contains('active')) {
        e.preventDefault();
        saveCurrentDocument();
      }
    }
  });
}

/**
 * មុខងារបោះពុម្ពឯកសារ (Print Document)
 */
function printCurrentDocument() {
  applyPrintOrientationStyles(currentPrintOrientation);
  window.print();
}

function quickPrintDocument(subCode) {
  openDocumentModal(subCode);
  setTimeout(() => {
    printCurrentDocument();
  }, 400);
}

/**
 * មុខងារទាញយកឯកសារ Word ផ្ទាល់តាមលេខកូដសូចនាកររង
 * បើមានទិន្នន័យកែប្រែ រក្សាទុក នឹងទាញយកទិន្នន័យកែប្រែនោះ
 */
function downloadWordFile(subCode) {
  const sub = getSubIndicatorByCode(subCode);
  if (!sub) return;

  // ពិនិត្យមើលថាតើអ្នកប្រើប្រាស់ធ្លាប់បានកែប្រែ និងរក្សាទុកឯកសារនេះឬនៅ
  const savedHtml = localStorage.getItem('saved_doc_' + subCode);
  if (savedHtml) {
    exportSavedHtmlToWord(sub, savedHtml);
    return;
  }

  const std = getStandardById(sub.standardId);
  const stdFolder = `${std.title}`;
  const fileName = `ឯកសារ_${sub.code}_${sub.templateTitle}.doc`;
  
  // Link to pre-generated file in folder
  const filePath = `${stdFolder}/${fileName}`;
  const a = document.createElement('a');
  a.href = filePath;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  showToast(`បានបើក/ទាញយក៖ ${fileName}`);
}

/**
 * មុខងារទាញយកជាឯកសារ Microsoft Word (.doc) ពីក្នុងផ្ទាំង Modal (រក្សាទុកទិន្នន័យកែប្រែទាំងអស់)
 */
function exportToWord() {
  const printableArea = document.getElementById('printableDocument');
  if (!printableArea || !currentActiveDoc) return;

  // បង្កើត Clone ដើម្បីលុបចេញនូវធាតុមិនចាំបាច់ពេលទាញយក (ប៊ូតុងលុប របារបញ្ជា)
  const clone = printableArea.cloneNode(true);
  clone.querySelectorAll('.no-print').forEach(el => el.remove());
  clone.querySelectorAll('[contenteditable]').forEach(el => el.removeAttribute('contenteditable'));

  // ពិនិត្យមើលថាតើតារាងមានជួរឈរច្រើន ឬស្ថិតក្នុងទម្រង់ Landscape ដែរឬទេ
  const table = printableArea.querySelector('.doc-table');
  const colCount = table ? table.querySelectorAll('thead th:not(.row-action-col)').length : 5;
  const isLandscape = (currentPrintOrientation === 'landscape' || colCount >= 7);

  const header = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' 
          xmlns:w='urn:schemas-microsoft-com:office:word' 
          xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <meta charset="utf-8">
      <title>${currentActiveDoc.templateTitle}</title>
      <!--[if gte mso 9]>
      <xml>
        <w:WordDocument>
          <w:View>Print</w:View>
          <w:Zoom>100</w:Zoom>
          <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
      </xml>
      <![endif]-->
      <style>
        @page Section1 {
          size: ${isLandscape ? '841.9pt 595.3pt' : '595.3pt 841.9pt'};
          mso-page-orientation: ${isLandscape ? 'landscape' : 'portrait'};
          margin: 0.5in 0.6in 0.5in 0.6in;
          mso-header-margin: 0.5in;
          mso-footer-margin: 0.5in;
        }
        div.Section1 { page: Section1; }
        body { font-family: 'Khmer OS Siemreap', 'Kantumruy Pro', Arial; font-size: ${isLandscape ? '9.5pt' : '11pt'}; }
        .indicator-badge-top-right { float: right; border: 2px solid black; padding: 5px 10px; text-align: center; }
        table { border-collapse: collapse; width: 100%; margin: 15px 0; table-layout: auto; }
        th, td { border: 1px solid black; padding: ${isLandscape ? '4px 6px' : '6px 8px'}; font-size: ${isLandscape ? '9pt' : '10pt'}; word-break: break-word; }
        th { background-color: #f2f2f2; text-align: center; font-weight: bold; }
        .text-center { text-align: center; }
        .doc-header { margin-bottom: 20px; }
        .doc-main-title { text-align: center; font-family: 'Khmer OS Muol Light'; font-size: 14pt; margin: 15px 0; }
        .doc-signatures { margin-top: 40px; width: 100%; }
        .doc-report-page + .doc-report-page { page-break-before: always; margin-top: 30pt; }
        .document-page-2 { page-break-before: always; margin-top: 30pt; }
        .annex-photo-card { border: 1px solid black; margin-bottom: 20px; page-break-inside: avoid; }
        .annex-photo-img-wrap { text-align: center; margin-bottom: 10px; }
        .annex-photo-img-wrap img { max-width: 100%; max-height: 280pt; }
        .annex-pdf-wrap { text-align: center; padding: 20pt; border-bottom: 1px solid #ccc; background-color: #fef2f2; }
        .annex-photo-caption-wrap { padding: 8px; }
      </style>
    </head>
    <body>
      <div class="Section1">
  `;
  const footer = `</div></body></html>`;
  const sourceHTML = header + clone.innerHTML + footer;

  const source = 'data:application/vnd.ms-word;charset=utf-8,\uFEFF' + encodeURIComponent(sourceHTML);
  const fileDownload = document.createElement("a");
  document.body.appendChild(fileDownload);
  fileDownload.href = source;
  fileDownload.download = `ទម្រង់ឯកសារ_${currentActiveDoc.code}_${currentActiveDoc.templateTitle}.doc`;
  fileDownload.click();
  document.body.removeChild(fileDownload);

  showToast(`បានទាញយកទម្រង់ ${currentActiveDoc.code} ជាឯកសារ Word ជោគជ័យ!`);
}

/**
 * មុខងារទាញយក Word ពីទិន្នន័យដែលបានកែប្រែ និងរក្សាទុកក្នុង LocalStorage
 */
function exportSavedHtmlToWord(sub, savedHtml) {
  let cleanSavedHtml = (savedHtml || '')
    .replace(/<div class="doc-meta-item">\s*<strong>\s*សាលាបឋមសិក្សា[៖:]\s*<\/strong>[\s\S]*?<\/div>\s*/g, '')
    .replace(/ប្រធានគណៈកម្មការ\s*គ\.ក\.ស\./g, 'ប្រធាន គ.គ.ស')
    .replace(/នាយកសាលាបឋមសិក្សាតាពីង/g, 'នាយកសាលា');
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = cleanSavedHtml;

  // បន្ថែមឧបសម្ព័ន្ធភស្តុតាងប្រសិនបើមាន
  const evList = getEvidenceList(sub.code).filter(item => item && (item.imageData || hasLinkUrl(item) || (item.caption && item.caption.trim())));
  if (evList.length > 0) {
    tempDiv.innerHTML += `<div class="document-page-2" style="page-break-before: always; margin-top: 30pt;">${renderEvidenceAnnexHtml(sub)}</div>`;
  }

  tempDiv.querySelectorAll('.no-print').forEach(el => el.remove());
  tempDiv.querySelectorAll('[contenteditable]').forEach(el => el.removeAttribute('contenteditable'));

  // ពិនិត្យមើលថាតើតារាងមានជួរឈរច្រើន ឬស្ថិតក្នុងទម្រង់ Landscape ដែរឬទេ
  const table = tempDiv.querySelector('.doc-table');
  const colCount = table ? table.querySelectorAll('thead th:not(.row-action-col)').length : (sub.headers ? sub.headers.length : 5);
  const isLandscape = (currentPrintOrientation === 'landscape' || colCount >= 7);

  const header = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' 
          xmlns:w='urn:schemas-microsoft-com:office:word' 
          xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <meta charset="utf-8">
      <title>${sub.templateTitle}</title>
      <!--[if gte mso 9]>
      <xml>
        <w:WordDocument>
          <w:View>Print</w:View>
          <w:Zoom>100</w:Zoom>
          <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
      </xml>
      <![endif]-->
      <style>
        @page Section1 {
          size: ${isLandscape ? '841.9pt 595.3pt' : '595.3pt 841.9pt'};
          mso-page-orientation: ${isLandscape ? 'landscape' : 'portrait'};
          margin: 0.5in 0.6in 0.5in 0.6in;
          mso-header-margin: 0.5in;
          mso-footer-margin: 0.5in;
        }
        div.Section1 { page: Section1; }
        body { font-family: 'Khmer OS Siemreap', 'Kantumruy Pro', Arial; font-size: ${isLandscape ? '9.5pt' : '11pt'}; }
        .indicator-badge-top-right { float: right; border: 2px solid black; padding: 5px 10px; text-align: center; }
        table { border-collapse: collapse; width: 100%; margin: 15px 0; table-layout: auto; }
        th, td { border: 1px solid black; padding: ${isLandscape ? '4px 6px' : '6px 8px'}; font-size: ${isLandscape ? '9pt' : '10pt'}; word-break: break-word; }
        th { background-color: #f2f2f2; text-align: center; font-weight: bold; }
        .text-center { text-align: center; }
        .doc-header { margin-bottom: 20px; }
        .doc-main-title { text-align: center; font-family: 'Khmer OS Muol Light'; font-size: 14pt; margin: 15px 0; }
        .doc-signatures { margin-top: 40px; width: 100%; }
        .doc-report-page + .doc-report-page { page-break-before: always; margin-top: 30pt; }
        .document-page-2 { page-break-before: always; margin-top: 30pt; }
        .annex-photo-card { border: 1px solid black; margin-bottom: 20px; page-break-inside: avoid; }
        .annex-photo-img-wrap img { max-width: 100%; max-height: 280pt; }
        .annex-pdf-wrap { text-align: center; padding: 20pt; border-bottom: 1px solid #ccc; background-color: #fef2f2; }
        .annex-photo-caption-wrap { padding: 8px; }
        .annex-link-wrap { text-align: center; padding: 15pt; border-bottom: 1px solid #ccc; background-color: #eff6ff; }
        .annex-link-platform-badge { font-weight: bold; color: #2563eb; margin-bottom: 4pt; font-size: 10pt; }
        .annex-link-url { font-family: 'Courier New', monospace; font-size: 9.5pt; color: #1d4ed8; word-break: break-all; }
        .annex-link-print-url { font-size: 9pt; color: #333; margin-top: 6pt; }
      </style>
    </head>
    <body>
      <div class="Section1">
  `;
  const footer = `</div></body></html>`;
  const sourceHTML = header + tempDiv.innerHTML + footer;

  const source = 'data:application/vnd.ms-word;charset=utf-8,\uFEFF' + encodeURIComponent(sourceHTML);
  const fileDownload = document.createElement("a");
  document.body.appendChild(fileDownload);
  fileDownload.href = source;
  fileDownload.download = `ទម្រង់ឯកសារ_${sub.code}_${sub.templateTitle}_(កែសម្រួល).doc`;
  fileDownload.click();
  document.body.removeChild(fileDownload);

  showToast(`បានទាញយកទិន្នន័យកែប្រែនៃ ${sub.code} ជាឯកសារ Word ជោគជ័យ!`);
}

/**
 * ចម្លងអត្ថបទឯកសារ (Copy Content)
 */
function copyDocumentText() {
  const printableArea = document.getElementById('printableDocument');
  if (!printableArea || !currentActiveDoc) return;

  navigator.clipboard.writeText(printableArea.innerText).then(() => {
    showToast(`បានចម្លងខ្លឹមសារឯកសារ ${currentActiveDoc.code} រួចរាល់!`);
  }).catch(() => {
    showToast("សូមអភ័យទោស មិនអាចចម្លងបានឡើយ។");
  });
}

/**
 * បង្ហាញសារជូនដំណឹង (Toast Message)
 */
function showToast(msg) {
  let toast = document.getElementById('appToast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'appToast';
    toast.className = 'toast-msg';
    document.body.appendChild(toast);
  }
  toast.innerHTML = `<i class="fas fa-check-circle" style="color: #4ade80;"></i> <span>${msg}</span>`;
  toast.classList.add('show');
  setTimeout(() => {
    toast.classList.remove('show');
  }, 3200);
}

/**
 * បំប្លែងលេខអារ៉ាប់ទៅជាលេខខ្មែរ
 */
function khmerNumber(num) {
  const khmerDigits = ['០', '១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩'];
  return num.toString().split('').map(d => khmerDigits[d] !== undefined ? khmerDigits[d] : d).join('');
}

/* ==========================================================================
   EVIDENCE MANAGEMENT & ANNEX PAGE LOGIC (ប្រព័ន្ធគ្រប់គ្រងភស្តុតាង និងទំព័រឧបសម្ព័ន្ធ)
   IndexedDB High-Capacity Storage Engine (គាំទ្ររូបភាព/PDF ៥+ និង ៥ លីង)
   ========================================================================== */

const EVIDENCE_DB_NAME = 'SchoolStandardEvidenceDB';
const EVIDENCE_DB_VERSION = 1;
const EVIDENCE_STORE_NAME = 'evidence_store';
let evidenceDB = null;
let evidenceMemoryCache = {};
let isEvidenceDbInitialized = false;

// Populate initial cache immediately from localStorage (fast fallback)
(function initCacheFromLocalStorage() {
  try {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith('evidence_')) {
        const subCode = key.substring('evidence_'.length);
        const data = localStorage.getItem(key);
        if (data) {
          try {
            const list = JSON.parse(data);
            if (Array.isArray(list)) {
              list.forEach(item => {
                if (item && item.date) item.date = formatKhmerMonthNameToWord(item.date);
              });
              evidenceMemoryCache[subCode] = list;
            }
          } catch (e) {}
        }
      }
    }
  } catch (e) {}
})();

/**
 * បង្កើត និងបើក IndexedDB Database សម្រាប់រក្សាទុកភស្តុតាងទំហំធំ (គាំទ្ររាប់សិប MB/GB គ្មានដែនកំណត់ 5MB)
 */
function openEvidenceDatabase() {
  return new Promise((resolve) => {
    if (evidenceDB) {
      resolve(evidenceDB);
      return;
    }
    if (!window.indexedDB) {
      console.warn('Browser does not support IndexedDB, falling back to localStorage');
      resolve(null);
      return;
    }
    try {
      const request = indexedDB.open(EVIDENCE_DB_NAME, EVIDENCE_DB_VERSION);
      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(EVIDENCE_STORE_NAME)) {
          db.createObjectStore(EVIDENCE_STORE_NAME, { keyPath: 'subCode' });
        }
      };
      request.onsuccess = (e) => {
        evidenceDB = e.target.result;
        resolve(evidenceDB);
      };
      request.onerror = (e) => {
        console.error('Failed to open IndexedDB:', e);
        resolve(null);
      };
    } catch (err) {
      console.error('Error opening IndexedDB:', err);
      resolve(null);
    }
  });
}

/**
 * ដំណើរការ Auto-Migration ពី LocalStorage និងផ្ទុកទិន្នន័យភស្តុតាងទាំងអស់ចូលក្នុង Memory Cache
 */
async function initEvidenceDB() {
  const db = await openEvidenceDatabase();
  if (!db) {
    isEvidenceDbInitialized = true;
    return;
  }

  try {
    const tx = db.transaction(EVIDENCE_STORE_NAME, 'readonly');
    const store = tx.objectStore(EVIDENCE_STORE_NAME);
    const getAllReq = store.getAll();

    getAllReq.onsuccess = () => {
      const records = getAllReq.result || [];
      records.forEach(rec => {
        if (rec && rec.subCode && Array.isArray(rec.items)) {
          rec.items.forEach(item => {
            if (item && item.date) item.date = formatKhmerMonthNameToWord(item.date);
          });
          evidenceMemoryCache[rec.subCode] = rec.items;
        }
      });

      // Auto-Migrate: ប្រសិនបើមានទិន្នន័យក្នុង LocalStorage តែពុំទាន់មានក្នុង IndexedDB ត្រូវ Save ចូល IndexedDB
      Object.keys(evidenceMemoryCache).forEach(subCode => {
        const items = evidenceMemoryCache[subCode];
        if (items && items.length > 0) {
          saveEvidenceToDB(subCode, items);
        }
      });

      // សម្អាត Base64 ធំៗចេញពី LocalStorage ដើម្បីកុំឱ្យស្ទះទំហំផ្ទុក 5MB
      cleanupLocalStorageEvidenceImages();

      isEvidenceDbInitialized = true;
      updateAllEvidenceBadges();
    };

    getAllReq.onerror = () => {
      isEvidenceDbInitialized = true;
    };
  } catch (err) {
    console.error('Error loading all evidence from IndexedDB:', err);
    isEvidenceDbInitialized = true;
  }
}

/**
 * រក្សាទុកទិន្នន័យភស្តុតាងចូល IndexedDB
 */
function saveEvidenceToDB(subCode, items) {
  return new Promise((resolve) => {
    if (!evidenceDB) {
      resolve(false);
      return;
    }
    try {
      const tx = evidenceDB.transaction(EVIDENCE_STORE_NAME, 'readwrite');
      const store = tx.objectStore(EVIDENCE_STORE_NAME);
      const req = store.put({ subCode: subCode, items: items, updatedAt: new Date().toISOString() });
      req.onsuccess = () => resolve(true);
      req.onerror = () => resolve(false);
    } catch (e) {
      console.error('Error saving evidence to IndexedDB:', e);
      resolve(false);
    }
  });
}

/**
 * ធ្វើបច្ចុប្បន្នកម្ម Badge ភស្តុតាងទាំងអស់លើកាត
 */
function updateAllEvidenceBadges() {
  if (typeof STANDARDS_DATA === 'undefined') return;
  STANDARDS_DATA.forEach(std => {
    if (!std.keyIndicators) return;
    std.keyIndicators.forEach(keyInd => {
      if (!keyInd.subIndicators) return;
      keyInd.subIndicators.forEach(sub => {
        updateCardEvidenceBadge(sub.code);
      });
    });
  });
}

/**
 * ទទួលយកបញ្ជីភស្តុតាងពី Memory Cache (IndexedDB) ឬ LocalStorage
 */
function getEvidenceList(subCode) {
  if (evidenceMemoryCache[subCode] && Array.isArray(evidenceMemoryCache[subCode])) {
    return evidenceMemoryCache[subCode];
  }
  try {
    const data = localStorage.getItem('evidence_' + subCode);
    if (!data) return [];
    const list = JSON.parse(data);
    if (Array.isArray(list)) {
      list.forEach(item => {
        if (item && item.date) {
          item.date = formatKhmerMonthNameToWord(item.date);
        }
      });
      evidenceMemoryCache[subCode] = list;
      return list;
    }
    return [];
  } catch (e) {
    return [];
  }
}

/**
 * រាប់ចំនួនភស្តុតាងដែលមានរូបភាព ឬឯកសារ ឬតំណលីង ឬចំណងជើង
 */
function getEvidenceCount(subCode) {
  const list = getEvidenceList(subCode);
  return list.filter(item => item && (item.imageData || hasLinkUrl(item) || (item.caption && item.caption.trim()))).length;
}

/**
 * Escape HTML characters to prevent XSS
 */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * ពិនិត្យថាតើធាតុភស្តុតាងមានឯកសាររូបភាព ឬ PDF ផ្ទុកមកជាមួយឬទេ
 */
function hasMediaFile(item) {
  return Boolean(item && item.imageData);
}

/**
 * ទទួលយកបញ្ជីតំណលីងពីធាតុភស្តុតាង (គាំទ្ររហូតដល់ ៥ លីង)
 */
function getItemLinks(item) {
  if (!item) return [];
  if (Array.isArray(item.links) && item.links.length > 0) {
    const valid = item.links.filter(l => l && (typeof l === 'string' ? l.trim() : (l.url && l.url.trim())));
    if (valid.length > 0) return valid;
  }
  if (item.linkUrl && item.linkUrl.trim()) {
    return [{ url: item.linkUrl.trim() }];
  }
  return [];
}

/**
 * ពិនិត្យថាតើធាតុភស្តុតាងមានតំណលីង (URL / Web link) ឬទេ
 */
function hasLinkUrl(item) {
  if (!item) return false;
  if (item.linkUrl && item.linkUrl.trim()) return true;
  if (Array.isArray(item.links) && item.links.some(l => l && (typeof l === 'string' ? l.trim() : (l.url && l.url.trim())))) return true;
  return false;
}

/**
 * ពិនិត្យថាតើធាតុភស្តុតាងជាតំណលីងសុទ្ធ (Link Only - គ្មានហ្វាល់រូបភាព/PDF) ឬទេ
 */
function isLinkItem(item) {
  return hasLinkUrl(item) && !hasMediaFile(item);
}

/**
 * ពិនិត្យថាតើធាតុភស្តុតាងជាឯកសារ PDF ឬទេ
 */
function isPdfItem(item) {
  if (!item || !item.imageData) return false;
  if (item.isPdf === true || item.fileType === 'application/pdf') return true;
  if (typeof item.imageData === 'string' && item.imageData.startsWith('data:application/pdf')) return true;
  if (typeof item.fileName === 'string' && item.fileName.toLowerCase().endsWith('.pdf')) return true;
  return false;
}

/**
 * ពិនិត្យថាតើធាតុភស្តុតាងជារូបភាព ឬទេ
 */
function isImageItem(item) {
  return hasMediaFile(item) && !isPdfItem(item);
}

/**
 * វិភាគ និងផ្ទៀងផ្ទាត់ប្រភេទតំណលីង (Google Drive, Docs, YouTube, etc.)
 */
function getLinkPlatformInfo(url) {
  if (!url || typeof url !== 'string') {
    return { name: 'តំណភ្ជាប់', icon: 'fas fa-link', color: '#2563eb', bg: '#eff6ff' };
  }
  const low = url.toLowerCase();
  if (low.includes('drive.google.com')) {
    return { name: 'Google Drive', icon: 'fab fa-google-drive', color: '#0F9D58', bg: '#e8f5e9' };
  }
  if (low.includes('docs.google.com/document') || low.includes('docs.google.com')) {
    return { name: 'Google Docs', icon: 'fas fa-file-alt', color: '#4285F4', bg: '#e8f0fe' };
  }
  if (low.includes('docs.google.com/spreadsheets')) {
    return { name: 'Google Sheets', icon: 'fas fa-file-excel', color: '#0F9D58', bg: '#e8f5e9' };
  }
  if (low.includes('youtube.com') || low.includes('youtu.be')) {
    return { name: 'YouTube Video', icon: 'fab fa-youtube', color: '#FF0000', bg: '#fef2f2' };
  }
  if (low.includes('facebook.com') || low.includes('fb.watch')) {
    return { name: 'Facebook', icon: 'fab fa-facebook', color: '#1877F2', bg: '#eff6ff' };
  }
  if (low.includes('dropbox.com')) {
    return { name: 'Dropbox', icon: 'fab fa-dropbox', color: '#0061FF', bg: '#eff6ff' };
  }
  if (low.includes('onedrive') || low.includes('sharepoint.com')) {
    return { name: 'OneDrive / SharePoint', icon: 'fas fa-cloud', color: '#0078D4', bg: '#eff6ff' };
  }
  return { name: 'តំណភ្ជាប់គេហទំព័រ', icon: 'fas fa-globe', color: '#0284c7', bg: '#f0f9ff' };
}

/**
 * បំប្លែង Base64 DataURL ទៅជា Blob URL ដើម្បីគាំទ្រការបើកមើល PDF/រូបភាពដោយរលូនក្នុង browser
 */
function dataUrlToBlobUrl(dataUrl, defaultMime = 'application/octet-stream') {
  if (!dataUrl || typeof dataUrl !== 'string') return '';
  try {
    const parts = dataUrl.split(',');
    if (parts.length < 2) return dataUrl;
    const match = parts[0].match(/:(.*?);/);
    const mime = match ? match[1] : defaultMime;
    const bstr = atob(parts[1]);
    let n = bstr.length;
    const u8arr = new Uint8Array(n);
    while (n--) {
      u8arr[n] = bstr.charCodeAt(n);
    }
    const blob = new Blob([u8arr], { type: mime });
    return URL.createObjectURL(blob);
  } catch (e) {
    console.error('Error converting dataUrl to Blob:', e);
    return dataUrl;
  }
}

/**
 * ធ្វើបច្ចុប្បន្នកម្ម Badge និងប៊ូតុងភស្តុតាងលើ Card
 */
function updateCardEvidenceBadge(subCode) {
  const cardId = 'card-' + subCode.replace(/\./g, '-');
  const card = document.getElementById(cardId);
  if (!card) return;

  const evCount = getEvidenceCount(subCode);
  const headerDiv = card.querySelector('.card-header-row > div');
  if (headerDiv) {
    let badge = headerDiv.querySelector('.card-evidence-badge');
    if (evCount > 0) {
      if (!badge) {
        badge = document.createElement('span');
        badge.className = 'card-evidence-badge';
        headerDiv.insertBefore(badge, headerDiv.children[1] || null);
      }
      badge.title = `មានភស្តុតាងភ្ជាប់ចំនួន ${khmerNumber(evCount)}`;
      badge.innerHTML = `<i class="fas fa-images"></i> ភស្តុតាង (${khmerNumber(evCount)})`;
    } else if (badge) {
      badge.remove();
    }
  }

  const evBtn = card.querySelector('.btn-evidence');
  if (evBtn) {
    if (evCount > 0) {
      evBtn.classList.add('has-evidence');
      evBtn.innerHTML = `<i class="fas fa-images"></i> ភស្តុតាង <span class="evidence-badge-num">${khmerNumber(evCount)}</span>`;
    } else {
      evBtn.classList.remove('has-evidence');
      evBtn.innerHTML = `<i class="fas fa-images"></i> ភស្តុតាង`;
    }
  }
}

/**
 * បង្កើតកូដ HTML នៃទំព័រឧបសម្ព័ន្ធ (ទំព័រទី២) សម្រាប់បង្ហាញក្នុង Preview Modal
 */
function renderEvidenceAnnexHtml(sub) {
  const list = getEvidenceList(sub.code);
  const validItems = list.filter(item => item && (item.imageData || hasLinkUrl(item) || (item.caption && item.caption.trim())));

  if (validItems.length === 0) {
    return `
      <div class="no-evidence-annex-box">
        <i class="fas fa-images" style="font-size: 3.2rem; color: #0d9488; margin-bottom: 0.75rem;"></i>
        <h3 style="font-size: 1.15rem; color: #0f172a; margin-bottom: 0.4rem;">ពុំទាន់មានភស្តុតាងបញ្ជាក់បន្ថែមនៅឡើយទេ</h3>
        <p style="font-size: 0.9rem; color: #64748b; margin-bottom: 1.25rem;">លោកអ្នកអាចបញ្ចូលរូបភាពសកម្មភាព ឯកសារ PDF ឬតំណលីង ដើម្បីភ្ជាប់ជាមួយទម្រង់ឯកសារនេះ។</p>
        <button type="button" class="btn-add-evidence-inline" onclick="openEvidenceModal('${sub.code}')">
          <i class="fas fa-plus-circle"></i> បញ្ចូលភស្តុតាងបញ្ជាក់ (+)
        </button>
      </div>
    `;
  }

  const cardItems = [];
  validItems.forEach((item, idx) => {
    const hasFile = hasMediaFile(item);
    const isPdf = hasFile && isPdfItem(item);
    const isImg = hasFile && !isPdf;
    const hasLink = hasLinkUrl(item);
    const plat = hasLink ? getLinkPlatformInfo(item.linkUrl) : null;

    let mediaVisualHtml = '';
    let typeBadges = '';

    if (isImg) {
      typeBadges += `<span class="media-type-badge badge-image" style="margin-right: 4px;"><i class="fas fa-image"></i> រូបភាព</span>`;
      mediaVisualHtml = `
        <div class="annex-photo-img-wrap" onclick="openMediaViewerByDocCode('${sub.code}', ${idx})" title="ចុចដើម្បីបើកមើលរូបភាពពេញលេញ">
          <img src="${item.imageData}" alt="${escapeHtml(item.caption || 'រូបភាពភស្តុតាង')}">
          <div class="annex-view-overlay-btn no-print">
            <i class="fas fa-search-plus"></i> ពង្រីកមើល
          </div>
        </div>
      `;
    } else if (isPdf) {
      typeBadges += `<span class="media-type-badge badge-pdf" style="margin-right: 4px;"><i class="fas fa-file-pdf"></i> PDF</span>`;
      mediaVisualHtml = `
        <div class="annex-pdf-wrap" onclick="openMediaViewerByDocCode('${sub.code}', ${idx})" title="ចុចដើម្បីបើកមើលឯកសារ PDF នេះ">
          <i class="fas fa-file-pdf annex-pdf-icon-big"></i>
          <div class="annex-pdf-name">${escapeHtml(item.fileName || 'ឯកសារភស្តុតាង PDF')}</div>
          ${item.fileSize ? `<div class="annex-pdf-meta-tag"><span class="badge-pdf-tag">PDF</span> <span>${escapeHtml(item.fileSize)}</span></div>` : ''}
          <button type="button" class="btn-annex-open-pdf no-print">
            <i class="fas fa-eye"></i> បើកអានឯកសារ PDF នេះ
          </button>
        </div>
      `;
    } else if (hasLink) {
      const itemLinks = getItemLinks(item);
      const primaryUrl = typeof itemLinks[0] === 'string' ? itemLinks[0] : (itemLinks[0]?.url || item.linkUrl);
      const plat = getLinkPlatformInfo(primaryUrl);
      typeBadges += `<span class="media-type-badge badge-link" style="margin-right: 4px;"><i class="${plat.icon}"></i> ${itemLinks.length > 1 ? `តំណលីង (${khmerNumber(itemLinks.length)})` : escapeHtml(plat.name)}</span>`;
      mediaVisualHtml = `
        <div class="annex-link-wrap" style="background: ${plat.bg};" onclick="openMediaViewerByDocCode('${sub.code}', ${idx})" title="ចុចដើម្បីបើកមើលព័ត៌មានតំណលីង">
          <i class="${itemLinks.length > 1 ? 'fas fa-link' : plat.icon} annex-link-icon-big" style="color: ${plat.color};"></i>
          <div class="annex-link-platform-badge" style="background: ${plat.color}; color: white;">
            ${itemLinks.length > 1 ? `តំណភ្ជាប់ចំនួន ${khmerNumber(itemLinks.length)}` : escapeHtml(plat.name)}
          </div>
          <div class="annex-link-url" title="${escapeHtml(primaryUrl)}">
            ${escapeHtml(primaryUrl)}
          </div>
          <a href="${escapeHtml(primaryUrl)}" target="_blank" rel="noopener noreferrer" class="btn-annex-open-link no-print" onclick="event.stopPropagation()">
            <i class="fas fa-external-link-alt"></i> បើកមើលតំណភ្ជាប់
          </a>
          <div class="annex-link-print-url">
            តំណភ្ជាប់៖ ${escapeHtml(primaryUrl)}
          </div>
        </div>
      `;
    } else {
      typeBadges += `<span class="media-type-badge" style="margin-right: 4px; background: #64748b; color: white;"><i class="fas fa-file-alt"></i> ឯកសារ</span>`;
      mediaVisualHtml = `
        <div class="annex-photo-img-wrap" style="display: flex; flex-direction: column; align-items: center; justify-content: center; background: #f8fafc; color: #64748b;">
          <i class="fas fa-file-alt" style="font-size: 3rem; color: #0d9488; margin-bottom: 0.5rem;"></i>
          <div style="font-weight: bold; font-size: 0.95rem;">ឯកសារភស្តុតាង</div>
        </div>
      `;
    }

    if (hasLink && hasFile) {
      const itemLinks = getItemLinks(item);
      typeBadges += `<span class="media-type-badge badge-link" style="margin-right: 4px;"><i class="fas fa-link"></i> ${itemLinks.length > 1 ? `តំណភ្ជាប់ (${khmerNumber(itemLinks.length)})` : 'តំណភ្ជាប់'}</span>`;
    }

    // Inline link box if this card has a file AND also a link, or multiple links
    let inlineLinkHtml = '';
    const itemLinks = getItemLinks(item);
    if (itemLinks.length > 0) {
      if (itemLinks.length === 1 && hasFile) {
        const firstLkUrl = typeof itemLinks[0] === 'string' ? itemLinks[0] : (itemLinks[0]?.url || item.linkUrl);
        const lkPlat = getLinkPlatformInfo(firstLkUrl);
        inlineLinkHtml = `
          <div class="annex-card-inline-link">
            <div class="inline-link-left">
              <i class="${lkPlat.icon}" style="color: ${lkPlat.color}; font-size: 1rem;"></i>
              <a href="${escapeHtml(firstLkUrl)}" target="_blank" rel="noopener noreferrer" class="inline-link-anchor" title="${escapeHtml(firstLkUrl)}">
                ${escapeHtml(firstLkUrl)}
              </a>
            </div>
            <div class="inline-link-actions no-print">
              <a href="${escapeHtml(firstLkUrl)}" target="_blank" rel="noopener noreferrer" class="btn-inline-link-open" title="បើកតំណភ្ជាប់">
                <i class="fas fa-external-link-alt"></i> បើកលីង
              </a>
              <button type="button" class="btn-inline-link-copy" onclick="copyEvidenceLink('${escapeHtml(firstLkUrl)}')" title="ចម្លងលីង">
                <i class="fas fa-copy"></i>
              </button>
            </div>
          </div>
          <div class="annex-link-print-url">
            តំណភ្ជាប់ពិនិត្យបន្ថែម (${escapeHtml(lkPlat.name)}) ៖ ${escapeHtml(firstLkUrl)}
          </div>
        `;
      } else if (itemLinks.length > 1) {
        let multiRowsHtml = '';
        itemLinks.forEach((lk, lkIdx) => {
          const lkUrl = typeof lk === 'string' ? lk : (lk?.url || '');
          if (!lkUrl.trim()) return;
          const lkPlat = getLinkPlatformInfo(lkUrl);
          multiRowsHtml += `
            <div class="annex-multi-link-row">
              <div class="annex-multi-link-info">
                <span class="slot-link-index-badge" style="font-size: 0.68rem; padding: 1px 5px;">លីង ${khmerNumber(lkIdx + 1)}</span>
                <i class="${lkPlat.icon}" style="color: ${lkPlat.color}; font-size: 0.9rem;"></i>
                <a href="${escapeHtml(lkUrl)}" target="_blank" rel="noopener noreferrer" class="annex-multi-link-anchor" title="${escapeHtml(lkUrl)}">
                  ${escapeHtml(lkUrl)}
                </a>
              </div>
              <div class="inline-link-actions no-print" style="display: flex; gap: 4px;">
                <a href="${escapeHtml(lkUrl)}" target="_blank" rel="noopener noreferrer" class="btn-inline-link-open" title="បើកតំណភ្ជាប់" style="padding: 2px 7px; font-size: 0.72rem;">
                  <i class="fas fa-external-link-alt"></i>
                </a>
                <button type="button" class="btn-inline-link-copy" onclick="copyEvidenceLink('${escapeHtml(lkUrl)}')" title="ចម្លងលីង" style="padding: 2px 7px; font-size: 0.72rem;">
                  <i class="fas fa-copy"></i>
                </button>
              </div>
            </div>
            <div class="annex-link-print-url">
              តំណភ្ជាប់ទី ${khmerNumber(lkIdx + 1)} (${escapeHtml(lkPlat.name)}) ៖ ${escapeHtml(lkUrl)}
            </div>
          `;
        });
        inlineLinkHtml = `
          <div class="annex-multi-links-card">
            <div style="font-weight: 700; font-size: 0.8rem; color: #1e293b; display: flex; align-items: center; gap: 5px;">
              <i class="fas fa-link" style="color: #2563eb;"></i> តំណភ្ជាប់ភស្តុតាង (${khmerNumber(itemLinks.length)}) ៖
            </div>
            ${multiRowsHtml}
          </div>
        `;
      }
    }

    // Card Action Buttons
    let actionsHtml = '<div class="annex-card-actions no-print">';
    if (isImg) {
      actionsHtml += `
        <button type="button" class="btn-annex-tool" onclick="openMediaViewerByDocCode('${sub.code}', ${idx})">
          <i class="fas fa-search-plus"></i> ពង្រីកមើល
        </button>
        <button type="button" class="btn-annex-tool" onclick="downloadEvidenceItem('${sub.code}', ${idx})">
          <i class="fas fa-download"></i> ទាញយក
        </button>
      `;
    } else if (isPdf) {
      actionsHtml += `
        <button type="button" class="btn-annex-tool" onclick="openMediaViewerByDocCode('${sub.code}', ${idx})">
          <i class="fas fa-eye"></i> បើកមើល PDF
        </button>
        <button type="button" class="btn-annex-tool" onclick="downloadEvidenceItem('${sub.code}', ${idx})">
          <i class="fas fa-download"></i> ទាញយក
        </button>
      `;
    }
    if (hasLink) {
      actionsHtml += `
        <a href="${escapeHtml(item.linkUrl)}" target="_blank" rel="noopener noreferrer" class="btn-annex-tool" style="text-decoration: none;">
          <i class="fas fa-external-link-alt"></i> បើកលីង
        </a>
        <button type="button" class="btn-annex-tool" onclick="copyEvidenceLink('${escapeHtml(item.linkUrl)}')">
          <i class="fas fa-copy"></i> ចម្លងលីង
        </button>
      `;
    }
    actionsHtml += '</div>';

    const singleCardHtml = `
      <div class="annex-photo-card ${isPdf ? 'annex-pdf-card' : ''} ${(!hasFile && hasLink) ? 'annex-link-card' : ''}">
        ${mediaVisualHtml}
        <div class="annex-photo-caption-wrap">
          <div class="annex-photo-title">
            ${typeBadges}
            ឯកសារទី ${khmerNumber(idx + 1)}៖ ${escapeHtml(item.caption || (isPdf ? item.fileName : (plat ? plat.name : 'ភស្តុតាងបញ្ជាក់')))}
          </div>
          ${item.date ? `<div class="annex-photo-date"><i class="far fa-calendar-alt"></i> ${escapeHtml(formatKhmerMonthNameToWord(item.date))}</div>` : ''}
          ${item.description ? `<div class="annex-photo-desc">${escapeHtml(item.description)}</div>` : ''}
          ${inlineLinkHtml}
          ${actionsHtml}
        </div>
      </div>
    `;

    cardItems.push(singleCardHtml);
  });

  // Group cards into 2-card rows to guarantee perfect alignment and prevent any dropping/scatter bugs in print
  let rowsHtml = '';
  for (let i = 0; i < cardItems.length; i += 2) {
    const card1 = cardItems[i];
    const card2 = cardItems[i + 1];
    rowsHtml += `
      <div class="annex-cards-row">
        ${card1}
        ${card2 ? card2 : '<div class="annex-photo-card annex-card-spacer" aria-hidden="true"></div>'}
      </div>
    `;
  }

  return `
    <div class="annex-page-header">
      <div class="annex-doc-badge-line no-print">
        <span class="annex-sub-badge">[សូចនាករ ${sub.code}]</span>
        <span class="annex-sub-name">${escapeHtml(sub.name)}</span>
      </div>
      <h2 class="annex-doc-title">ឧបសម្ព័ន្ធ ៖ ភស្តុតាងបញ្ជាក់ និងរូបភាពសកម្មភាព</h2>
      <div class="annex-print-sub-header">
        ស្តង់ដាសាលាបឋមសិក្សាគំរូ | សូចនាកររង ${sub.code} ៖ ${escapeHtml(sub.name)} (សាលាបឋមសិក្សាតាពីង)
      </div>
    </div>

    <div class="annex-photos-grid">
      ${rowsHtml}
    </div>
  `;
}

function openEvidenceModal(subCode) {
  const sub = getSubIndicatorByCode(subCode);
  if (!sub) return;

  currentEvidenceSubCode = subCode;
  const modal = document.getElementById('evidenceModal');
  const titleEl = document.getElementById('evidenceModalTitle');
  const subTitleEl = document.getElementById('evidenceModalSubTitle');

  if (titleEl) {
    titleEl.innerHTML = `<span style="color: #a7f3d0;">[សូចនាករ ${sub.code}]</span> បញ្ចូលភស្តុតាងបញ្ជាក់`;
  }
  if (subTitleEl) {
    subTitleEl.innerHTML = `<strong>${sub.name}</strong> | ឯកសារតម្រូវ៖ ${sub.col2Docs}`;
  }

  // Load existing evidence list or create 1 empty slot
  currentEvidenceList = getEvidenceList(subCode);
  if (currentEvidenceList.length === 0) {
    currentEvidenceList = [
      {
        id: 'ev_' + Date.now(),
        mode: 'file',
        imageData: '',
        fileName: '',
        linkUrl: '',
        caption: '',
        date: '',
        description: ''
      }
    ];
  }

  renderEvidenceSlots();

  if (modal) modal.classList.add('active');
  document.body.style.overflow = 'hidden';
}

/**
 * បង្ហាញប្រអប់ភស្តុតាងទាំងអស់ក្នុង Modal
 */
function renderEvidenceSlots() {
  const container = document.getElementById('evidenceSlotsContainer');
  if (!container) return;

  container.innerHTML = '';

  currentEvidenceList.forEach((item, index) => {
    const card = document.createElement('div');
    card.className = 'evidence-slot-card';
    card.dataset.index = index;

    const hasMedia = hasMediaFile(item);
    const isPdf = hasMedia && isPdfItem(item);
    const isImg = hasMedia && !isPdf;
    const hasLink = hasLinkUrl(item);
    const itemLinks = getItemLinks(item);
    const primaryUrl = typeof itemLinks[0] === 'string' ? itemLinks[0] : (itemLinks[0]?.url || item.linkUrl || '');
    const plat = primaryUrl ? getLinkPlatformInfo(primaryUrl) : null;

    let statusText = '';
    if (hasMedia && hasLink) {
      statusText = `<span class="text-success"><i class="fas fa-check-circle"></i> មានទាំងឯកសារ (${isPdf ? 'PDF' : 'រូបភាព'}) និងតំណ (${khmerNumber(itemLinks.length)} លីង)</span>`;
    } else if (hasMedia) {
      if (isPdf) {
        statusText = '<span class="text-success" style="color: #dc2626;"><i class="fas fa-file-pdf"></i> បានភ្ជាប់ឯកសារ PDF រួចរាល់</span>';
      } else {
        statusText = '<span class="text-success"><i class="fas fa-image"></i> បានជ្រើសរូបភាពរួចរាល់</span>';
      }
    } else if (hasLink) {
      statusText = `<span class="text-success" style="color: #2563eb;"><i class="${plat ? plat.icon : 'fas fa-link'}"></i> បានភ្ជាប់តំណលីង (${khmerNumber(itemLinks.length)} លីង)</span>`;
    } else {
      statusText = '<span class="text-warning"><i class="fas fa-clock"></i> រង់ចាំការបញ្ចូលភស្តុតាង</span>';
    }

    let previewContent = '';
    if (hasMedia) {
      if (isPdf) {
        previewContent = `
          <div class="slot-pdf-preview" onclick="event.stopPropagation(); openMediaViewer(${index}, 'slot')">
            <i class="fas fa-file-pdf slot-pdf-icon"></i>
            <div class="slot-pdf-filename" title="${escapeHtml(item.fileName || 'ឯកសារ PDF')}">${escapeHtml(item.fileName || 'ឯកសារ PDF ភស្តុតាង')}</div>
            <div class="slot-pdf-meta">
              <span class="media-type-badge badge-pdf"><i class="fas fa-file-pdf"></i> PDF</span>
              ${item.fileSize ? `<span style="color: #64748b; font-size: 0.75rem;">(${escapeHtml(item.fileSize)})</span>` : ''}
            </div>
            <div class="slot-pdf-actions">
              <button type="button" class="btn-slot-action-view" onclick="event.stopPropagation(); openMediaViewer(${index}, 'slot')">
                <i class="fas fa-eye"></i> បើកមើល PDF
              </button>
              <button type="button" class="btn-slot-action-change" onclick="event.stopPropagation(); triggerSlotFileInput(${index})">
                <i class="fas fa-sync-alt"></i> ប្តូរ
              </button>
              <button type="button" class="btn-slot-clear-file" onclick="event.stopPropagation(); clearSlotFile(${index})" title="លុបឯកសារ PDF ចេញ">
                <i class="fas fa-trash-alt"></i> លុបចេញ
              </button>
            </div>
          </div>
        `;
      } else {
        previewContent = `
          <div class="slot-img-preview" onclick="event.stopPropagation(); openMediaViewer(${index}, 'slot')">
            <img src="${item.imageData}" alt="Preview">
            <div class="slot-img-controls">
              <button type="button" class="btn-slot-img-tool" onclick="event.stopPropagation(); openMediaViewer(${index}, 'slot')">
                <i class="fas fa-search-plus"></i> ពង្រីក
              </button>
              <button type="button" class="btn-slot-img-tool" onclick="event.stopPropagation(); triggerSlotFileInput(${index})">
                <i class="fas fa-camera"></i> ប្តូរ
              </button>
              <button type="button" class="btn-slot-img-tool btn-slot-clear-file" onclick="event.stopPropagation(); clearSlotFile(${index})" title="លុបរូបភាពចេញ">
                <i class="fas fa-trash-alt"></i> លុបចេញ
              </button>
            </div>
          </div>
        `;
      }
    } else {
      previewContent = `
        <div class="slot-empty-prompt">
          <i class="fas fa-cloud-arrow-up slot-cloud-icon"></i>
          <div class="slot-prompt-title">ចុចទីនេះដើម្បីជ្រើសរូបភាព ឬឯកសារ PDF</div>
          <div class="slot-prompt-sub">គាំទ្រ៖ JPG, PNG, WEBP, PDF (រហូតដល់ 25MB)</div>
        </div>
      `;
    }

    const uploadSideHtml = `
      <div style="margin-bottom: 0.45rem; display: flex; align-items: center; justify-content: space-between;">
        <span style="font-weight: 600; font-size: 0.82rem; color: #475569;">
          <i class="fas fa-paperclip" style="color: #0d9488;"></i> ឯកសារភ្ជាប់ (រូបភាព ឬ PDF)
        </span>
        ${hasMedia ? `<span style="font-size: 0.72rem; color: #0d9488; font-weight: 600;"><i class="fas fa-check"></i> មានឯកសារ</span>` : `<span style="font-size: 0.72rem; color: #94a3b8;">(ជាជម្រើស)</span>`}
      </div>
      <input type="file" id="evFileInput_${index}" accept="image/*,.pdf" style="display: none;" onchange="onEvidenceFileChosen(event, ${index})">
      <div class="slot-dropzone ${hasMedia ? 'has-image' : ''}" id="dropzone_${index}" onclick="triggerSlotFileInput(${index})">
        ${previewContent}
      </div>
    `;

    card.innerHTML = `
      <div class="slot-header">
        <div class="slot-title">
          <span class="slot-number-badge">ភស្តុតាងទី ${khmerNumber(index + 1)}</span>
          <span class="slot-status-text">
            ${statusText}
          </span>
        </div>
        ${currentEvidenceList.length > 1 ? `
          <button type="button" class="btn-delete-slot" onclick="removeEvidenceSlot(${index})" title="លុបប្រអប់នេះ">
            <i class="fas fa-trash-alt"></i> លុបប្រអប់នេះ
          </button>
        ` : ''}
      </div>
      <div class="slot-body">
        <div class="slot-upload-side">
          ${uploadSideHtml}
        </div>
        <div class="slot-info-side">
          <div class="slot-form-group">
            <label>ចំណងជើងភស្តុតាង <span class="required">*</span></label>
            <input type="text" class="slot-caption-input" id="evCaption_${index}" value="${escapeHtml(item.caption || '')}" placeholder="ឧ. រូបភាពទីចាត់ការ ឬ របាយការណ៍សាលាគំរូ">
          </div>
          <div class="slot-form-group">
            <label>កាលបរិច្ឆេទ</label>
            <div class="slot-date-field-wrap">
              <input type="text" class="slot-date-input" id="evDate_${index}" value="${escapeHtml(formatKhmerMonthNameToWord(item.date || ''))}" placeholder="ឧ. ថ្ងៃទី ១៥ ខែកញ្ញា ឆ្នាំ ២០២៣" onblur="onSlotDateBlurred(${index})" oninput="onSlotDateInput(${index})">
              <button type="button" class="btn-slot-date-picker" onclick="triggerSlotDatePicker(${index})" title="ចុចជ្រើសរើសកាលបរិច្ឆេទពីប្រតិទិន">
                <i class="fas fa-calendar-alt"></i> ប្រតិទិន
              </button>
              <input type="date" id="evDateHidden_${index}" class="native-date-input-hidden no-print" onchange="onSlotNativeDatePicked(${index}, this.value)">
            </div>
          </div>
          <div class="slot-form-group">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.35rem;">
              <label style="margin-bottom: 0;"><i class="fas fa-link" style="color: #2563eb;"></i> តំណភ្ជាប់ភស្តុតាង (Google Drive / Docs / Sheets / Web - គាំទ្ររហូតដល់ ៥) ៖</label>
              <span style="font-size: 0.72rem; color: #2563eb; font-weight: 600;">(ផ្ទុកបាន ៥ លីង)</span>
            </div>
            <div class="slot-multi-links-container" id="slotMultiLinks_${index}">
              ${renderSlotMultiLinksHtml(item, index)}
            </div>
            <div class="slot-link-help" style="font-size: 0.76rem; color: #64748b; margin-top: 4px;">
              <i class="fas fa-info-circle"></i> អាចភ្ជាប់ឯកសារ Google Drive, Docs, Sheets, YouTube, Facebook... (បន្ថែមលើរូបភាព ឬប្រើតែលីងក៏បាន)
            </div>
          </div>
          <div class="slot-form-group">
            <label>ការពិពណ៌នា / ចំណាំបន្ថែម</label>
            <textarea class="slot-desc-input" id="evDesc_${index}" rows="2" placeholder="ព័ត៌មានលម្អិតបញ្ជាក់ពីសកម្មភាព ឬឯកសារភស្តុតាង...">${escapeHtml(item.description || '')}</textarea>
          </div>
        </div>
      </div>
    `;

    container.appendChild(card);

    // Attach drag and drop listeners (supports multiple files at once!)
    const dropzone = card.querySelector(`#dropzone_${index}`);
    if (dropzone) {
      dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
      });
      dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
      });
      dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
          if (e.dataTransfer.files.length > 1) {
            handleMultipleEvidenceFiles(e.dataTransfer.files);
          } else {
            handleSlotFile(e.dataTransfer.files[0], index);
          }
        }
      });
    }
  });

  // Update summary counter at bottom
  const counterEl = document.getElementById('evidenceSummaryCounter');
  if (counterEl) {
    const fileCount = currentEvidenceList.filter(it => it && it.imageData).length;
    let totalLinks = 0;
    currentEvidenceList.forEach(it => {
      totalLinks += getItemLinks(it).filter(l => l && (typeof l === 'string' ? l.trim() : (l.url && l.url.trim()))).length;
    });
    counterEl.innerHTML = `<i class="fas fa-layer-group" style="color: #0d9488;"></i> ភស្តុតាងសរុប៖ <strong>${khmerNumber(currentEvidenceList.length)}</strong> (ឯកសារ/PDF៖ <span style="color: #0d9488;">${khmerNumber(fileCount)}</span>, តំណលីង៖ <span style="color: #2563eb;">${khmerNumber(totalLinks)}</span>)`;
  }
}

/**
 * បង្កើតកូដ HTML នៃបញ្ជី Multi-Links ក្នុងស្លុតមួយ (គាំទ្ររហូតដល់ ៥ លីង)
 */
function renderSlotMultiLinksHtml(item, slotIndex) {
  let itemLinks = getItemLinks(item);
  if (itemLinks.length === 0) itemLinks = [{ url: '' }];

  let rows = '';
  itemLinks.forEach((lk, lkIdx) => {
    const urlVal = typeof lk === 'string' ? lk : (lk.url || '');
    const hasThisLink = Boolean(urlVal && urlVal.trim());
    const plat = hasThisLink ? getLinkPlatformInfo(urlVal) : null;
    rows += `
      <div class="slot-link-row-item">
        <span class="slot-link-index-badge">លីង ${khmerNumber(lkIdx + 1)}</span>
        <input type="url" class="slot-link-input" id="evLink_${slotIndex}_${lkIdx}" value="${escapeHtml(urlVal)}" placeholder="https://drive.google.com/... ឬ https://..." oninput="onSlotMultiLinkChanged(${slotIndex}, ${lkIdx}, this.value)" onblur="onSlotMultiLinkBlurred(${slotIndex})">
        ${hasThisLink ? `
          <span class="annex-link-platform-badge" style="background: ${plat.color}; color: white; font-size: 0.7rem; padding: 2px 6px; white-space: nowrap;"><i class="${plat.icon}"></i> ${escapeHtml(plat.name)}</span>
          <a href="${escapeHtml(urlVal)}" target="_blank" rel="noopener noreferrer" class="btn-test-link" title="សាកល្បងបើកតំណភ្ជាប់">
            <i class="fas fa-external-link-alt"></i>
          </a>
          <button type="button" class="btn-test-link" onclick="copyEvidenceLink('${escapeHtml(urlVal)}')" title="ចម្លងលីង" style="background: #f1f5f9; color: #475569; border: 1px solid #cbd5e1;">
            <i class="fas fa-copy"></i>
          </button>
        ` : ''}
        ${itemLinks.length > 1 ? `
          <button type="button" class="btn-slot-link-remove" onclick="removeSlotLink(${slotIndex}, ${lkIdx})" title="លុបតំណលីងនេះ">
            <i class="fas fa-times"></i>
          </button>
        ` : ''}
      </div>
    `;
  });

  if (itemLinks.length < 5) {
    rows += `
      <button type="button" class="btn-add-another-link" onclick="addAnotherLinkToSlot(${slotIndex})">
        <i class="fas fa-plus"></i> បន្ថែមតំណលីងទី ${khmerNumber(itemLinks.length + 1)} (រហូតដល់ ៥)
      </button>
    `;
  }
  return rows;
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលវាយ ឬកែប្រែតំណលីងណាមួយក្នុង Multi-Links
 */
function onSlotMultiLinkChanged(slotIndex, linkIndex, val) {
  const item = currentEvidenceList[slotIndex];
  if (!item) return;
  if (!Array.isArray(item.links)) {
    item.links = item.linkUrl ? [{ url: item.linkUrl }] : [];
  }
  while (item.links.length <= linkIndex) {
    item.links.push({ url: '' });
  }
  item.links[linkIndex] = { url: val.trim() };
  item.linkUrl = item.links[0] ? (typeof item.links[0] === 'string' ? item.links[0] : item.links[0].url) : '';

  if (!item.imageData) {
    item.mode = 'link';
  }

  const plat = getLinkPlatformInfo(val.trim());
  if (!item.caption || item.caption.startsWith('តំណភ្ជាប់') || item.caption.startsWith('Google') || item.caption.startsWith('YouTube') || item.caption.startsWith('Facebook')) {
    if (val.trim()) {
      item.caption = `តំណភ្ជាប់ ${plat.name} (ភស្តុតាងទី ${khmerNumber(slotIndex + 1)})`;
      const captionInput = document.getElementById(`evCaption_${slotIndex}`);
      if (captionInput && (!captionInput.value || captionInput.value.startsWith('តំណភ្ជាប់') || captionInput.value.startsWith('Google') || captionInput.value.startsWith('YouTube') || captionInput.value.startsWith('Facebook'))) {
        captionInput.value = item.caption;
      }
    }
  }
  if (!item.date && val.trim()) {
    item.date = getKhmerCurrentDate();
    const dateInput = document.getElementById(`evDate_${slotIndex}`);
    if (dateInput && !dateInput.value) dateInput.value = item.date;
  }
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលបញ្ចប់ការវាយតំណលីង (Blur)
 */
function onSlotMultiLinkBlurred(slotIndex) {
  syncEvidenceInputsFromDom();
  renderEvidenceSlots();
}

/**
 * បន្ថែមតំណលីងមួយទៀតក្នុងស្លុត (រហូតដល់ ៥ លីង)
 */
function addAnotherLinkToSlot(slotIndex) {
  syncEvidenceInputsFromDom();
  const item = currentEvidenceList[slotIndex];
  if (!item) return;
  if (!Array.isArray(item.links)) {
    item.links = item.linkUrl ? [{ url: item.linkUrl }] : [];
  }
  if (item.links.length >= 5) {
    showToast('អាចបញ្ចូលតំណលីងបានអតិបរមាចំនួន ៥ ក្នុងប្រអប់មួយ!');
    return;
  }
  item.links.push({ url: '' });
  renderEvidenceSlots();
  setTimeout(() => {
    const newIdx = item.links.length - 1;
    const input = document.getElementById(`evLink_${slotIndex}_${newIdx}`);
    if (input) input.focus();
  }, 50);
}

/**
 * លុបតំណលីងមួយចេញពីស្លុត
 */
function removeSlotLink(slotIndex, linkIndex) {
  syncEvidenceInputsFromDom();
  const item = currentEvidenceList[slotIndex];
  if (!item || !Array.isArray(item.links)) return;
  item.links.splice(linkIndex, 1);
  if (item.links.length === 0) item.links = [{ url: '' }];
  item.linkUrl = item.links[0] ? (typeof item.links[0] === 'string' ? item.links[0] : item.links[0].url) : '';
  renderEvidenceSlots();
}

/**
 * ធ្វើសមកាលកម្មទិន្នន័យពី Form Inputs ចូលក្នុង currentEvidenceList
 */
function syncEvidenceInputsFromDom() {
  currentEvidenceList.forEach((item, idx) => {
    const captionInput = document.getElementById(`evCaption_${idx}`);
    const dateInput = document.getElementById(`evDate_${idx}`);
    const descInput = document.getElementById(`evDesc_${idx}`);

    if (captionInput) item.caption = captionInput.value.trim();
    if (dateInput) item.date = formatKhmerMonthNameToWord(dateInput.value.trim());
    if (descInput) item.description = descInput.value.trim();

    // Read multi-links from DOM
    const linkItems = [];
    for (let lkIdx = 0; lkIdx < 5; lkIdx++) {
      const linkInput = document.getElementById(`evLink_${idx}_${lkIdx}`);
      if (linkInput) {
        const val = linkInput.value.trim();
        if (val) {
          linkItems.push({ url: val });
        }
      }
    }
    // Fallback if legacy input exists
    if (linkItems.length === 0) {
      const oldLinkInput = document.getElementById(`evLink_${idx}`);
      if (oldLinkInput && oldLinkInput.value.trim()) {
        linkItems.push({ url: oldLinkInput.value.trim() });
      }
    }
    item.links = linkItems;
    item.linkUrl = linkItems.length > 0 ? (typeof linkItems[0] === 'string' ? linkItems[0] : linkItems[0].url) : '';
  });
}

/**
 * លុបឯកសាររូបភាព ឬ PDF ចេញពីប្រអប់ភស្តុតាង
 */
function clearSlotFile(index) {
  syncEvidenceInputsFromDom();
  if (currentEvidenceList[index]) {
    currentEvidenceList[index].imageData = '';
    currentEvidenceList[index].fileName = '';
    currentEvidenceList[index].fileType = '';
    currentEvidenceList[index].fileSize = '';
    currentEvidenceList[index].isPdf = false;
    renderEvidenceSlots();
    showToast('បានលុបឯកសារចេញពីប្រអប់ភស្តុតាង');
  }
}

/**
 * ប្តូរ Mode នៃប្រអប់ភស្តុតាង (សម្រាប់ភាពឆបគ្នា)
 */
function switchSlotMode(index, mode) {
  syncEvidenceInputsFromDom();
  if (currentEvidenceList[index]) {
    currentEvidenceList[index].mode = mode;
    renderEvidenceSlots();
  }
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលវាយ ឬកែប្រែតំណលីងចាស់ (Legacy compatibility)
 */
function onEvidenceLinkChanged(index, val) {
  onSlotMultiLinkChanged(index, 0, val);
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលបញ្ចប់ការវាយតំណលីងចាស់ (Legacy compatibility)
 */
function onEvidenceLinkBlurred(index) {
  onSlotMultiLinkBlurred(index);
}

/**
 * ចម្លងតំណលីងភស្តុតាងទៅ Clipboard
 */
function copyEvidenceLink(url) {
  if (!url) return;
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(url).then(() => {
      showToast('បានចម្លងតំណភ្ជាប់ (URL) ជោគជ័យ!');
    }).catch(() => {
      fallbackCopyText(url);
    });
  } else {
    fallbackCopyText(url);
  }
}

function fallbackCopyText(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.opacity = '0';
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  try {
    document.execCommand('copy');
    showToast('បានចម្លងតំណភ្ជាប់ (URL) ជោគជ័យ!');
  } catch (e) {
    alert('តំណភ្ជាប់៖ ' + text);
  }
  document.body.removeChild(ta);
}

/**
 * បន្ថែមប្រអប់បញ្ចូលភស្តុតាងថ្មីមួយ
 */
function addNewEvidenceSlot() {
  syncEvidenceInputsFromDom();

  currentEvidenceList.push({
    id: 'ev_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4),
    mode: 'file',
    imageData: '',
    fileName: '',
    linkUrl: '',
    links: [],
    caption: '',
    date: '',
    description: ''
  });

  renderEvidenceSlots();

  setTimeout(() => {
    const container = document.getElementById('evidenceSlotsContainer');
    if (container && container.lastElementChild) {
      container.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }, 100);

  showToast('បានបន្ថែមប្រអប់បញ្ចូលភស្តុតាងថ្មី (+)');
}

/**
 * បង្កើតប្រអប់ភស្តុតាងចំនួន ៥ ស្វ័យប្រវត្តិ (Quick 5 Slots)
 */
function createFiveEvidenceSlots() {
  syncEvidenceInputsFromDom();
  const needed = 5 - currentEvidenceList.length;
  if (needed <= 0) {
    showToast(`មានប្រអប់ភស្តុតាងចំនួន ${khmerNumber(currentEvidenceList.length)} រួចរាល់ហើយ!`);
    return;
  }
  for (let i = 0; i < needed; i++) {
    currentEvidenceList.push({
      id: 'ev_' + Date.now() + '_' + i + '_' + Math.random().toString(36).substr(2, 4),
      mode: 'file',
      imageData: '',
      fileName: '',
      linkUrl: '',
      links: [],
      caption: `ភស្តុតាងទី ${khmerNumber(currentEvidenceList.length + 1)}`,
      date: getKhmerCurrentDate(),
      description: ''
    });
  }
  renderEvidenceSlots();
  showToast(`បានរៀបចំប្រអប់ភស្តុតាងចំនួន ៥ ស្វ័យប្រវត្តិរួចរាល់!`);
}

/**
 * បន្ថែមប្រអប់តំណលីងសុទ្ធមួយ (Link-Only Slot)
 */
function addNewLinkEvidenceSlot() {
  syncEvidenceInputsFromDom();
  currentEvidenceList.push({
    id: 'ev_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4),
    mode: 'link',
    imageData: '',
    fileName: '',
    linkUrl: '',
    links: [{ url: '' }],
    caption: `តំណភ្ជាប់ភស្តុតាងទី ${khmerNumber(currentEvidenceList.length + 1)}`,
    date: getKhmerCurrentDate(),
    description: ''
  });
  renderEvidenceSlots();
  setTimeout(() => {
    const container = document.getElementById('evidenceSlotsContainer');
    if (container && container.lastElementChild) {
      container.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'center' });
      const newSlotIdx = currentEvidenceList.length - 1;
      const input = document.getElementById(`evLink_${newSlotIdx}_0`);
      if (input) input.focus();
    }
  }, 100);
  showToast('បានបន្ថែមប្រអប់តំណលីងថ្មី (+)');
}

/**
 * លុបប្រអប់ភស្តុតាងណាមួយ
 */
function removeEvidenceSlot(index) {
  syncEvidenceInputsFromDom();

  if (currentEvidenceList.length <= 1) {
    currentEvidenceList[0] = {
      id: 'ev_' + Date.now(),
      mode: 'file',
      imageData: '',
      fileName: '',
      linkUrl: '',
      links: [],
      caption: '',
      date: '',
      description: ''
    };
  } else {
    currentEvidenceList.splice(index, 1);
  }

  renderEvidenceSlots();
  showToast('បានលុបប្រអប់ភស្តុតាង');
}

/**
 * បើក File Dialog សម្រាប់ជ្រើសឯកសារទោល
 */
function triggerSlotFileInput(index) {
  const input = document.getElementById(`evFileInput_${index}`);
  if (input) input.click();
}

/**
 * បើក File Dialog សម្រាប់ជ្រើសឯកសារច្រើនក្នុងពេលតែមួយ (Multi-Upload ៥+)
 */
function triggerMultiUpload() {
  const input = document.getElementById('multiEvidenceFileInput');
  if (input) input.click();
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលជ្រើសឯកសារច្រើន (Multi-Files)
 */
function onMultiEvidenceFilesChosen(event) {
  const files = event.target.files;
  if (files && files.length > 0) {
    handleMultipleEvidenceFiles(files);
  }
  event.target.value = '';
}

/**
 * ដំណើរការផ្ទុកឯកសារច្រើនក្នុងពេលតែមួយ (Multi-Upload ៥+ រូបភាព/PDF)
 */
function handleMultipleEvidenceFiles(fileList) {
  syncEvidenceInputsFromDom();
  const files = Array.from(fileList);
  if (files.length === 0) return;

  // Check file size max 25MB each
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    if (file.size > 25 * 1024 * 1024) {
      alert(`ឯកសារ «${file.name}» មានទំហំធំពេក (${(file.size / (1024 * 1024)).toFixed(1)} MB)! \nសូមជ្រើសរើសឯកសារក្រោម 25 MB។`);
      return;
    }
  }

  showToast(`កំពុងដំណើរការផ្ទុកឯកសារចំនួន ${khmerNumber(files.length)}...`);

  // Check if first slot is empty and can be replaced
  let startIdx = 0;
  if (currentEvidenceList.length === 1 && !currentEvidenceList[0].imageData && !currentEvidenceList[0].linkUrl && (!currentEvidenceList[0].caption || currentEvidenceList[0].caption.startsWith('រូបភាពភស្តុតាង') || currentEvidenceList[0].caption.startsWith('ឯកសារភស្តុតាង') || currentEvidenceList[0].caption.startsWith('ភស្តុតាងទី'))) {
    startIdx = 0;
  } else {
    startIdx = currentEvidenceList.length;
  }

  let processedCount = 0;
  files.forEach((file, fIdx) => {
    const targetSlotIdx = startIdx + fIdx;
    const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');

    processAndCompressImage(file, (result) => {
      const slotData = {
        id: 'ev_' + Date.now() + '_' + fIdx + '_' + Math.random().toString(36).substr(2, 4),
        mode: 'file',
        imageData: result.data,
        fileName: result.name,
        fileType: result.fileType || (isPdf ? 'application/pdf' : 'image/jpeg'),
        fileSize: result.size,
        isPdf: isPdf,
        linkUrl: '',
        links: [],
        caption: isPdf 
          ? `ឯកសារភស្តុតាង (PDF) ទី ${khmerNumber(targetSlotIdx + 1)}` 
          : `រូបភាពភស្តុតាងទី ${khmerNumber(targetSlotIdx + 1)}`,
        date: getKhmerCurrentDate(),
        description: ''
      };

      if (targetSlotIdx < currentEvidenceList.length) {
        currentEvidenceList[targetSlotIdx] = slotData;
      } else {
        currentEvidenceList.push(slotData);
      }

      processedCount++;
      if (processedCount === files.length) {
        renderEvidenceSlots();
        showToast(`បានបញ្ចូលឯកសារចំនួន ${khmerNumber(files.length)} ដោយជោគជ័យ!`);
      }
    });
  });
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលជ្រើសរូបភាព ឬ PDF ក្នុងស្លុតទោល
 */
function onEvidenceFileChosen(event, index) {
  const file = event.target.files && event.target.files[0];
  if (file) {
    handleSlotFile(file, index);
  }
}

/**
 * ដំណើរការ និងបង្រួមទំហំរូបភាព (គាំទ្រទំហំរហូតដល់ 25MB)
 */
function handleSlotFile(file, index) {
  syncEvidenceInputsFromDom();

  if (file.size > 25 * 1024 * 1024) {
    alert(`ឯកសារ «${file.name}» មានទំហំធំពេក (${(file.size / (1024 * 1024)).toFixed(1)} MB)! \nសូមជ្រើសរើសឯកសារ PDF ឬរូបភាពដែលមានទំហំក្រោម 25 MB ដើម្បីធានាការរក្សាទុកក្នុងប្រព័ន្ធ។`);
    return;
  }

  const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
  showToast(isPdf ? 'កំពុងផ្ទុក និងដំណើរការឯកសារ PDF...' : 'កំពុងដំណើរការ និងបង្រួមរូបភាព...');

  processAndCompressImage(file, (result) => {
    currentEvidenceList[index].imageData = result.data;
    currentEvidenceList[index].fileName = result.name;
    currentEvidenceList[index].fileType = result.fileType || (isPdf ? 'application/pdf' : 'image/jpeg');
    currentEvidenceList[index].isPdf = isPdf;
    currentEvidenceList[index].fileSize = result.size;

    // ប្រសិនបើចំណងជើងនៅទទេ ដាក់ចំណងជើងលំនាំដើម
    if (!currentEvidenceList[index].caption) {
      currentEvidenceList[index].caption = isPdf 
        ? `ឯកសារភស្តុតាង (PDF) ទី ${khmerNumber(index + 1)}` 
        : `រូបភាពភស្តុតាងទី ${khmerNumber(index + 1)}`;
    }
    if (!currentEvidenceList[index].date) {
      currentEvidenceList[index].date = getKhmerCurrentDate();
    }

    renderEvidenceSlots();
    showToast(isPdf ? `បានភ្ជាប់ឯកសារ PDF ដោយជោគជ័យ! (${result.size})` : `បានបញ្ចូលរូបភាពដោយជោគជ័យ! (${result.size})`);
  });
}

function processAndCompressImage(file, callback) {
  if (!file) return;

  const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');

  if (isPdf || !file.type.startsWith('image/')) {
    // Non-image file (PDF etc.)
    const reader = new FileReader();
    reader.onload = (e) => {
      callback({
        isImage: false,
        isPdf: isPdf,
        fileType: file.type || (isPdf ? 'application/pdf' : 'application/octet-stream'),
        name: file.name,
        size: (file.size / 1024).toFixed(1) + ' KB',
        data: e.target.result
      });
    };
    reader.readAsDataURL(file);
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    const img = new Image();
    img.onload = () => {
      // Compress to max width/height 1200px
      const MAX_WIDTH = 1200;
      const MAX_HEIGHT = 1200;
      let width = img.width;
      let height = img.height;

      if (width > height) {
        if (width > MAX_WIDTH) {
          height = Math.round((height * MAX_WIDTH) / width);
          width = MAX_WIDTH;
        }
      } else {
        if (height > MAX_HEIGHT) {
          width = Math.round((width * MAX_HEIGHT) / height);
          height = MAX_HEIGHT;
        }
      }

      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, width, height);

      // Export as compressed JPEG data URL at 82% quality for great clarity
      const compressedDataUrl = canvas.toDataURL('image/jpeg', 0.82);
      const approxKb = Math.round(compressedDataUrl.length * 0.75 / 1024);

      callback({
        isImage: true,
        isPdf: false,
        fileType: 'image/jpeg',
        name: file.name,
        size: approxKb + ' KB',
        data: compressedDataUrl
      });
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

/**
 * បើកផ្ទាំងប្រតិទិន Date Picker សម្រាប់ប្រអប់ភស្តុតាង
 */
function triggerSlotDatePicker(index) {
  const hiddenInput = document.getElementById(`evDateHidden_${index}`);
  if (!hiddenInput) return;

  const dateInput = document.getElementById(`evDate_${index}`);
  if (dateInput && dateInput.value) {
    const iso = parseKhmerDateToIso(dateInput.value);
    if (iso) {
      hiddenInput.value = iso;
    } else {
      hiddenInput.value = new Date().toISOString().split('T')[0];
    }
  } else {
    hiddenInput.value = new Date().toISOString().split('T')[0];
  }

  if (typeof hiddenInput.showPicker === 'function') {
    hiddenInput.showPicker();
  } else {
    hiddenInput.focus();
    hiddenInput.click();
  }
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលជ្រើសរើសកាលបរិច្ឆេទពីប្រតិទិនក្នុងប្រអប់ភស្តុតាង
 */
function onSlotNativeDatePicked(index, isoDateStr) {
  if (!isoDateStr) return;
  const khmerDate = formatKhmerDateFromIso(isoDateStr);
  if (!khmerDate) return;

  const normalized = formatKhmerMonthNameToWord(khmerDate);
  const dateInput = document.getElementById(`evDate_${index}`);
  if (dateInput) {
    dateInput.value = normalized;
  }
  if (currentEvidenceList[index]) {
    currentEvidenceList[index].date = normalized;
  }
  showToast(`បានជ្រើសរើសកាលបរិច្ឆេទ៖ ${normalized}`);
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលអ្នកប្រើប្រាស់បញ្ចប់ការវាយកាលបរិច្ឆេទ (Blur) ក្នុងប្រអប់ភស្តុតាង
 * បំប្លែងលេខខែ «ខែ៩» ទៅជា «ខែកញ្ញា» ភ្លាមៗ
 */
function onSlotDateBlurred(index) {
  const dateInput = document.getElementById(`evDate_${index}`);
  if (!dateInput) return;
  const val = dateInput.value.trim();
  if (val) {
    const formatted = formatKhmerMonthNameToWord(val);
    dateInput.value = formatted;
    if (currentEvidenceList[index]) {
      currentEvidenceList[index].date = formatted;
    }
  }
}

/**
 * ពិនិត្យការវាយអត្ថបទកាលបរិច្ឆេទផ្ទាល់ (Live input)
 */
function onSlotDateInput(index) {
  const dateInput = document.getElementById(`evDate_${index}`);
  if (!dateInput) return;
  const val = dateInput.value;
  // ប្រសិនបើអ្នកប្រើវាយចប់ «ខែ៩ » ឬចុចដកឃ្លាបន្ទាប់ពីលេខខែ
  if (/ខែ\s*([០-៩\d]{1,2})\s+/.test(val)) {
    const formatted = formatKhmerMonthNameToWord(val);
    dateInput.value = formatted;
    if (currentEvidenceList[index]) {
      currentEvidenceList[index].date = formatted;
    }
  }
}

/**
 * រក្សាទុកទិន្នន័យភស្តុតាងទាំងអស់ទៅក្នុង IndexedDB និង Sync ជាមួយ LocalStorage
 */
async function saveEvidenceData() {
  if (!currentEvidenceSubCode) return;

  syncEvidenceInputsFromDom();

  // Filter items that have an image OR a link OR a caption
  const validItems = currentEvidenceList.filter(item => {
    if (!item) return false;
    const hasMedia = Boolean(item.imageData);
    const hasLks = hasLinkUrl(item);
    const hasCap = Boolean(item.caption && item.caption.trim());
    return hasMedia || hasLks || hasCap;
  });

  // 1. រក្សាទុកក្នុង In-Memory Cache ភ្លាមៗ
  evidenceMemoryCache[currentEvidenceSubCode] = validItems;

  // 2. រក្សាទុកក្នុង IndexedDB (ទំហំផ្ទុកធំទូលាយ គាំទ្រឯកសារ ៥+ និង PDF ធំៗ)
  await saveEvidenceToDB(currentEvidenceSubCode, validItems);

  // 3. ព្យាយាមរក្សាទុកក្នុង LocalStorage (តែដក Base64 imageData ចេញ ដើម្បីការពារកុំឱ្យពេញ 5MB LocalStorage)
  try {
    const lightweightItems = validItems.map(item => {
      const copy = { ...item };
      if (copy.imageData && copy.imageData.length > 500) {
        delete copy.imageData;
        copy.hasIndexedDbMedia = true;
      }
      return copy;
    });
    localStorage.setItem('evidence_' + currentEvidenceSubCode, JSON.stringify(lightweightItems));
  } catch (err) {
    console.warn('LocalStorage quota reached; data safely preserved in IndexedDB:', err);
  }

  // Update badge on cards
  updateCardEvidenceBadge(currentEvidenceSubCode);

  // If Document Preview modal is open, refresh Page 2 annex immediately
  const annexPage = document.getElementById('docEvidenceAnnex');
  if (annexPage && currentActiveDoc && currentActiveDoc.code === currentEvidenceSubCode) {
    annexPage.innerHTML = renderEvidenceAnnexHtml(currentActiveDoc);
    if (validItems.length > 0) {
      annexPage.classList.remove('has-no-evidence');
    } else {
      annexPage.classList.add('has-no-evidence');
    }
  }

  showToast(`បានរក្សាទុកភស្តុតាងចំនួន ${khmerNumber(validItems.length)} ដោយជោគជ័យ (IndexedDB)!`);
}

/**
 * បិទផ្ទាំង Evidence Modal
 */
function closeEvidenceModal() {
  const modal = document.getElementById('evidenceModal');
  if (modal) modal.classList.remove('active');

  const docModal = document.getElementById('documentModal');
  if (!docModal || !docModal.classList.contains('active')) {
    document.body.style.overflow = '';
  }
}

/**
 * មើលភស្តុតាងក្នុងផ្ទាំងទម្រង់ឯកសារគំរូ (ទំព័រទី២)
 */
function viewEvidenceInDocModal() {
  saveEvidenceData();
  const subCode = currentEvidenceSubCode;
  closeEvidenceModal();

  if (subCode) {
    openDocumentModal(subCode);
    setTimeout(() => {
      const annexEl = document.getElementById('docEvidenceAnnex') || document.querySelector('.evidence-page-separator');
      if (annexEl) {
        annexEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }, 300);
  }
}

/**
 * Event Listeners សម្រាប់ Evidence Modal
 */
function initEvidenceModalActions() {
  const modal = document.getElementById('evidenceModal');
  if (modal) {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeEvidenceModal();
    });
  }
}

/* ==========================================================================
   BACKUP & RESTORE DATA SYSTEM (ប្រព័ន្ធបម្រុងទុក និងផ្ទេរទិន្នន័យឆ្លងកុំព្យូទ័រ)
   ========================================================================== */

/**
 * នាំចេញទិន្នន័យកែប្រែ និងរូបភាពភស្តុតាងទាំងអស់ទុកជាហ្វាល់ JSON
 */
function backupAllData() {
  const backupObj = {
    version: '1.0',
    system: 'ប្រព័ន្ធគ្រប់គ្រងឯកសារស្តង់ដាសាលាបឋមសិក្សាគំរូ',
    schoolName: 'សាលាបឋមសិក្សាតាពីង',
    exportDate: new Date().toISOString(),
    exportDateKhmer: formatKhmerDateTime(new Date()),
    data: {}
  };

  let count = 0;
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key && (key.startsWith('saved_doc_') || key.startsWith('evidence_'))) {
      backupObj.data[key] = localStorage.getItem(key);
      count++;
    }
  }

  // រួមបញ្ចូលទិន្នន័យភស្តុតាងទាំងអស់ពី IndexedDB / Memory Cache
  if (typeof evidenceMemoryCache === 'object') {
    Object.keys(evidenceMemoryCache).forEach(subCode => {
      const key = 'evidence_' + subCode;
      const items = evidenceMemoryCache[subCode];
      if (items && Array.isArray(items) && items.length > 0) {
        if (!backupObj.data[key]) {
          count++;
        }
        backupObj.data[key] = JSON.stringify(items);
      }
    });
  }

  if (count === 0) {
    showToast('ពុំទាន់មានទិន្នន័យកែប្រែ ឬភស្តុតាងដើម្បីបម្រុងទុកនៅឡើយទេ។');
    return;
  }

  const jsonStr = JSON.stringify(backupObj, null, 2);
  const blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const dateStr = new Date().toISOString().slice(0, 10);
  a.download = `ទិន្នន័យបម្រុង_ស្តង់ដាសាលាតាពីង_${dateStr}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  showToast(`បានទាញយកទិន្នន័យបម្រុងចំនួន ${khmerNumber(count)} ធាតុដោយជោគជ័យ!`);
}

/**
 * បើក File Dialog សម្រាប់រើសហ្វាល់ Backup JSON
 */
function triggerRestoreData() {
  const input = document.getElementById('restoreFileInput');
  if (input) input.click();
}

/**
 * ដាក់បញ្ចូលទិន្នន័យបម្រុងពីហ្វាល់ JSON
 */
function handleRestoreFile(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const backupObj = JSON.parse(e.target.result);
      if (!backupObj || !backupObj.data || typeof backupObj.data !== 'object') {
        alert('ហ្វាល់ទិន្នន័យបម្រុងមិនត្រឹមត្រូវឡើយ!');
        return;
      }

      const keys = Object.keys(backupObj.data);
      if (keys.length === 0) {
        alert('ពុំមានទិន្នន័យនៅក្នុងហ្វាល់បម្រុងនេះឡើយ!');
        return;
      }

      const confirmRestore = confirm(`តើលោកអ្នកពិតជាចង់ដាក់បញ្ចូលទិន្នន័យបម្រុងចំនួន ${khmerNumber(keys.length)} ធាតុនេះមែនទេ? ទិន្នន័យចាស់ៗលើ Browser នេះនឹងត្រូវជំនួសដោយស្វ័យប្រវត្តិ។`);
      if (!confirmRestore) {
        event.target.value = '';
        return;
      }

      let restoredCount = 0;
      for (const key of keys) {
        try {
          localStorage.setItem(key, backupObj.data[key]);
        } catch (err) {
          console.warn('LocalStorage quota limit reached during restore; saving directly to IndexedDB:', err);
        }
        if (key.startsWith('evidence_')) {
          const subCode = key.substring('evidence_'.length);
          try {
            const parsed = JSON.parse(backupObj.data[key]);
            if (Array.isArray(parsed)) {
              evidenceMemoryCache[subCode] = parsed;
              await saveEvidenceToDB(subCode, parsed);
            }
          } catch (err) {}
        }
        restoredCount++;
      }

      alert(`បានដាក់បញ្ចូលទិន្នន័យ និងភស្តុតាងចំនួន ${khmerNumber(restoredCount)} ដោយជោគជ័យ (IndexedDB)! កម្មវិធីនឹង Refresh ទំព័រឡើងវិញ។`);
      window.location.reload();
    } catch (err) {
      alert('កំហុសក្នុងការអានហ្វាល់បម្រុងទុក៖ ' + err.message);
    }
  };
  reader.readAsText(file, 'utf-8');
  event.target.value = '';
}

/* ==========================================================================
   MEDIA VIEWER CONTROLLER (រូបភាព និងឯកសារ PDF ភស្តុតាង)
   ========================================================================== */

let currentViewerMedia = null;

/**
 * បើកផ្ទាំងបង្ហាញរូបភាព ឬឯកសារ PDF ឬតំណលីងពេញលេញ (Lightbox / PDF Reader / Link Viewer)
 */
function openMediaViewer(itemOrIndex, source) {
  let item;
  if (source === 'slot') {
    syncEvidenceInputsFromDom();
    item = currentEvidenceList[itemOrIndex];
  } else {
    item = itemOrIndex;
  }

  if (!item || (!item.imageData && !item.linkUrl)) {
    showToast('មិនមានទិន្នន័យរូបភាព ឯកសារ ឬតំណលីងដើម្បីបើកមើលឡើយ');
    return;
  }

  const hasMedia = hasMediaFile(item);
  const isPdf = hasMedia && isPdfItem(item);
  const isImg = hasMedia && !isPdf;
  const hasLink = hasLinkUrl(item);
  const isLinkOnly = !hasMedia && hasLink;

  const modal = document.getElementById('mediaViewerModal');
  const badgeEl = document.getElementById('mediaViewerBadge');
  const titleEl = document.getElementById('mediaViewerTitle');
  const bodyEl = document.getElementById('mediaViewerBody');
  const dateEl = document.getElementById('mediaViewerDate');
  const descEl = document.getElementById('mediaViewerDesc');
  const openNewTabBtn = document.getElementById('mediaViewerOpenNewTabBtn');
  const downloadBtn = document.getElementById('mediaViewerDownloadBtn');

  if (isLinkOnly) {
    const itemLinks = getItemLinks(item).filter(l => l && (typeof l === 'string' ? l.trim() : (l.url && l.url.trim())));
    const primaryUrl = typeof itemLinks[0] === 'string' ? itemLinks[0] : (itemLinks[0]?.url || item.linkUrl || '');
    const plat = getLinkPlatformInfo(primaryUrl);
    currentViewerMedia = {
      title: item.caption || (itemLinks.length > 1 ? `តំណភ្ជាប់ភស្តុតាង (${khmerNumber(itemLinks.length)} លីង)` : `តំណភ្ជាប់ ${plat.name}`),
      date: item.date || '',
      desc: item.description || '',
      linkUrl: primaryUrl,
      isLink: true,
      isPdf: false,
      plat: plat
    };

    if (badgeEl) {
      badgeEl.className = 'media-type-badge badge-link';
      badgeEl.innerHTML = `<i class="${plat.icon}"></i> ${itemLinks.length > 1 ? `តំណភ្ជាប់ (${khmerNumber(itemLinks.length)})` : escapeHtml(plat.name)}`;
    }

    if (titleEl) titleEl.textContent = currentViewerMedia.title;

    if (bodyEl) {
      if (itemLinks.length > 1) {
        let multiBoxesHtml = '';
        itemLinks.forEach((lk, lkIdx) => {
          const lkUrl = typeof lk === 'string' ? lk : lk.url;
          const lkPlat = getLinkPlatformInfo(lkUrl);
          multiBoxesHtml += `
            <div style="background: white; border: 1px solid #e2e8f0; border-radius: 12px; padding: 1rem 1.25rem; margin-bottom: 0.75rem; text-align: left; display: flex; align-items: center; justify-content: space-between; gap: 10px; box-shadow: 0 2px 6px rgba(0,0,0,0.04);">
              <div style="display: flex; align-items: center; gap: 10px; overflow: hidden; flex: 1;">
                <div style="width: 42px; height: 42px; border-radius: 8px; background: ${lkPlat.bg}; color: ${lkPlat.color}; display: flex; align-items: center; justify-content: center; font-size: 1.3rem; flex-shrink: 0;">
                  <i class="${lkPlat.icon}"></i>
                </div>
                <div style="overflow: hidden; flex: 1;">
                  <div style="font-weight: 700; font-size: 0.9rem; color: #1e293b;">តំណភ្ជាប់ទី ${khmerNumber(lkIdx + 1)} (${escapeHtml(lkPlat.name)})</div>
                  <div style="color: #2563eb; font-family: monospace; font-size: 0.8rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${escapeHtml(lkUrl)}</div>
                </div>
              </div>
              <div style="display: flex; gap: 6px; flex-shrink: 0;">
                <a href="${escapeHtml(lkUrl)}" target="_blank" rel="noopener noreferrer" class="btn-viewer-tool" style="background: #0d9488; color: white; border: none; padding: 0.5rem 0.9rem; font-size: 0.85rem; border-radius: 6px; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; font-weight: 600;">
                  <i class="fas fa-external-link-alt"></i> បើកលីង
                </a>
                <button type="button" class="btn-viewer-tool" onclick="copyEvidenceLink('${escapeHtml(lkUrl)}')" style="background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1; padding: 0.5rem 0.8rem; font-size: 0.85rem; border-radius: 6px; cursor: pointer;">
                  <i class="fas fa-copy"></i>
                </button>
              </div>
            </div>
          `;
        });
        bodyEl.innerHTML = `
          <div style="max-width: 700px; margin: auto; padding: 1rem;">
            <div style="text-align: center; margin-bottom: 1.25rem;">
              <h3 style="font-size: 1.3rem; color: #1e293b; margin-bottom: 0.3rem;">តំណភ្ជាប់ភស្តុតាងទាំងអស់ (${khmerNumber(itemLinks.length)})</h3>
              <p style="color: #64748b; font-size: 0.88rem; margin: 0;">លោកគ្រូ-អ្នកគ្រូ អាចចុចបើកមើលតំណភ្ជាប់នីមួយៗ ឬចម្លងលីងបានងាយស្រួល</p>
            </div>
            ${multiBoxesHtml}
          </div>
        `;
      } else {
        bodyEl.innerHTML = `
          <div style="background: white; border-radius: 16px; padding: 2.5rem 1.5rem; text-align: center; max-width: 650px; margin: auto; box-shadow: 0 10px 25px rgba(0,0,0,0.15);">
            <div style="width: 80px; height: 80px; margin: 0 auto 1.25rem; border-radius: 50%; background: ${plat.bg}; display: flex; align-items: center; justify-content: center; font-size: 2.5rem; color: ${plat.color};">
              <i class="${plat.icon}"></i>
            </div>
            <h3 style="font-size: 1.35rem; font-weight: 700; color: #1e293b; margin-bottom: 0.5rem;">${escapeHtml(plat.name)}</h3>
            <p style="color: #64748b; font-size: 0.95rem; margin-bottom: 1.25rem;">ឯកសារ ឬប្រភពភស្តុតាងត្រូវបានតម្កល់លើសេវាកម្មខាងក្រៅ</p>
            <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 0.85rem 1rem; word-break: break-all; color: #0284c7; font-family: monospace; font-size: 0.95rem; margin-bottom: 1.5rem;">
              ${escapeHtml(primaryUrl)}
            </div>
            <div style="display: flex; gap: 12px; justify-content: center; flex-wrap: wrap;">
              <a href="${escapeHtml(primaryUrl)}" target="_blank" rel="noopener noreferrer" class="btn-viewer-tool" style="background: #0d9488; color: white; border: none; padding: 0.75rem 1.5rem; font-size: 1rem; border-radius: 8px; text-decoration: none; display: inline-flex; align-items: center; gap: 8px; font-weight: 600;">
                <i class="fas fa-external-link-alt"></i> បើកមើលតំណភ្ជាប់ផ្ទាល់ (Open Link)
              </a>
              <button type="button" class="btn-viewer-tool" onclick="copyEvidenceLink('${escapeHtml(primaryUrl)}')" style="background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1; padding: 0.75rem 1.25rem; font-size: 1rem; border-radius: 8px; cursor: pointer; display: inline-flex; align-items: center; gap: 8px; font-weight: 600;">
                <i class="fas fa-copy"></i> ចម្លងលីង (Copy Link)
              </button>
            </div>
          </div>
        `;
      }
    }

    if (downloadBtn) downloadBtn.style.display = 'none';
    if (openNewTabBtn) {
      openNewTabBtn.innerHTML = '<i class="fas fa-external-link-alt"></i> <span class="hide-mobile">បើកលីង</span>';
    }
  } else {
    const mimeType = isPdf ? 'application/pdf' : 'image/jpeg';
    const blobUrl = dataUrlToBlobUrl(item.imageData, mimeType);

    currentViewerMedia = {
      title: item.caption || item.fileName || (isPdf ? 'ឯកសារ PDF ភស្តុតាង' : 'រូបភាពភស្តុតាង'),
      date: item.date || '',
      desc: item.description || '',
      dataUrl: item.imageData,
      fileName: item.fileName || (isPdf ? 'evidence.pdf' : 'evidence.jpg'),
      isPdf: isPdf,
      isLink: false,
      blobUrl: blobUrl
    };

    if (badgeEl) {
      if (isPdf) {
        badgeEl.className = 'media-type-badge badge-pdf';
        badgeEl.innerHTML = '<i class="fas fa-file-pdf"></i> ឯកសារ PDF';
      } else {
        badgeEl.className = 'media-type-badge badge-image';
        badgeEl.innerHTML = '<i class="fas fa-image"></i> រូបភាព';
      }
    }

    if (titleEl) titleEl.textContent = currentViewerMedia.title;

    if (bodyEl) {
      if (isPdf) {
        bodyEl.innerHTML = `
          <iframe src="${blobUrl}#toolbar=1&navpanes=1" class="media-viewer-pdf-frame" title="PDF Viewer">
            <div style="padding: 2.5rem; color: white; text-align: center;">
              <i class="fas fa-file-pdf" style="font-size: 3.5rem; color: #ef4444; margin-bottom: 1rem;"></i>
              <p style="margin-bottom: 1rem; font-size: 1rem;">ឯកសារ PDF នេះមិនអាច Preview ផ្ទាល់ក្នុងទំព័របានឡើយ។</p>
              <button onclick="openMediaInNewTab()" class="btn-viewer-tool">
                <i class="fas fa-external-link-alt"></i> ចុចបើកមើលក្នុងផ្ទាំងថ្មី (Open New Tab)
              </button>
            </div>
          </iframe>
        `;
      } else {
        bodyEl.innerHTML = `
          <img src="${item.imageData}" class="media-viewer-image" alt="${escapeHtml(currentViewerMedia.title)}">
        `;
      }
    }

    if (downloadBtn) downloadBtn.style.display = 'inline-flex';
    if (openNewTabBtn) {
      openNewTabBtn.innerHTML = '<i class="fas fa-external-link-alt"></i> <span class="hide-mobile">បើកផ្ទាំងថ្មី</span>';
    }
  }

  if (dateEl) {
    const displayDate = formatKhmerMonthNameToWord(currentViewerMedia.date);
    dateEl.innerHTML = displayDate ? `<i class="far fa-calendar-alt"></i> ${escapeHtml(displayDate)}` : '';
  }
  if (descEl) {
    let descContent = escapeHtml(currentViewerMedia.desc || '');
    if (hasLink && !isLinkOnly) {
      const p = getLinkPlatformInfo(item.linkUrl);
      descContent += `
        <div style="margin-top: 10px; padding: 8px 12px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; display: flex; align-items: center; justify-content: space-between; gap: 8px; flex-wrap: wrap;">
          <div style="display: flex; align-items: center; gap: 6px; font-size: 0.88rem; color: #1e293b;">
            <i class="${p.icon}" style="color: ${p.color};"></i>
            <span>តំណភ្ជាប់បន្ថែម (${escapeHtml(p.name)}) ៖</span>
            <a href="${escapeHtml(item.linkUrl)}" target="_blank" rel="noopener noreferrer" style="color: #0284c7; text-decoration: underline; font-family: monospace; word-break: break-all;">
              ${escapeHtml(item.linkUrl)}
            </a>
          </div>
          <div style="display: flex; gap: 6px;">
            <a href="${escapeHtml(item.linkUrl)}" target="_blank" rel="noopener noreferrer" class="btn-viewer-tool" style="font-size: 0.8rem; padding: 4px 10px; text-decoration: none; background: #0d9488; color: white; border: none; border-radius: 4px; display: inline-flex; align-items: center; gap: 4px;">
              <i class="fas fa-external-link-alt"></i> បើកលីង
            </a>
            <button type="button" class="btn-viewer-tool" onclick="copyEvidenceLink('${escapeHtml(item.linkUrl)}')" style="font-size: 0.8rem; padding: 4px 10px; background: #e2e8f0; border: 1px solid #cbd5e1; border-radius: 4px; cursor: pointer;">
              <i class="fas fa-copy"></i>
            </button>
          </div>
        </div>
      `;
    }
    descEl.innerHTML = descContent;
  }

  if (modal) modal.classList.add('active');
}

function openMediaViewerByDocCode(subCode, index) {
  const list = getEvidenceList(subCode).filter(item => item && (item.imageData || item.linkUrl || item.caption));
  if (list && list[index]) {
    openMediaViewer(list[index], 'direct');
  }
}

/**
 * បិទផ្ទាំង Media Viewer Modal
 */
function closeMediaViewerModal() {
  const modal = document.getElementById('mediaViewerModal');
  if (modal) modal.classList.remove('active');
  const bodyEl = document.getElementById('mediaViewerBody');
  if (bodyEl) bodyEl.innerHTML = '';

  if (currentViewerMedia && currentViewerMedia.blobUrl && currentViewerMedia.blobUrl.startsWith('blob:')) {
    URL.revokeObjectURL(currentViewerMedia.blobUrl);
  }
  currentViewerMedia = null;
}

/**
 * បើកមើលរូបភាព ឬ PDF ឬតំណលីងក្នុង Browser Tab ថ្មី
 */
function openMediaInNewTab() {
  if (!currentViewerMedia) return;
  if (currentViewerMedia.isLink && currentViewerMedia.linkUrl) {
    window.open(currentViewerMedia.linkUrl, '_blank', 'noopener,noreferrer');
    return;
  }
  if (currentViewerMedia.blobUrl) {
    window.open(currentViewerMedia.blobUrl, '_blank');
  } else if (currentViewerMedia.dataUrl) {
    const win = window.open();
    if (win) {
      if (currentViewerMedia.isPdf) {
        win.document.write(`<iframe src="${currentViewerMedia.dataUrl}" style="border:0; top:0; left:0; bottom:0; right:0; width:100%; height:100%;" allowfullscreen></iframe>`);
      } else {
        win.document.write(`<img src="${currentViewerMedia.dataUrl}" style="max-width:100%; height:auto; display:block; margin:auto;">`);
      }
    }
  }
}

/**
 * ទាញយកឯកសារភស្តុតាងដែលកំពុងមើល
 */
function downloadCurrentMedia() {
  if (!currentViewerMedia || currentViewerMedia.isLink) return;
  const link = document.createElement('a');
  link.href = currentViewerMedia.blobUrl || currentViewerMedia.dataUrl;
  link.download = currentViewerMedia.fileName || (currentViewerMedia.isPdf ? 'evidence.pdf' : 'evidence.jpg');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast('បានទាញយកឯកសារភស្តុតាងជោគជ័យ!');
}

/**
 * ទាញយកឯកសារភស្តុតាងពី Card ឧបសម្ព័ន្ធដោយផ្ទាល់
 */
function downloadEvidenceItem(subCode, index) {
  const list = getEvidenceList(subCode).filter(item => item && (item.imageData || item.linkUrl || item.caption));
  const item = list[index];
  if (!item || !item.imageData) return;
  const isPdf = isPdfItem(item);
  const blobUrl = dataUrlToBlobUrl(item.imageData, isPdf ? 'application/pdf' : 'image/jpeg');
  const link = document.createElement('a');
  link.href = blobUrl;
  link.download = item.fileName || (isPdf ? `${subCode}_evidence_${index + 1}.pdf` : `${subCode}_evidence_${index + 1}.jpg`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
  showToast('បានទាញយកឯកសារភស្តុតាងជោគជ័យ!');
}

// Attach event listeners for Media Viewer Modal
document.addEventListener('DOMContentLoaded', () => {
  const mediaViewerModal = document.getElementById('mediaViewerModal');
  if (mediaViewerModal) {
    mediaViewerModal.addEventListener('click', (e) => {
      if (e.target.id === 'mediaViewerModal') {
        closeMediaViewerModal();
      }
    });
  }
});

// Close media viewer on ESC key
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    const mediaModal = document.getElementById('mediaViewerModal');
    if (mediaModal && mediaModal.classList.contains('active')) {
      closeMediaViewerModal();
    }
  }
});

// Expose globally to window
window.openMediaViewer = openMediaViewer;
window.openMediaViewerByDocCode = openMediaViewerByDocCode;
window.closeMediaViewerModal = closeMediaViewerModal;
window.openMediaInNewTab = openMediaInNewTab;
window.downloadCurrentMedia = downloadCurrentMedia;
window.downloadEvidenceItem = downloadEvidenceItem;
window.switchSlotMode = switchSlotMode;
window.onEvidenceLinkChanged = onEvidenceLinkChanged;
window.copyEvidenceLink = copyEvidenceLink;




window.clearSlotFile = clearSlotFile;
window.onEvidenceLinkBlurred = onEvidenceLinkBlurred;

/* ==========================================================================
   EXCEL & GOOGLE SHEETS IMPORT & AUTO-REPORT GENERATION ENGINE
   ========================================================================== */

let currentImportSubCode = null;
let currentWorkbook = null;
let currentUploadedFile = null;
let currentGoogleSheetUrl = '';
let currentParsedData = null;

/**
 * បើកផ្ទាំង Excel / Google Sheets Import Modal
 */
function openExcelImportModal(subCode, defaultTab) {
  if (!subCode) {
    if (currentActiveDoc) subCode = currentActiveDoc.code;
    else return;
  }

  const sub = getSubIndicatorByCode(subCode);
  if (!sub) return;

  currentImportSubCode = subCode;
  currentWorkbook = null;
  currentUploadedFile = null;
  currentGoogleSheetUrl = '';
  currentParsedData = null;

  // Update Modal Header
  const titleEl = document.getElementById('importModalTitle');
  const subTitleEl = document.getElementById('importModalSubtitle');
  if (titleEl) {
    titleEl.innerHTML = `<span style="color: #86efac;">[សូចនាករ ${sub.code}]</span> នាំចូលទិន្នន័យពី Excel / Google Sheet`;
  }
  if (subTitleEl) {
    subTitleEl.innerHTML = `<strong>${escapeHtml(sub.name)}</strong> | ទម្រង់៖ ${escapeHtml(sub.templateTitle)}`;
  }

  // Reset inputs & preview
  const fileInput = document.getElementById('excelFileInput');
  if (fileInput) fileInput.value = '';
  const fileInfoBar = document.getElementById('excelFileInfoBar');
  if (fileInfoBar) fileInfoBar.style.display = 'none';
  const urlInput = document.getElementById('googleSheetUrlInput');
  if (urlInput) urlInput.value = '';
  const pasteTextarea = document.getElementById('pastedDataTable');
  if (pasteTextarea) pasteTextarea.value = '';
  const previewArea = document.getElementById('importPreviewArea');
  if (previewArea) previewArea.style.display = 'none';
  const btnApply = document.getElementById('btnApplyImport');
  if (btnApply) btnApply.disabled = true;
  const statusMsg = document.getElementById('importStatusMsg');
  if (statusMsg) statusMsg.innerHTML = '';

  // Switch to requested Tab (default to 'file')
  const initialTab = defaultTab || 'file';
  switchImportTab(initialTab);
  if (initialTab === 'paste' && pasteTextarea) {
    setTimeout(() => pasteTextarea.focus(), 150);
  }

  const modal = document.getElementById('excelImportModal');
  if (modal) modal.classList.add('active');
  document.body.style.overflow = 'hidden';

  // Attach Drag & Drop listeners on dropzone
  const dropzone = document.getElementById('excelDropzone');
  if (dropzone && !dropzone.dataset.dragAttached) {
    dropzone.dataset.dragAttached = 'true';
    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        processExcelFile(e.dataTransfer.files[0]);
      }
    });
  }
}

/**
 * បិទផ្ទាំង Excel Import Modal
 */
function closeExcelImportModal() {
  const modal = document.getElementById('excelImportModal');
  if (modal) modal.classList.remove('active');

  const docModal = document.getElementById('documentModal');
  if (!docModal || !docModal.classList.contains('active')) {
    document.body.style.overflow = '';
  }
}

/**
 * ប្តូរ Tab ក្នុងផ្ទាំងនាំចូលទិន្នន័យ (File / Sheet / Paste)
 */
function switchImportTab(tabName) {
  const tabs = ['file', 'sheet', 'paste'];
  tabs.forEach(t => {
    const btn = document.getElementById('tabBtn' + t.charAt(0).toUpperCase() + t.slice(1));
    const content = document.getElementById('tabContent' + t.charAt(0).toUpperCase() + t.slice(1));
    if (btn) {
      if (t === tabName) btn.classList.add('active');
      else btn.classList.remove('active');
    }
    if (content) {
      if (t === tabName) content.style.display = 'block';
      else content.style.display = 'none';
    }
  });
}

/**
 * ចាប់ព្រឹត្តិការណ៍ពេលជ្រើសហ្វាល់ Excel តាម File Input
 */
function handleExcelFileInput(event) {
  const file = event.target.files && event.target.files[0];
  if (file) {
    processExcelFile(file);
  }
}

/**
 * ដំណើរការអាន និង Parse ហ្វាល់ Excel (.xlsx, .xls, .csv)
 */
function processExcelFile(file) {
  if (!file) return;
  currentUploadedFile = file;
  const fileName = file.name;
  const fileSize = (file.size / 1024).toFixed(1) + ' KB';

  const nameEl = document.getElementById('loadedFileName');
  const sizeEl = document.getElementById('loadedFileSize');
  const infoBar = document.getElementById('excelFileInfoBar');
  if (nameEl) nameEl.textContent = fileName;
  if (sizeEl) sizeEl.textContent = fileSize;
  if (infoBar) infoBar.style.display = 'flex';

  const isCsv = fileName.toLowerCase().endsWith('.csv');

  showImportStatus('កំពុងអានទិន្នន័យពីហ្វាល់ «' + fileName + '»...', 'loading');

  const reader = new FileReader();

  if (isCsv || typeof XLSX === 'undefined') {
    reader.onload = (e) => {
      const text = e.target.result;
      const result = parseCsvOrTsv(text);
      if (result.rows.length > 0) {
        renderImportPreview(result.headers, result.rows, 'file');
        showImportStatus(`បានអានហ្វាល់ CSV ជោគជ័យ! រកឃើញ ${khmerNumber(result.rows.length)} ជួរដេក`, 'success');
      } else {
        showImportStatus('មិនមានទិន្នន័យក្នុងហ្វាល់ CSV នេះឡើយ', 'error');
      }
    };
    reader.readAsText(file);
  } else {
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        currentWorkbook = workbook;

        // Populate Sheet Select
        const sheetSelect = document.getElementById('excelSheetSelect');
        const sheetWrap = document.getElementById('sheetSelectorWrap');
        if (sheetSelect) {
          sheetSelect.innerHTML = '';
          workbook.SheetNames.forEach(name => {
            const opt = document.createElement('option');
            opt.value = name;
            opt.textContent = name;
            sheetSelect.appendChild(opt);
          });
        }

        if (sheetWrap) {
          sheetWrap.style.display = workbook.SheetNames.length > 1 ? 'flex' : 'none';
        }

        loadSheetData(workbook.SheetNames[0]);
      } catch (err) {
        console.error('Error parsing Excel:', err);
        showImportStatus('មិនអាចអានហ្វាល់ Excel បានទេ៖ ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }
}

/**
 * ដំណើរការអានទិន្នន័យពី Sheet ណាមួយនៃ Excel Workbook
 */
function onExcelSheetSelected(sheetName) {
  loadSheetData(sheetName);
}

function loadSheetData(sheetName) {
  if (!currentWorkbook) return;
  const sheet = currentWorkbook.Sheets[sheetName];
  if (!sheet) return;

  try {
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rawData || rawData.length === 0) {
      showImportStatus(`សន្លឹកការងារ «${sheetName}» ពុំមានទិន្នន័យឡើយ`, 'error');
      return;
    }

    const cleanData = rawData.filter(row => row && row.some(cell => String(cell).trim() !== ''));
    if (cleanData.length === 0) {
      showImportStatus(`សន្លឹកការងារ «${sheetName}» ពុំមានទិន្នន័យឡើយ`, 'error');
      return;
    }

    let headers = [];
    let rows = [];
    if (cleanData.length === 1) {
      headers = cleanData[0].map((h, i) => `កូឡោនទី ${khmerNumber(i + 1)}`);
      rows = [cleanData[0]];
    } else {
      headers = cleanData[0].map((h, i) => String(h || '').trim() || `កូឡោនទី ${khmerNumber(i + 1)}`);
      rows = cleanData.slice(1);
    }

    renderImportPreview(headers, rows, 'file');
    showImportStatus(`បានអានទិន្នន័យពី Sheet «${sheetName}» ជោគជ័យ! រកឃើញ ${khmerNumber(rows.length)} ជួរដេក`, 'success');
  } catch (err) {
    showImportStatus('កំហុសពេលទាញទិន្នន័យពី Sheet៖ ' + err.message, 'error');
  }
}

/**
 * ទាញយកទិន្នន័យពី Google Sheets តាម URL
 */
function fetchGoogleSheetData() {
  const urlInput = document.getElementById('googleSheetUrlInput');
  const url = urlInput ? urlInput.value.trim() : '';
  if (!url) {
    alert('សូមបញ្ចូលតំណភ្ជាប់ Google Sheet (URL) ជាមុនសិន!');
    return;
  }

  currentGoogleSheetUrl = url;

  // Extract Spreadsheet ID
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!match || !match[1]) {
    alert('តំណភ្ជាប់ Google Sheet មិនត្រឹមត្រូវឡើយ! សូមពិនិត្យមើលតំណភ្ជាប់ឡើងវិញ (ឧ. https://docs.google.com/spreadsheets/d/.../edit)');
    return;
  }

  const sheetId = match[1];

  // Check custom sheet name or GID input
  const sheetNameInput = document.getElementById('googleSheetNameInput');
  const customSheet = sheetNameInput ? sheetNameInput.value.trim() : '';

  let sheetParam = '';
  if (customSheet) {
    if (/^[0-9]+$/.test(customSheet)) {
      sheetParam = `&gid=${customSheet}`;
    } else {
      sheetParam = `&sheet=${encodeURIComponent(customSheet)}`;
    }
  } else {
    const gidMatch = url.match(/[#&?]gid=([0-9]+)/);
    if (gidMatch && gidMatch[1]) {
      sheetParam = `&gid=${gidMatch[1]}`;
    }
  }

  showImportStatus('កំពុងទាញយកទិន្នន័យពី Google Sheet...', 'loading');

  const csvUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv${sheetParam}`;

  fetch(csvUrl)
    .then(res => {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.text();
    })
    .then(csvText => {
      if (!csvText || csvText.includes('<!DOCTYPE html>')) {
        throw new Error('Google Sheet មិនទាន់បានកំណត់សិទ្ធិមើលជាសាធារណៈ (Anyone with link) ឡើយ!');
      }
      const result = parseCsvOrTsv(csvText);
      if (result.rows.length > 0) {
        renderImportPreview(result.headers, result.rows, 'sheet');
        showImportStatus(`បានទាញយកទិន្នន័យពី Google Sheet ជោគជ័យ! (${khmerNumber(result.rows.length)} ជួរដេក)`, 'success');
      } else {
        showImportStatus('មិនមានទិន្នន័យក្នុង Google Sheet នេះឡើយ', 'error');
      }
    })
    .catch(err => {
      console.warn('Direct CSV fetch failed, trying GViz fallback:', err);
      const gvizUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv${sheetParam}`;
      fetch(gvizUrl)
        .then(res => res.text())
        .then(csvText => {
          if (!csvText || csvText.includes('<!DOCTYPE html>')) {
            throw new Error('Private sheet');
          }
          const result = parseCsvOrTsv(csvText);
          if (result.rows.length > 0) {
            renderImportPreview(result.headers, result.rows, 'sheet');
            showImportStatus(`បានទាញយកទិន្នន័យពី Google Sheet ជោគជ័យ! (${khmerNumber(result.rows.length)} ជួរដេក)`, 'success');
          } else {
            throw new Error('Empty sheet');
          }
        })
        .catch(() => {
          showImportStatus('មិនអាចទាញយកដោយស្វ័យប្រវត្តិបានទេ (ដោយសារ Google Sheet មិនទាន់បើកសិទ្ធិសាធារណៈ ឬជាប់ CORS)។ សូមប្រើជម្រើស «ចម្លង & បិទភ្ជាប់» ជំនួសវិញ (ចុច Ctrl+A -> Ctrl+C ក្នុង Sheet រួចមក Paste នៅទីនេះ)។', 'warning');
          setTimeout(() => switchImportTab('paste'), 3000);
        });
    });
}

function processPastedData() {
  const textarea = document.getElementById('pastedDataTable');
  const text = textarea ? textarea.value.trim() : '';
  if (!text) {
    alert('សូមបិទភ្ជាប់ (Paste) ទិន្នន័យក្នុងប្រអប់ជាមុនសិន!');
    return;
  }

  const result = parseCsvOrTsv(text);
  if (result.rows.length > 0) {
    renderImportPreview(result.headers, result.rows, 'paste');
    showImportStatus(`បានដំណើរការទិន្នន័យដែលបានបិទភ្ជាប់ជោគជ័យ! (${khmerNumber(result.rows.length)} ជួរដេក)`, 'success');
  } else {
    showImportStatus('មិនអាចសម្គាល់ជួរដេកទិន្នន័យបានឡើយ សូមពិនិត្យទិន្នន័យដែលបាន Paste', 'error');
  }
}

/**
 * បិទភ្ជាប់ទិន្នន័យពី Clipboard ចូលក្នុង Textarea នៃ Tab 3 ស្វ័យប្រវត្តិ
 */
async function pasteFromClipboardToTextarea() {
  const textarea = document.getElementById('pastedDataTable');
  if (!textarea) return;

  if (navigator.clipboard && navigator.clipboard.readText) {
    try {
      const text = await navigator.clipboard.readText();
      if (text && text.trim()) {
        textarea.value = text;
        processPastedData();
        showImportStatus(`បានបិទភ្ជាប់ទិន្នន័យពី Clipboard រួចរាល់!`, 'success');
        return;
      }
    } catch (err) {
      console.warn('Clipboard readText failed:', err);
    }
  }

  textarea.focus();
  showImportStatus('សូមចុចបញ្ជា Ctrl+V ក្នុងប្រអប់ខាងលើដើម្បីបិទភ្ជាប់ទិន្នន័យ។', 'info');
}

function onPastedDataChanged(val) {
  if (val && val.trim().length > 10) {
    clearTimeout(window._pasteDebounce);
    window._pasteDebounce = setTimeout(() => {
      processPastedData();
    }, 600);
  }
}

/**
 * កំណត់ Placeholder និងប្រភេទបិទភ្ជាប់រហ័ស (តារាងទាំងមូល / មួយជួរឈរ / មួយជួរដេក)
 */
function setPasteModePlaceholder(mode) {
  window._preferredImportMode = mode;
  const textarea = document.getElementById('pastedDataTable');

  // Highlight quick paste chips
  document.querySelectorAll('.quick-paste-btn').forEach(btn => {
    btn.classList.remove('active');
  });
  if (typeof event !== 'undefined' && event && event.currentTarget && event.currentTarget.classList.contains('quick-paste-btn')) {
    event.currentTarget.classList.add('active');
  }

  if (textarea) {
    if (mode === 'column') {
      textarea.placeholder = 'ឧទាហរណ៍ មួយជួរឈរ (Column) ៖ ជ្រើសរើសជួរឈរណាមួយក្នុង Excel (ឧ. ជួរឈរឈ្មោះ ឬពិន្ទុ) ចុច Ctrl+C រួចចុច Ctrl+V នៅទីនេះ...';
    } else if (mode === 'row') {
      textarea.placeholder = 'ឧទាហរណ៍ មួយជួរដេក (Row) ៖ ជ្រើសរើសជួរដេកណាមួយក្នុង Excel ចុច Ctrl+C រួចចុច Ctrl+V នៅទីនេះ...';
    } else {
      textarea.placeholder = 'បិទភ្ជាប់ទិន្នន័យដែលបាន Copy ពី Excel ឬ Google Sheet នៅទីនេះ (តារាងទាំងមូល ឬតែមួយជួរឈរ ឬមួយជួរដេក)...';
    }
    textarea.focus();
  }

  onImportScopeModeChanged(mode);
}

/**
 * មុខងារ Built-in CSV / TSV Parser (ដំណើរការ Offline ១០០% ទាំងពេលគ្មានអ៊ីនធឺណិត)
 * គាំទ្រការបំបែកទាំងតារាងទាំងមូល មួយជួរឈរ ឬមួយជួរដេក
 */
function parseCsvOrTsv(text) {
  if (!text) return { headers: [], rows: [] };

  const firstLine = text.split(/\r?\n/)[0] || '';
  let delimiter = ',';
  if (firstLine.includes('\t')) {
    delimiter = '\t';
  } else if (!firstLine.includes(',') && firstLine.includes(';')) {
    delimiter = ';';
  }

  const rawRows = [];
  let currentRow = [];
  let currentCell = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        currentCell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      currentRow.push(currentCell.trim());
      currentCell = '';
    } else if ((char === '\r' || char === '\n') && !inQuotes) {
      if (char === '\r' && nextChar === '\n') i++;
      currentRow.push(currentCell.trim());
      currentCell = '';
      if (currentRow.some(c => c !== '')) {
        rawRows.push(currentRow);
      }
      currentRow = [];
    } else {
      currentCell += char;
    }
  }

  if (currentCell !== '' || currentRow.length > 0) {
    currentRow.push(currentCell.trim());
    if (currentRow.some(c => c !== '')) {
      rawRows.push(currentRow);
    }
  }

  if (rawRows.length === 0) return { headers: [], rows: [] };

  // ប្រសិនបើបិទភ្ជាប់តែមួយជួរដេកគត់
  if (rawRows.length === 1) {
    const headers = rawRows[0].map((h, idx) => `កូឡោនទី ${khmerNumber(idx + 1)}`);
    const rows = [rawRows[0]];
    return { headers, rows };
  }

  const headers = rawRows[0].map((h, idx) => h || `កូឡោនទី ${khmerNumber(idx + 1)}`);
  const rows = rawRows.slice(1);

  return { headers, rows };
}

/**
 * ទទួលយកបញ្ជីជួរឈរនៃតារាងឯកសារបច្ចុប្បន្ន (សម្រាប់ដាក់ក្នុង Dropdown ជ្រើសរើសជួរឈរគោលដៅ)
 */
function getDocTableColumns(subCode) {
  const sub = getSubIndicatorByCode(subCode);
  if (!sub) return [];

  // 1. ពិនិត្យមើលតារាងដែលកំពុងបើកបង្ហាញផ្ទាល់លើផ្ទាំង
  const liveTable = document.querySelector('#printableDocument .doc-table');
  if (liveTable && currentActiveDoc && currentActiveDoc.code === subCode) {
    const headers = Array.from(liveTable.querySelectorAll('thead th:not(.row-action-col)'));
    if (headers.length > 0) {
      return headers.map((th, idx) => ({
        index: idx,
        name: th.innerText.replace(/[\r\n\t]+/g, ' ').trim() || `កូឡោនទី ${khmerNumber(idx + 1)}`
      }));
    }
  }

  // 2. ពិនិត្យមើលទិន្នន័យឯកសារដែលធ្លាប់រក្សាទុកក្នុង LocalStorage
  const savedHtml = localStorage.getItem('saved_doc_' + subCode);
  if (savedHtml) {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString('<div>' + savedHtml + '</div>', 'text/html');
      const headers = Array.from(doc.querySelectorAll('.doc-table thead th:not(.row-action-col)'));
      if (headers.length > 0) {
        return headers.map((th, idx) => ({
          index: idx,
          name: th.innerText.replace(/[\r\n\t]+/g, ' ').trim() || `កូឡោនទី ${khmerNumber(idx + 1)}`
        }));
      }
    } catch (e) {
      console.warn('Error reading saved doc columns:', e);
    }
  }

  // 3. Fallback ទៅកាន់ headers លំនាំដើមនៃសូចនាករ
  return (sub.headers || []).map((h, idx) => ({
    index: idx,
    name: String(h).replace(/[\r\n\t]+/g, ' ').trim() || `កូឡោនទី ${khmerNumber(idx + 1)}`
  }));
}

/**
 * ទទួលយកបញ្ជីជួរដេកនៃតារាងឯកសារបច្ចុប្បន្ន (សម្រាប់ដាក់ក្នុង Dropdown ជ្រើសរើសជួរដេកគោលដៅ)
 */
function getDocTableRows(subCode) {
  const sub = getSubIndicatorByCode(subCode);
  if (!sub) return [];

  // 1. ពិនិត្យមើលតារាងដែលកំពុងបើកបង្ហាញផ្ទាល់លើផ្ទាំង
  const liveTable = document.querySelector('#printableDocument .doc-table');
  if (liveTable && currentActiveDoc && currentActiveDoc.code === subCode) {
    const trs = Array.from(liveTable.querySelectorAll('tbody tr'));
    if (trs.length > 0) {
      return trs.map((tr, idx) => {
        const cells = Array.from(tr.querySelectorAll('td:not(.row-action-col)')).map(td => td.innerText.trim());
        const summary = cells.filter(Boolean).slice(0, 3).join(' | ');
        return {
          index: idx,
          summary: summary ? `ជួរទី ${khmerNumber(idx + 1)}៖ ${summary.substring(0, 45)}` : `ជួរទី ${khmerNumber(idx + 1)}`
        };
      });
    }
  }

  // 2. ពិនិត្យមើលទិន្នន័យឯកសារដែលធ្លាប់រក្សាទុកក្នុង LocalStorage
  const savedHtml = localStorage.getItem('saved_doc_' + subCode);
  if (savedHtml) {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString('<div>' + savedHtml + '</div>', 'text/html');
      const trs = Array.from(doc.querySelectorAll('.doc-table tbody tr'));
      if (trs.length > 0) {
        return trs.map((tr, idx) => {
          const cells = Array.from(tr.querySelectorAll('td:not(.row-action-col)')).map(td => td.innerText.trim());
          const summary = cells.filter(Boolean).slice(0, 3).join(' | ');
          return {
            index: idx,
            summary: summary ? `ជួរទី ${khmerNumber(idx + 1)}៖ ${summary.substring(0, 45)}` : `ជួរទី ${khmerNumber(idx + 1)}`
          };
        });
      }
    } catch (e) {
      console.warn('Error reading saved doc rows:', e);
    }
  }

  // 3. Fallback ទៅកាន់ sampleRows លំនាំដើមនៃសូចនាករ
  return (sub.sampleRows || []).map((row, idx) => {
    const summary = row.filter(Boolean).slice(0, 3).join(' | ');
    return {
      index: idx,
      summary: summary ? `ជួរទី ${khmerNumber(idx + 1)}៖ ${summary.substring(0, 45)}` : `ជួរទី ${khmerNumber(idx + 1)}`
    };
  });
}

/**
 * បង្ហាញការពិនិត្យមើលទិន្នន័យជាមុន (Live Preview)
 */
function renderImportPreview(headers, rows, source) {
  currentParsedData = { headers, rows, source };

  const previewArea = document.getElementById('importPreviewArea');
  const badgeEl = document.getElementById('previewRowCountBadge');
  const thead = document.getElementById('previewTableHead');
  const tbody = document.getElementById('previewTableBody');
  const btnApply = document.getElementById('btnApplyImport');

  if (badgeEl) {
    badgeEl.textContent = `${khmerNumber(rows.length)} ជួរដេក (${khmerNumber(headers.length)} កូឡោន)`;
  }

  // 1. បំពេញបញ្ជីជួរឈរប្រភព (Source Column Select)
  const selectSourceCol = document.getElementById('selectSourceColumn');
  if (selectSourceCol) {
    selectSourceCol.innerHTML = '';
    headers.forEach((h, idx) => {
      const opt = document.createElement('option');
      opt.value = idx;
      opt.textContent = `កូឡោនទី ${khmerNumber(idx + 1)}៖ ${h}`;
      selectSourceCol.appendChild(opt);
    });
  }

  // 2. បំពេញបញ្ជីជួរឈរគោលដៅ (Target Column Select)
  const selectTargetCol = document.getElementById('selectTargetColumn');
  if (selectTargetCol) {
    selectTargetCol.innerHTML = '';
    const docCols = getDocTableColumns(currentImportSubCode);
    docCols.forEach(col => {
      const opt = document.createElement('option');
      opt.value = col.index;
      opt.textContent = `🔄 ជំនួសជួរឈរទី ${khmerNumber(col.index + 1)}៖ ${col.name}`;
      selectTargetCol.appendChild(opt);
    });

    const newOpt = document.createElement('option');
    newOpt.value = '__NEW__';
    newOpt.textContent = '➕ បង្កើតជាជួរឈរថ្មី (Append New Column)';
    selectTargetCol.appendChild(newOpt);

    // Auto select target column if user clicked a cell before opening modal, or default to column 2 (index 1)
    if (window._lastActiveCell) {
      const tr = window._lastActiveCell.closest('tr');
      if (tr) {
        const cells = Array.from(tr.querySelectorAll('td:not(.row-action-col), th:not(.row-action-col)'));
        const idx = cells.indexOf(window._lastActiveCell);
        if (idx !== -1 && docCols.some(c => c.index === idx)) {
          selectTargetCol.value = idx;
        }
      }
    } else if (docCols.length > 1) {
      selectTargetCol.value = 1;
    }
  }

  // 3. បំពេញបញ្ជីជួរដេកប្រភព (Source Row Select)
  const selectSourceRow = document.getElementById('selectSourceRow');
  if (selectSourceRow) {
    selectSourceRow.innerHTML = '';
    const allOpt = document.createElement('option');
    allOpt.value = 'all';
    allOpt.textContent = `⚡ គ្រប់ជួរដេកទាំងអស់ (ចំនួន ${khmerNumber(rows.length)} ជួរ)`;
    selectSourceRow.appendChild(allOpt);

    rows.forEach((row, idx) => {
      const opt = document.createElement('option');
      opt.value = idx;
      const previewText = row.filter(Boolean).slice(0, 3).join(' | ');
      opt.textContent = `ជួរដេកទី ${khmerNumber(idx + 1)}${previewText ? ' ៖ ' + previewText.substring(0, 35) : ''}`;
      selectSourceRow.appendChild(opt);
    });
  }

  // 4. បំពេញបញ្ជីជួរដេកគោលដៅសម្រាប់ជំនួស (Target Row Index Select)
  const selectTargetRowIndex = document.getElementById('selectTargetRowIndex');
  if (selectTargetRowIndex) {
    selectTargetRowIndex.innerHTML = '';
    const docRows = getDocTableRows(currentImportSubCode);
    docRows.forEach(r => {
      const opt = document.createElement('option');
      opt.value = r.index;
      opt.textContent = r.summary;
      selectTargetRowIndex.appendChild(opt);
    });
  }

  // 5. Smart Auto-detect Import Mode ផ្អែកលើទម្រង់ទិន្នន័យ
  let targetMode = window._preferredImportMode || 'full';
  if (!window._preferredImportMode) {
    if (headers.length === 1 && rows.length > 0) {
      targetMode = 'column';
    } else if (rows.length === 1 && headers.length > 1) {
      targetMode = 'row';
    } else {
      targetMode = 'full';
    }
  }
  onImportScopeModeChanged(targetMode);

  // 6. បង្ហាញក្បាលតារាងពិនិត្យជាមុន (Table Header Preview)
  if (thead) {
    let thHtml = '<tr>';
    headers.forEach((h, i) => {
      thHtml += `<th title="កូឡោនទី ${i + 1}" data-col="${i}">${escapeHtml(h)}</th>`;
    });
    thHtml += '</tr>';
    thead.innerHTML = thHtml;
  }

  // 7. បង្ហាញតួទិន្នន័យពិនិត្យជាមុន (Table Body Preview)
  if (tbody) {
    let trHtml = '';
    const previewRows = rows.slice(0, 10);
    previewRows.forEach((row, rIdx) => {
      trHtml += `<tr data-row="${rIdx}">`;
      headers.forEach((h, cIdx) => {
        const val = row[cIdx] !== undefined ? row[cIdx] : '';
        trHtml += `<td data-col="${cIdx}">${escapeHtml(String(val))}</td>`;
      });
      trHtml += '</tr>';
    });

    if (rows.length > 10) {
      trHtml += `
        <tr class="preview-more-rows">
          <td colspan="${headers.length}" style="text-align: center; color: #64748b; font-style: italic; background: #f1f5f9; padding: 0.6rem;">
            ... នៅសល់ ${khmerNumber(rows.length - 10)} ជួរដេកទៀត ដែលនឹងត្រូវបញ្ចូលទាំងអស់ ...
          </td>
        </tr>
      `;
    }
    tbody.innerHTML = trHtml;
  }

  updatePreviewHighlight();

  if (previewArea) previewArea.style.display = 'flex';
  if (btnApply) btnApply.disabled = false;

  setTimeout(() => {
    if (previewArea) {
      previewArea.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, 100);
}

/**
 * ធ្វើបច្ចុប្បន្នកម្មការបន្លិចពណ៌ (Highlight) លើតារាង Preview តាមទម្រង់នៃការនាំចូល
 */
function updatePreviewHighlight() {
  const currentModeRadio = document.querySelector('input[name="importScopeMode"]:checked');
  const mode = currentModeRadio ? currentModeRadio.value : 'full';

  const previewTable = document.getElementById('importPreviewTable');
  if (!previewTable) return;

  // Clear previous highlights
  previewTable.querySelectorAll('.col-highlight, .row-highlight').forEach(el => {
    el.classList.remove('col-highlight', 'row-highlight');
  });

  if (mode === 'column') {
    const srcColSelect = document.getElementById('selectSourceColumn');
    const colIdx = srcColSelect ? parseInt(srcColSelect.value || '0', 10) : 0;
    previewTable.querySelectorAll(`[data-col="${colIdx}"]`).forEach(el => {
      el.classList.add('col-highlight');
    });
  } else if (mode === 'row') {
    const srcRowSelect = document.getElementById('selectSourceRow');
    const rowVal = srcRowSelect ? srcRowSelect.value : 'all';
    if (rowVal === 'all') {
      previewTable.querySelectorAll('tbody tr:not(.preview-more-rows)').forEach(tr => {
        tr.classList.add('row-highlight');
      });
    } else {
      const rowIdx = parseInt(rowVal, 10);
      const tr = previewTable.querySelector(`tbody tr[data-row="${rowIdx}"]`);
      if (tr) tr.classList.add('row-highlight');
    }
  }
}

/**
 * ប្តូរទម្រង់នាំចូលទិន្នន័យ (តារាងទាំងមូល / ម្តងមួយជួរឈរ / ម្តងមួយជួរដេក)
 */
function onImportScopeModeChanged(mode) {
  const modes = ['full', 'column', 'row'];
  modes.forEach(m => {
    const radio = document.getElementById(m === 'full' ? 'modeFullTable' : m === 'column' ? 'modeColumn' : 'modeRow');
    const lbl = document.getElementById(m === 'full' ? 'lblModeFullTable' : m === 'column' ? 'lblModeColumn' : 'lblModeRow');
    const cfg = document.getElementById(m === 'full' ? 'configFullTable' : m === 'column' ? 'configColumn' : 'configRow');
    const isCur = (m === mode);
    if (radio) radio.checked = isCur;
    if (lbl) {
      if (isCur) lbl.classList.add('active');
      else lbl.classList.remove('active');
    }
    if (cfg) {
      cfg.style.display = isCur ? 'block' : 'none';
    }
  });

  const btnApply = document.getElementById('btnApplyImport');
  if (btnApply) {
    if (mode === 'column') {
      btnApply.innerHTML = '<i class="fas fa-columns"></i> ⚡ ចម្លងទិន្នន័យចូលក្នុងជួរឈរ';
    } else if (mode === 'row') {
      btnApply.innerHTML = '<i class="fas fa-grip-lines"></i> ⚡ ចម្លងទិន្នន័យចូលក្នុងជួរដេក';
    } else {
      btnApply.innerHTML = '<i class="fas fa-bolt"></i> ⚡ ទាញទិន្នន័យបង្កើតជារបាយការណ៍ស្វ័យប្រវត្តិ';
    }
  }

  updatePreviewHighlight();
}

function onSourceColumnChanged() {
  updatePreviewHighlight();
}

function onTargetColumnChanged() {
  const targetColSelect = document.getElementById('selectTargetColumn');
  const chkColHeader = document.getElementById('chkColFirstRowIsHeader');
  if (targetColSelect && chkColHeader) {
    if (targetColSelect.value === '__NEW__') {
      chkColHeader.checked = true;
    }
  }
}

function onSourceRowChanged() {
  updatePreviewHighlight();
}

function onTargetRowActionChanged() {
  const actionSelect = document.getElementById('selectTargetRowAction');
  const wrapIndex = document.getElementById('wrapTargetRowIndex');
  if (actionSelect && wrapIndex) {
    wrapIndex.style.display = (actionSelect.value === 'replace_specific') ? 'flex' : 'none';
  }
}

function onColFirstRowHeaderToggled() {
  // Trigger update on preview if needed
}

function onRowFirstRowHeaderToggled() {
  updatePreviewHighlight();
}

/**
 * បង្ហាញសារជូនដំណឹងក្នុងផ្ទាំង Import
 */
function showImportStatus(msg, type) {
  const el = document.getElementById('importStatusMsg');
  if (!el) return;
  if (!msg) {
    el.innerHTML = '';
    return;
  }
  let color = '#0284c7';
  let icon = 'fa-info-circle';
  if (type === 'success') {
    color = '#16a34a';
    icon = 'fa-check-circle';
  } else if (type === 'error') {
    color = '#dc2626';
    icon = 'fa-exclamation-circle';
  } else if (type === 'warning') {
    color = '#d97706';
    icon = 'fa-exclamation-triangle';
  } else if (type === 'loading') {
    color = '#2563eb';
    icon = 'fa-spinner fa-spin';
  }
  el.innerHTML = `<span style="color: ${color}; display: inline-flex; align-items: center; gap: 6px;"><i class="fas ${icon}"></i> ${escapeHtml(msg)}</span>`;
}

/**
 * មុខងារបង្កើតរបាយការណ៍ / ចម្លងទិន្នន័យពី Excel ឬ Google Sheet ចូលក្នុងឯកសារ
 * គាំទ្រទាំង ៣ ទម្រង់៖ (១) តារាងទាំងមូល (២) ម្តងមួយជួរឈរ (៣) ម្តងមួយជួរដេក
 */
function applyImportedDataToDocument() {
  if (!currentParsedData || !currentImportSubCode) return;

  const sub = getSubIndicatorByCode(currentImportSubCode);
  if (!sub) return;

  const { headers, rows, source } = currentParsedData;
  if (rows.length === 0) {
    alert('ពុំមានទិន្នន័យជួរដេកដើម្បីដំណើរការឡើយ!');
    return;
  }

  const modeRadio = document.querySelector('input[name="importScopeMode"]:checked');
  const mode = modeRadio ? modeRadio.value : 'full';

  const chkEvidence = document.getElementById('chkAttachAsEvidence');
  const shouldAttachEvidence = chkEvidence ? chkEvidence.checked : true;

  if (mode === 'column') {
    applyColumnImport(sub, headers, rows, shouldAttachEvidence, source);
  } else if (mode === 'row') {
    applyRowImport(sub, headers, rows, shouldAttachEvidence, source);
  } else {
    applyFullTableImport(sub, headers, rows, shouldAttachEvidence, source);
  }
}

/**
 * មុខងារ Toggle រវាងជម្រើសផ្លាស់ប្តូរតារាងទាំងស្រុង និងរក្សាក្បាលតារាងដើម
 */
function onReplaceTableCompletelyToggled() {
  const chkReplaceCompletely = document.getElementById('chkReplaceTableCompletely');
  const chkKeepOfficial = document.getElementById('chkKeepOriginalHeaders');
  const chkReplace = document.getElementById('chkReplaceExisting');
  if (chkReplaceCompletely && chkReplaceCompletely.checked) {
    if (chkKeepOfficial) chkKeepOfficial.checked = false;
    if (chkReplace) chkReplace.checked = true;
  }
}

function onKeepOriginalHeadersToggled() {
  const chkReplaceCompletely = document.getElementById('chkReplaceTableCompletely');
  const chkKeepOfficial = document.getElementById('chkKeepOriginalHeaders');
  if (chkKeepOfficial && chkKeepOfficial.checked) {
    if (chkReplaceCompletely) chkReplaceCompletely.checked = false;
  }
}

/**
 * នាំចូលទិន្នន័យជាទម្រង់ «តារាងទាំងមូល (Full Table)»
 */
function applyFullTableImport(sub, headers, rows, shouldAttachEvidence, source) {
  const chkReplace = document.getElementById('chkReplaceExisting');
  const shouldReplace = chkReplace ? chkReplace.checked : true;

  const chkReplaceCompletely = document.getElementById('chkReplaceTableCompletely');
  const shouldReplaceCompletely = chkReplaceCompletely ? chkReplaceCompletely.checked : true;

  const chkKeepOfficialHeaders = document.getElementById('chkKeepOriginalHeaders');
  const shouldKeepOfficialHeaders = chkKeepOfficialHeaders ? chkKeepOfficialHeaders.checked : false;

  // Get current document sheet base HTML
  let existingHtml = localStorage.getItem('saved_doc_' + currentImportSubCode);
  let baseHtml = existingHtml || buildDefaultDocumentHtml(sub);

  const parser = new DOMParser();
  const doc = parser.parseFromString('<div>' + baseHtml + '</div>', 'text/html');
  const container = doc.body.firstElementChild;
  const existingTable = container.querySelector('.doc-table');

  // ចំនួនជួរឈរផ្លូវការ
  let officialColsCount = 0;
  if (existingTable) {
    officialColsCount = existingTable.querySelectorAll('thead th:not(.row-action-col)').length;
  }
  if (!officialColsCount) officialColsCount = (sub.headers || []).length || headers.length;

  const isReplacingAll = shouldReplaceCompletely || (!shouldKeepOfficialHeaders && shouldReplace);
  const totalColsToRender = isReplacingAll ? headers.length : officialColsCount;
  const firstHeaderTitle = String(headers[0] || '');
  const isCol0Index = Boolean(firstHeaderTitle && /^(ល\.?រ|no|n°|លរ|#)/i.test(firstHeaderTitle.trim()));

  // Build Table Headers (សម្រាប់ករណីផ្លាស់ប្តូរក្បាលតារាង)
  let tableHeadersHtml = '';
  headers.forEach((h, idx) => {
    const title = h || `កូឡោនទី ${khmerNumber(idx + 1)}`;
    tableHeadersHtml += `<th class="editable-header" contenteditable="true">${escapeHtml(title)}</th>`;
  });
  tableHeadersHtml += `<th class="row-action-col no-print" style="width: 58px; border: none; background: transparent;"></th>`;

  // Build Table Rows
  let tableRowsHtml = '';
  rows.forEach((row, rowIndex) => {
    tableRowsHtml += '<tr>';
    for (let colIndex = 0; colIndex < totalColsToRender; colIndex++) {
      let cellValue = row[colIndex] !== undefined ? String(row[colIndex]).trim() : '';
      
      // Auto convert index if column 0 is No/ល.រ
      if (colIndex === 0 && isCol0Index) {
        if (!cellValue || !isNaN(cellValue) || /^[០-៩\d]+$/.test(cellValue)) {
          cellValue = khmerNumber(cellValue || (rowIndex + 1));
        }
      }

      const isCenter = (colIndex === 0 && isCol0Index) || isNumericData(cellValue);
      tableRowsHtml += `<td class="${isCenter ? 'text-center ' : ''}editable-cell" contenteditable="true">${escapeHtml(cellValue)}</td>`;
    }
    tableRowsHtml += `<td class="row-action-col no-print">${getRowActionButtonsHtml()}</td>`;
    tableRowsHtml += '</tr>';
  });

  if (existingTable) {
    if (isReplacingAll) {
      const theadTr = existingTable.querySelector('thead tr');
      if (theadTr) {
        theadTr.innerHTML = tableHeadersHtml;
      } else {
        let thead = existingTable.querySelector('thead');
        if (!thead) {
          thead = document.createElement('thead');
          existingTable.insertBefore(thead, existingTable.firstChild);
        }
        thead.innerHTML = `<tr>${tableHeadersHtml}</tr>`;
      }
    }

    if (shouldReplace || shouldReplaceCompletely) {
      let tbody = existingTable.querySelector('tbody');
      if (!tbody) {
        tbody = document.createElement('tbody');
        existingTable.appendChild(tbody);
      }
      tbody.innerHTML = tableRowsHtml;
    } else {
      let tbody = existingTable.querySelector('tbody');
      if (!tbody) {
        tbody = document.createElement('tbody');
        existingTable.appendChild(tbody);
      }
      tbody.insertAdjacentHTML('beforeend', tableRowsHtml);
    }

    // Dynamic wide table classes
    if (totalColsToRender >= 10) {
      existingTable.classList.add('table-ultra-wide');
      existingTable.classList.remove('table-wide');
    } else if (totalColsToRender >= 7) {
      existingTable.classList.add('table-wide');
      existingTable.classList.remove('table-ultra-wide');
    } else {
      existingTable.classList.remove('table-wide', 'table-ultra-wide');
    }
  }

  const finalHtmlToSave = container.innerHTML;
  const evCaption = source === 'sheet' ? `តារាងទិន្នន័យ Google Sheet (${sub.name})` : `ឯកសារ Excel (${sub.name})`;
  const evDesc = `ឯកសារតារាងទិន្នន័យបាននាំចូល និងផ្លាស់ប្តូរតារាងរបាយការណ៍ទាំងស្រុងចំនួន ${khmerNumber(totalColsToRender)} ជួរឈរ និង ${khmerNumber(rows.length)} ជួរដេក។`;
  const toastMsg = `⚡ បានទាញទិន្នន័យផ្លាស់ប្តូរតារាងទាំងស្រុងពី Excel/Sheet [${sub.code}] ជោគជ័យ (${khmerNumber(totalColsToRender)} ជួរឈរ, ${khmerNumber(rows.length)} ជួរដេក)!`;

  saveImportResultAndOpen(sub, finalHtmlToSave, shouldAttachEvidence, source, evCaption, evDesc, toastMsg);
}

/**
 * នាំចូលទិន្នន័យជាទម្រង់ «ម្តងមួយជួរឈរ (Single Column)»
 */
function applyColumnImport(sub, headers, rows, shouldAttachEvidence, source) {
  const srcColSelect = document.getElementById('selectSourceColumn');
  const srcColIdx = srcColSelect ? parseInt(srcColSelect.value || '0', 10) : 0;

  const targetColSelect = document.getElementById('selectTargetColumn');
  const targetVal = targetColSelect ? targetColSelect.value : '__NEW__';

  const chkFirstRowHeader = document.getElementById('chkColFirstRowIsHeader');
  const firstRowIsHeader = chkFirstRowHeader ? chkFirstRowHeader.checked : true;

  const rawHeaderName = headers[srcColIdx] || `កូឡោនថ្មី`;
  let dataValues = rows.map(r => r[srcColIdx] !== undefined ? String(r[srcColIdx]).trim() : '');

  let finalHeaderName = rawHeaderName;
  if (!firstRowIsHeader) {
    // ជួរដេកទី១ ជាទិន្នន័យ យកមកបញ្ចូលជាទិន្នន័យជួរទី១
    dataValues = [rawHeaderName, ...dataValues];
    finalHeaderName = `កូឡោនថ្មី`;
  }

  // Get current document HTML
  let existingHtml = localStorage.getItem('saved_doc_' + currentImportSubCode);
  let baseHtml = existingHtml || buildDefaultDocumentHtml(sub);

  const parser = new DOMParser();
  const doc = parser.parseFromString('<div>' + baseHtml + '</div>', 'text/html');
  const container = doc.body.firstElementChild;
  const table = container.querySelector('.doc-table');

  if (!table) {
    alert('មិនអាចរកឃើញតារាងក្នុងឯកសារនេះឡើយ!');
    return;
  }

  const theadRow = table.querySelector('thead tr');
  const tbody = table.querySelector('tbody');
  const actionTh = theadRow.querySelector('.row-action-col');
  const dataHeaders = Array.from(theadRow.querySelectorAll('th:not(.row-action-col)'));
  let tbodyRows = Array.from(tbody.querySelectorAll('tr'));

  let targetColName = finalHeaderName;

  if (targetVal === '__NEW__') {
    // 1. បង្កើតជាជួរឈរថ្មី (Append New Column)
    const newTh = doc.createElement('th');
    newTh.className = 'editable-header';
    newTh.setAttribute('contenteditable', 'true');
    newTh.textContent = finalHeaderName;
    if (actionTh) {
      theadRow.insertBefore(newTh, actionTh);
    } else {
      theadRow.appendChild(newTh);
    }

    const newColCount = dataHeaders.length + 1;

    // ប្រសិនបើទិន្នន័យដែលនាំចូលមានចំនួនច្រើនជាងជួរដេកដែលមានស្រាប់ សូមពង្រីកជួរដេកស្វ័យប្រវត្តិ
    while (tbodyRows.length < dataValues.length) {
      const newTr = doc.createElement('tr');
      for (let c = 0; c < newColCount; c++) {
        const td = doc.createElement('td');
        const isCenter = (c === 0);
        td.className = (isCenter ? 'text-center ' : '') + 'editable-cell';
        td.setAttribute('contenteditable', 'true');
        if (c === 0) {
          td.textContent = khmerNumber(tbodyRows.length + 1);
        } else {
          td.textContent = '';
        }
        newTr.appendChild(td);
      }
      const actionTd = doc.createElement('td');
      actionTd.className = 'row-action-col no-print';
      actionTd.innerHTML = getRowActionButtonsHtml();
      newTr.appendChild(actionTd);
      tbody.appendChild(newTr);
      tbodyRows.push(newTr);
    }

    // បញ្ចូលក្រឡាថ្មីទៅក្នុងគ្រប់ជួរដេកទាំងអស់នៃតារាង
    tbodyRows.forEach((row, rIdx) => {
      const val = dataValues[rIdx] !== undefined ? dataValues[rIdx] : '';
      const actionTd = row.querySelector('.row-action-col');
      const dataCells = Array.from(row.querySelectorAll('td:not(.row-action-col)'));

      if (dataCells.length < newColCount) {
        const td = doc.createElement('td');
        const isCenter = isNumericData(val);
        td.className = (isCenter ? 'text-center ' : '') + 'editable-cell';
        td.setAttribute('contenteditable', 'true');
        td.textContent = val;
        if (actionTd) {
          row.insertBefore(td, actionTd);
        } else {
          row.appendChild(td);
        }
      } else {
        const targetTd = dataCells[newColCount - 1];
        targetTd.textContent = val;
        if (isNumericData(val)) targetTd.classList.add('text-center');
        else targetTd.classList.remove('text-center');
      }
    });

  } else {
    // 2. ជំនួសចូលក្នុងជួរឈរណាមួយដែលមានស្រាប់
    const chkKeepHeader = document.getElementById('chkKeepTargetColHeader');
    const shouldKeepHeader = chkKeepHeader ? chkKeepHeader.checked : true;

    const targetColIdx = parseInt(targetVal, 10);
    if (dataHeaders[targetColIdx]) {
      const origHeaderName = dataHeaders[targetColIdx].innerText.replace(/[\r\n\t]+/g, ' ').trim();
      targetColName = origHeaderName;
      if (!shouldKeepHeader && firstRowIsHeader) {
        dataHeaders[targetColIdx].textContent = finalHeaderName;
        targetColName = finalHeaderName;
      }
    }

    const curColCount = dataHeaders.length;

    // ពង្រីកជួរដេកបើទិន្នន័យលើសពីចំនួនជួរដេកបច្ចុប្បន្ន
    while (tbodyRows.length < dataValues.length) {
      const newTr = doc.createElement('tr');
      for (let c = 0; c < curColCount; c++) {
        const td = doc.createElement('td');
        const isCenter = (c === 0);
        td.className = (isCenter ? 'text-center ' : '') + 'editable-cell';
        td.setAttribute('contenteditable', 'true');
        if (c === 0) {
          td.textContent = khmerNumber(tbodyRows.length + 1);
        } else {
          td.textContent = '';
        }
        newTr.appendChild(td);
      }
      const actionTd = doc.createElement('td');
      actionTd.className = 'row-action-col no-print';
      actionTd.innerHTML = getRowActionButtonsHtml();
      newTr.appendChild(actionTd);
      tbody.appendChild(newTr);
      tbodyRows.push(newTr);
    }

    // ជំនួសតម្លៃក្រឡាក្នុងជួរឈរដែលបានជ្រើស
    tbodyRows.forEach((row, rIdx) => {
      const dataCells = Array.from(row.querySelectorAll('td:not(.row-action-col)'));
      if (dataCells[targetColIdx]) {
        const val = dataValues[rIdx] !== undefined ? dataValues[rIdx] : '';
        dataCells[targetColIdx].textContent = val;
        if (targetColIdx === 0 || isNumericData(val)) {
          dataCells[targetColIdx].classList.add('text-center');
        } else {
          dataCells[targetColIdx].classList.remove('text-center');
        }
      }
    });
  }

  const finalHtmlToSave = container.innerHTML;
  const evCaption = `ជួរឈរទិន្នន័យ Excel/Sheet (${targetColName})`;
  const evDesc = `ឯកសារបានចម្លងទិន្នន័យចូលក្នុងជួរឈរ «${targetColName}» ចំនួន ${khmerNumber(dataValues.length)} ជួរដេក។`;
  const toastMsg = `⚡ បានចម្លងទិន្នន័យចូលក្នុងជួរឈរ «${targetColName}» នៃរបាយការណ៍ [${sub.code}] ជោគជ័យ (${khmerNumber(dataValues.length)} ជួរ)!`;

  saveImportResultAndOpen(sub, finalHtmlToSave, shouldAttachEvidence, source, evCaption, evDesc, toastMsg);
}

/**
 * នាំចូលទិន្នន័យជាទម្រង់ «ម្តងមួយជួរដេក (Single Row)»
 */
function applyRowImport(sub, headers, rows, shouldAttachEvidence, source) {
  const srcRowSelect = document.getElementById('selectSourceRow');
  const srcRowVal = srcRowSelect ? srcRowSelect.value : 'all';

  const targetActionSelect = document.getElementById('selectTargetRowAction');
  const targetAction = targetActionSelect ? targetActionSelect.value : 'append';

  const chkRowFirstRowIsHeader = document.getElementById('chkRowFirstRowIsHeader');
  const skipHeaderRow = chkRowFirstRowIsHeader ? chkRowFirstRowIsHeader.checked : (rows.length > 1);

  // កំណត់ជួរដេកដែលត្រូវនាំចូល
  let rowsToImport = [];
  if (srcRowVal === 'all') {
    rowsToImport = rows.slice();
    if (!skipHeaderRow && headers && headers.length > 0) {
      rowsToImport = [headers, ...rowsToImport];
    }
  } else {
    const rowIdx = parseInt(srcRowVal, 10);
    if (rows[rowIdx]) {
      rowsToImport = [rows[rowIdx]];
    }
  }

  if (rowsToImport.length === 0) {
    alert('ពុំមានទិន្នន័យជួរដេកដើម្បីចម្លងចូលឡើយ!');
    return;
  }

  // Get current document HTML
  let existingHtml = localStorage.getItem('saved_doc_' + currentImportSubCode);
  let baseHtml = existingHtml || buildDefaultDocumentHtml(sub);

  const parser = new DOMParser();
  const doc = parser.parseFromString('<div>' + baseHtml + '</div>', 'text/html');
  const container = doc.body.firstElementChild;
  const table = container.querySelector('.doc-table');

  if (!table) {
    alert('មិនអាចរកឃើញតារាងក្នុងឯកសារនេះឡើយ!');
    return;
  }

  const theadRow = table.querySelector('thead tr');
  const tbody = table.querySelector('tbody');
  const dataHeaders = Array.from(theadRow.querySelectorAll('th:not(.row-action-col)'));
  const colCount = dataHeaders.length;
  let tbodyRows = Array.from(tbody.querySelectorAll('tr'));

  if (targetAction === 'replace_specific') {
    // ជំនួសជួរដេកណាមួយជាក់លាក់
    const targetRowIndexSelect = document.getElementById('selectTargetRowIndex');
    const targetRowIdx = targetRowIndexSelect ? parseInt(targetRowIndexSelect.value || '0', 10) : 0;
    const targetTr = tbodyRows[targetRowIdx];

    if (!targetTr) {
      alert('មិនអាចរកឃើញជួរដេកគោលដៅឡើយ!');
      return;
    }

    const sourceDataRow = rowsToImport[0];
    const dataCells = Array.from(targetTr.querySelectorAll('td:not(.row-action-col)'));

    dataCells.forEach((cell, cIdx) => {
      let val = sourceDataRow[cIdx] !== undefined ? String(sourceDataRow[cIdx]).trim() : '';
      if (cIdx === 0 && (!val || !isNaN(val))) {
        val = khmerNumber(targetRowIdx + 1);
      }
      cell.textContent = val;
      if (cIdx === 0 || isNumericData(val)) {
        cell.classList.add('text-center');
      } else {
        cell.classList.remove('text-center');
      }
    });

  } else {
    // បន្ថែមជាជួរដេកថ្មីនៅខាងក្រោមបង្អស់
    rowsToImport.forEach(importedRow => {
      const curTotalRows = tbody.querySelectorAll('tr').length;
      const newTr = doc.createElement('tr');

      for (let c = 0; c < colCount; c++) {
        let val = importedRow[c] !== undefined ? String(importedRow[c]).trim() : '';
        if (c === 0 && (!val || !isNaN(val))) {
          val = khmerNumber(curTotalRows + 1);
        }

        const td = doc.createElement('td');
        const isCenter = (c === 0) || isNumericData(val);
        td.className = (isCenter ? 'text-center ' : '') + 'editable-cell';
        td.setAttribute('contenteditable', 'true');
        td.textContent = val;
        newTr.appendChild(td);
      }

      const actionTd = doc.createElement('td');
      actionTd.className = 'row-action-col no-print';
      actionTd.innerHTML = getRowActionButtonsHtml();
      newTr.appendChild(actionTd);

      tbody.appendChild(newTr);
    });
  }

  const finalHtmlToSave = container.innerHTML;
  const evCaption = `ជួរដេកទិន្នន័យ Excel/Sheet (${sub.name})`;
  const evDesc = `ឯកសារបានចម្លងទិន្នន័យចូលក្នុងជួរដេកចំនួន ${khmerNumber(rowsToImport.length)} ជួរ។`;
  const toastMsg = `⚡ បានចម្លងទិន្នន័យ ${khmerNumber(rowsToImport.length)} ជួរដេក ចូលក្នុងរបាយការណ៍ [${sub.code}] ជោគជ័យ!`;

  saveImportResultAndOpen(sub, finalHtmlToSave, shouldAttachEvidence, source, evCaption, evDesc, toastMsg);
}

/**
 * រក្សាទុកលទ្ធផលនៃការនាំចូលទិន្នន័យ ភ្ជាប់ភស្តុតាង និងបើកឯកសារឡើងវិញ
 */
function saveImportResultAndOpen(sub, finalHtmlToSave, shouldAttachEvidence, source, evCaption, evDesc, toastMsg) {
  // រក្សាទុកក្នុង LocalStorage
  localStorage.setItem('saved_doc_' + sub.code, finalHtmlToSave);
  const timeFormatted = formatKhmerDateTime(new Date());
  localStorage.setItem('saved_doc_time_' + sub.code, timeFormatted);

  // ប្រសិនបើបានធីក៖ រក្សាទុកប្រភពជាភស្តុតាងបញ្ជាក់ក្នុងឧបសម្ព័ន្ធស្វ័យប្រវត្តិ
  if (shouldAttachEvidence) {
    let evList = getEvidenceList(sub.code);
    const dateStr = getKhmerCurrentDate();

    if (source === 'sheet' && currentGoogleSheetUrl) {
      evList.push({
        id: 'ev_' + Date.now(),
        mode: 'link',
        imageData: '',
        fileName: '',
        linkUrl: currentGoogleSheetUrl,
        caption: evCaption || `តារាងទិន្នន័យ Google Sheet (${escapeHtml(sub.name)})`,
        date: dateStr,
        description: evDesc || `ឯកសារតារាងទិន្នន័យ Google Sheet បាននាំចូលបង្កើតរបាយការណ៍។`
      });
      localStorage.setItem('evidence_' + sub.code, JSON.stringify(evList));
      updateCardEvidenceBadge(sub.code);
    } else if (source === 'file' && currentUploadedFile) {
      evList.push({
        id: 'ev_' + Date.now(),
        mode: 'file',
        imageData: '',
        fileName: currentUploadedFile.name,
        linkUrl: '',
        caption: evCaption || `ឯកសារ Excel៖ ${escapeHtml(currentUploadedFile.name)}`,
        date: dateStr,
        description: evDesc || `ហ្វាល់ Excel បាននាំចូល និងទាញបង្កើតជារបាយការណ៍។`
      });
      localStorage.setItem('evidence_' + sub.code, JSON.stringify(evList));
      updateCardEvidenceBadge(sub.code);
    }
  }

  // ធ្វើបច្ចុប្បន្នកម្ម Badge លើ Card មេ
  updateCardBadge(sub.code, true);

  // បិទផ្ទាំង Import Modal
  closeExcelImportModal();

  // បើកផ្ទាំង Document Modal ដើម្បីពិនិត្យ កែប្រែផ្ទាល់ បោះពុម្ព ឬទាញយក
  openDocumentModal(sub.code);

  showToast(toastMsg);
}

// Attach ESC listener for Excel Import Modal
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    const importModal = document.getElementById('excelImportModal');
    if (importModal && importModal.classList.contains('active')) {
      closeExcelImportModal();
    }
  }
});

// Close Excel Import Modal when clicking outside content box
document.addEventListener('DOMContentLoaded', () => {
  const importModal = document.getElementById('excelImportModal');
  if (importModal) {
    importModal.addEventListener('click', (e) => {
      if (e.target.id === 'excelImportModal') closeExcelImportModal();
    });
  }
});

// Expose Excel Import & Scope functions globally to window
window.openExcelImportModal = openExcelImportModal;
window.closeExcelImportModal = closeExcelImportModal;
window.switchImportTab = switchImportTab;
window.handleExcelFileInput = handleExcelFileInput;
window.onExcelSheetSelected = onExcelSheetSelected;
window.fetchGoogleSheetData = fetchGoogleSheetData;
window.processPastedData = processPastedData;
window.onPastedDataChanged = onPastedDataChanged;
window.setPasteModePlaceholder = setPasteModePlaceholder;
window.onImportScopeModeChanged = onImportScopeModeChanged;
window.onSourceColumnChanged = onSourceColumnChanged;
window.onTargetColumnChanged = onTargetColumnChanged;
window.onSourceRowChanged = onSourceRowChanged;
window.onTargetRowActionChanged = onTargetRowActionChanged;
window.onColFirstRowHeaderToggled = onColFirstRowHeaderToggled;
window.onRowFirstRowHeaderToggled = onRowFirstRowHeaderToggled;
window.applyImportedDataToDocument = applyImportedDataToDocument;
window.handleTableClipboardPaste = handleTableClipboardPaste;
window.pasteClipboardToTable = pasteClipboardToTable;
window.pasteFromClipboardToTextarea = pasteFromClipboardToTextarea;
window.isNumericData = isNumericData;
window.applyPastedMatrixToTable = applyPastedMatrixToTable;
window.parseHtmlTableToMatrix = parseHtmlTableToMatrix;
window.replaceDocumentTableWithData = replaceDocumentTableWithData;
window.applyPastedDataSmart = applyPastedDataSmart;
window.onReplaceTableCompletelyToggled = onReplaceTableCompletelyToggled;
window.onKeepOriginalHeadersToggled = onKeepOriginalHeadersToggled;

// Expose Table Column manipulation & Interactive pickers globally to window
window.addNewTableColumn = addNewTableColumn;
window.openDeleteColumnModal = openDeleteColumnModal;
window.closeDeleteColumnModal = closeDeleteColumnModal;
window.executeDeleteColumn = executeDeleteColumn;
window.deleteLastTableColumn = deleteLastTableColumn;
window.insertTableRowAfter = insertTableRowAfter;
window.deleteCurrentTableRow = deleteCurrentTableRow;
window.reindexTableRows = reindexTableRows;
window.ensureRowActionButtons = ensureRowActionButtons;
window.insertTableColumnAfter = insertTableColumnAfter;
window.deleteTableColumnAt = deleteTableColumnAt;
window.ensureColumnActionButtonsOnHeaders = ensureColumnActionButtonsOnHeaders;
window.getColumnHeaderTitle = getColumnHeaderTitle;
window.handleAcademicYearChange = handleAcademicYearChange;
window.openNativeDatePicker = openNativeDatePicker;
window.handleDatePicked = handleDatePicked;
window.syncTeacherSignatureDate = syncTeacherSignatureDate;
window.syncAcademicYearSelect = syncAcademicYearSelect;

// Expose Evidence Date picker and formatting functions globally to window
window.formatKhmerMonthNameToWord = formatKhmerMonthNameToWord;
window.getKhmerCurrentDate = getKhmerCurrentDate;
window.triggerSlotDatePicker = triggerSlotDatePicker;
window.onSlotNativeDatePicked = onSlotNativeDatePicked;
window.onSlotDateBlurred = onSlotDateBlurred;
window.onSlotDateInput = onSlotDateInput;

// Expose Table Cell Merge & Selection functions globally to window
window.getTableGridMatrix = getTableGridMatrix;
window.clearCellSelection = clearCellSelection;
window.selectCellsBetween = selectCellsBetween;
window.mergeSelectedTableCells = mergeSelectedTableCells;
window.mergeActiveCellRight = mergeActiveCellRight;
window.mergeActiveCellDown = mergeActiveCellDown;
window.unmergeSelectedTableCells = unmergeSelectedTableCells;
window.toggleMergeDropdown = toggleMergeDropdown;
window.closeMergeDropdown = closeMergeDropdown;
window.openTableContextMenu = openTableContextMenu;
window.closeTableContextMenu = closeTableContextMenu;
window.handleContextMenuAction = handleContextMenuAction;

// Expose Multi-page sheet functions globally to window
window.addNewDocumentSheetPage = addNewDocumentSheetPage;
window.deleteDocumentSheetPage = deleteDocumentSheetPage;
window.refreshDocumentPageNumbers = refreshDocumentPageNumbers;


