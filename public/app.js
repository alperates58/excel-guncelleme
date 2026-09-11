// ==============================================================================
// Excel SQL Connection & IP Updater - Frontend Application Logic
// ==============================================================================

const API_BASE = window.location.origin + '/api';

let currentScanData = null;
let progressInterval = null;
let serverWorkspace = '';
let userDesktop = '';
let userDownloads = '';
let currentBrowserDir = '';
let parentBrowserDir = '';

// DOM Elements
const serverStatus = document.getElementById('serverStatus');
const folderPathInput = document.getElementById('folderPath');
const btnBrowse = document.getElementById('btnBrowse');
const btnScan = document.getElementById('btnScan');
const btnOpenFolder = document.getElementById('btnOpenFolder');
const btnBackup = document.getElementById('btnBackup');
const btnExecuteUpdate = document.getElementById('btnExecuteUpdate');
const btnAddRule = document.getElementById('btnAddRule');
const rulesList = document.getElementById('rulesList');
const quickChips = document.getElementById('quickChips');

// Options
const chkQueries = document.getElementById('chkQueries');
const chkConnections = document.getElementById('chkConnections');
const chkVba = document.getElementById('chkVba');
const chkAutoBackup = document.getElementById('chkAutoBackup');
const chkAtomicBatch = document.getElementById('chkAtomicBatch');

// Stats Counters
const statTotalFiles = document.getElementById('statTotalFiles');
const statPowerQueries = document.getElementById('statPowerQueries');
const statConnections = document.getElementById('statConnections');
const statVbaModules = document.getElementById('statVbaModules');

// Progress & Table & Log
const progressSection = document.getElementById('progressSection');
const progressStatusText = document.getElementById('progressStatusText');
const progressPercent = document.getElementById('progressPercent');
const progressBar = document.getElementById('progressBar');

const filesTable = document.getElementById('filesTable').querySelector('tbody');
const tableSearch = document.getElementById('tableSearch');
const badgeFileCount = document.getElementById('badgeFileCount');

const logConsole = document.getElementById('logConsole');
const btnClearLog = document.getElementById('btnClearLog');

// Modals
const detailModal = document.getElementById('detailModal');
const modalFileName = document.getElementById('modalFileName');
const modalBody = document.getElementById('modalBody');
const btnModalClose = document.getElementById('btnModalClose');

const folderBrowserModal = document.getElementById('folderBrowserModal');
const btnFolderModalClose = document.getElementById('btnFolderModalClose');
const modalFolderPathInput = document.getElementById('modalFolderPathInput');
const subfoldersList = document.getElementById('subfoldersList');
const btnPathUp = document.getElementById('btnPathUp');
const btnSelectCurrentFolder = document.getElementById('btnSelectCurrentFolder');

// Shortcuts
const btnShortcutDesktop = document.getElementById('btnShortcutDesktop');
const btnShortcutDownloads = document.getElementById('btnShortcutDownloads');
const btnShortcutProject = document.getElementById('btnShortcutProject');
const btnShortcutC = document.getElementById('btnShortcutC');
const btnShortcutD = document.getElementById('btnShortcutD');

// Initialize App
document.addEventListener('DOMContentLoaded', () => {
    checkServerStatus();
    bindEvents();
});

function bindEvents() {
    btnBrowse.addEventListener('click', openFolderBrowserModal);
    btnFolderModalClose.addEventListener('click', () => { folderBrowserModal.style.display = 'none'; });

    btnPathUp.addEventListener('click', () => {
        if (parentBrowserDir) loadFolderTree(parentBrowserDir);
    });

    btnSelectCurrentFolder.addEventListener('click', selectFolderFromModal);

    btnShortcutDesktop.addEventListener('click', () => { if (userDesktop) loadFolderTree(userDesktop); });
    btnShortcutDownloads.addEventListener('click', () => { if (userDownloads) loadFolderTree(userDownloads); });
    btnShortcutProject.addEventListener('click', () => { if (serverWorkspace) loadFolderTree(serverWorkspace); });
    btnShortcutC.addEventListener('click', () => loadFolderTree('C:\\'));
    btnShortcutD.addEventListener('click', () => loadFolderTree('D:\\'));

    btnScan.addEventListener('click', runScan);
    btnOpenFolder.addEventListener('click', openFolder);
    btnBackup.addEventListener('click', runBackup);
    btnExecuteUpdate.addEventListener('click', runUpdate);
    btnAddRule.addEventListener('click', () => addRuleRow());
    btnClearLog.addEventListener('click', () => { logConsole.innerHTML = ''; });

    tableSearch.addEventListener('input', filterTable);

    btnModalClose.addEventListener('click', () => { detailModal.style.display = 'none'; });
    window.addEventListener('click', (e) => {
        if (e.target === detailModal) detailModal.style.display = 'none';
        if (e.target === folderBrowserModal) folderBrowserModal.style.display = 'none';
    });
}

function appendLog(message, type = 'info') {
    const time = new Date().toLocaleTimeString();
    const entry = document.createElement('div');
    entry.className = `log-entry log-${type}`;
    entry.textContent = `[${time}] ${message}`;
    logConsole.appendChild(entry);
    logConsole.scrollTop = logConsole.scrollHeight;
}

// ------------------------------------------------------------------------------
// LIVE PROGRESS POLLING
// ------------------------------------------------------------------------------
function startProgressPolling() {
    progressSection.classList.remove('hidden');
    if (progressInterval) clearInterval(progressInterval);

    progressInterval = setInterval(async () => {
        try {
            const res = await fetch(`${API_BASE}/progress`);
            const data = await res.json();
            if (data && data.active) {
                const label = data.type === 'scan' ? 'Taranıyor' : 'Güncelleniyor';
                progressStatusText.textContent = `${label} (%${data.percent}): ${data.currentFile} (${data.current} / ${data.total})`;
                progressBar.style.width = `${data.percent}%`;
                progressPercent.textContent = `${data.percent}%`;
            }
        } catch (e) { }
    }, 400);
}

function stopProgressPolling(finalStatus = 'İşlem Tamamlandı') {
    if (progressInterval) {
        clearInterval(progressInterval);
        progressInterval = null;
    }
    progressBar.style.width = '100%';
    progressPercent.textContent = '100%';
    progressStatusText.textContent = finalStatus;
}

// ------------------------------------------------------------------------------
// SERVER STATUS CHECK
// ------------------------------------------------------------------------------
async function checkServerStatus() {
    try {
        const res = await fetch(`${API_BASE}/status`);
        const data = await res.json();
        if (data.status === 'ok') {
            serverWorkspace = data.workspace || '';
            userDesktop = data.userDesktop || '';
            userDownloads = data.userDownloads || '';

            serverStatus.innerHTML = `
                <span class="status-dot online"></span>
                <span class="status-text">Sunucu Aktif (Port ${data.port})</span>
            `;
            appendLog(`Sunucuya bağlandı. Çalışma dizini: ${data.workspace}`, 'success');
        }
    } catch (err) {
        serverStatus.innerHTML = `
            <span class="status-dot pulsing"></span>
            <span class="status-text">Sunucuya Bağlanılamadı</span>
        `;
        appendLog(`Sunucu bağlantı hatası: ${err.message}`, 'error');
    }
}

// ------------------------------------------------------------------------------
// INTERACTIVE WEB FOLDER BROWSER MODAL
// ------------------------------------------------------------------------------
function openFolderBrowserModal() {
    const currentPath = folderPathInput.value.trim() || serverWorkspace;
    loadFolderTree(currentPath);
    folderBrowserModal.style.display = 'flex';
}

async function loadFolderTree(targetDir) {
    subfoldersList.innerHTML = `<div class="subfolder-card-placeholder"><i class="fa-solid fa-spinner fa-spin"></i> Klasörler yükleniyor...</div>`;

    try {
        const res = await fetch(`${API_BASE}/list-folders`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ directory: targetDir })
        });
        const data = await res.json();

        if (data.success) {
            currentBrowserDir = data.currentDir;
            parentBrowserDir = data.parentDir;
            modalFolderPathInput.value = data.currentDir;

            btnPathUp.disabled = !data.parentDir;

            renderSubfolders(data.subFolders, data.drives);
        } else {
            subfoldersList.innerHTML = `<p class="text-rose">Klasör okunamadı: ${data.error}</p>`;
        }
    } catch (err) {
        subfoldersList.innerHTML = `<p class="text-rose">Bağlantı hatası: ${err.message}</p>`;
    }
}

function renderSubfolders(subFolders, drives) {
    subfoldersList.innerHTML = '';

    const list = subFolders && subFolders.length > 0 ? subFolders : [];

    if (drives && drives.length > 0) {
        drives.forEach(d => {
            const card = document.createElement('div');
            card.className = 'subfolder-card';
            card.innerHTML = `<i class="fa-solid fa-hard-drive"></i> <strong>${escapeHtml(d.name)}</strong>`;
            card.addEventListener('click', () => loadFolderTree(d.path));
            subfoldersList.appendChild(card);
        });
    }

    if (list.length === 0 && (!drives || drives.length === 0)) {
        subfoldersList.innerHTML = `<p class="text-dim">Bu klasörün içinde alt klasör bulunmuyor.</p>`;
        return;
    }

    list.forEach(sf => {
        const card = document.createElement('div');
        card.className = 'subfolder-card';
        card.title = sf.path;
        card.innerHTML = `<i class="fa-solid fa-folder"></i> <span>${escapeHtml(sf.name)}</span>`;
        card.addEventListener('click', () => loadFolderTree(sf.path));
        subfoldersList.appendChild(card);
    });
}

function selectFolderFromModal() {
    const selected = modalFolderPathInput.value.trim();
    if (selected) {
        folderPathInput.value = selected;
        folderPathInput.title = selected;
        appendLog(`Yeni çalışma klasörü seçildi: "${selected}"`, 'success');
        folderBrowserModal.style.display = 'none';
        runScan();
    }
}

// ------------------------------------------------------------------------------
// OPEN WORKING FOLDER IN WINDOWS EXPLORER
// ------------------------------------------------------------------------------
async function openFolder() {
    const dir = folderPathInput.value.trim();
    if (!dir) return;

    try {
        const res = await fetch(`${API_BASE}/open-folder`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ directory: dir })
        });
        const data = await res.json();

        if (data.success) {
            appendLog(`Klasör Gezginde açıldı: "${dir}"`, 'success');
        } else {
            appendLog(`Klasör açılamadı: ${data.error}`, 'error');
            alert(`Klasör açılamadı: ${data.error}`);
        }
    } catch (err) {
        appendLog(`Klasör açma hatası: ${err.message}`, 'error');
    }
}

// ------------------------------------------------------------------------------
// SCAN FILES
// ------------------------------------------------------------------------------
async function runScan() {
    const dir = folderPathInput.value.trim();
    if (!dir) {
        alert('Lütfen geçerli bir klasör yolu girin.');
        return;
    }

    btnScan.disabled = true;
    btnScan.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Taranıyor...`;
    appendLog(`Excel dosyaları taranıyor: "${dir}"...`, 'info');
    startProgressPolling();

    try {
        const res = await fetch(`${API_BASE}/scan`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ directory: dir })
        });
        const data = await res.json();

        if (data.success) {
            currentScanData = data;
            renderScanResults(data);
            stopProgressPolling(`Tarama Tamamlandı! (${data.totalFiles} Dosya)`);
            appendLog(`Tarama tamamlandı! Toplam ${data.totalFiles} dosya incelendi.`, 'success');
        } else {
            stopProgressPolling(`Tarama Hatayla Sonlandı!`);
            appendLog(`Tarama hatası: ${data.error}`, 'error');
        }
    } catch (err) {
        stopProgressPolling(`Bağlantı Hatası!`);
        appendLog(`Bağlantı hatası: ${err.message}`, 'error');
    } finally {
        btnScan.disabled = false;
        btnScan.innerHTML = `<i class="fa-solid fa-magnifying-glass-chart"></i> Klasörü Tara`;
    }
}

function renderScanResults(data) {
    // Render Quick Chips
    quickChips.innerHTML = '';
    if (data.detectedIPs && data.detectedIPs.length > 0) {
        data.detectedIPs.forEach(item => {
            const chip = document.createElement('span');
            chip.className = 'chip';
            chip.innerHTML = `<i class="fa-solid fa-network-wired"></i> ${item.ip} (${item.count})`;
            chip.addEventListener('click', () => {
                setRuleOldIP(item.ip);
            });
            quickChips.appendChild(chip);
        });
    } else {
        quickChips.innerHTML = `<span class="chip-placeholder">IP adresi tespit edilmedi.</span>`;
    }

    // Render Stats
    statTotalFiles.textContent = data.totalFiles || 0;
    
    let totalQueries = 0;
    let totalConnections = 0;
    let totalVba = 0;

    data.files.forEach(f => {
        totalQueries += (f.queries ? f.queries.length : 0);
        totalConnections += (f.connections ? f.connections.length : 0);
        totalVba += (f.vbaMatches ? f.vbaMatches.length : 0);
    });

    statPowerQueries.textContent = totalQueries;
    statConnections.textContent = totalConnections;
    statVbaModules.textContent = totalVba;

    // Render Table
    renderTable(data.files);
}

function renderTable(files) {
    filesTable.innerHTML = '';

    if (!files || files.length === 0) {
        filesTable.innerHTML = `
            <tr class="empty-row">
                <td colspan="8">
                    <div class="empty-state">
                        <i class="fa-solid fa-folder-open empty-icon"></i>
                        <p>Belirtilen klasörde Excel dosyası (.xlsx, .xlsm) bulunamadı.</p>
                    </div>
                </td>
            </tr>
        `;
        badgeFileCount.textContent = '0 Dosya';
        return;
    }

    badgeFileCount.textContent = `${files.length} Dosya Listelendi`;

    files.forEach((f, idx) => {
        const tr = document.createElement('tr');
        
        const qCount = f.queries ? f.queries.length : 0;
        const cCount = f.connections ? f.connections.length : 0;
        const vCount = f.vbaMatches ? f.vbaMatches.length : 0;
        const ipsStr = f.foundIPs && f.foundIPs.length > 0 ? f.foundIPs.join(', ') : '-';

        let badgeClass = 'badge-info';
        if (f.status.startsWith('Error')) badgeClass = 'badge-danger';
        else if (f.status === 'Updated') badgeClass = 'badge-success';

        tr.innerHTML = `
            <td><strong>${escapeHtml(f.fileName)}</strong></td>
            <td><span class="badge ${f.extension === '.xlsm' ? 'badge-warning' : 'badge-info'}">${f.extension.toUpperCase()}</span></td>
            <td>${qCount > 0 ? `<span class="badge badge-info">${qCount} Sorgu</span>` : '<span class="text-dim">0</span>'}</td>
            <td>${cCount > 0 ? `<span class="badge badge-warning">${cCount} Bağlantı</span>` : '<span class="text-dim">0</span>'}</td>
            <td>${vCount > 0 ? `<span class="badge badge-success">${vCount} Makro</span>` : '<span class="text-dim">0</span>'}</td>
            <td><code class="text-primary">${escapeHtml(ipsStr)}</code></td>
            <td><span class="badge ${badgeClass}">${escapeHtml(f.status)}</span></td>
            <td>
                <button class="btn btn-outline-sm btn-detail" data-index="${idx}">
                    <i class="fa-solid fa-eye"></i> İncele
                </button>
            </td>
        `;

        tr.querySelector('.btn-detail').addEventListener('click', () => showFileModal(f));
        filesTable.appendChild(tr);
    });
}

function filterTable() {
    const q = tableSearch.value.toLowerCase();
    if (!currentScanData || !currentScanData.files) return;

    const filtered = currentScanData.files.filter(f => {
        return f.fileName.toLowerCase().includes(q) || 
               (f.foundIPs && f.foundIPs.some(ip => ip.includes(q)));
    });

    renderTable(filtered);
}

// ------------------------------------------------------------------------------
// MODAL DETAILED PREVIEW
// ------------------------------------------------------------------------------
function showFileModal(file) {
    modalFileName.innerHTML = `<i class="fa-solid fa-file-excel"></i> ${escapeHtml(file.fileName)}`;

    let html = `
        <div class="file-modal-details">
            <p><strong>Tam Yol:</strong> <code>${escapeHtml(file.filePath)}</code></p>
            <p><strong>Boyut:</strong> ${(file.sizeBytes / 1024 / 1024).toFixed(2)} MB</p>
            <hr class="divider">
    `;

    // Power Queries Section
    html += `<h4><i class="fa-solid fa-database"></i> Power Query Formülleri (${file.queries ? file.queries.length : 0})</h4>`;
    if (file.queries && file.queries.length > 0) {
        html += `<ul class="modal-list">`;
        file.queries.forEach(q => {
            html += `
                <li>
                    <strong>${escapeHtml(q.name)}:</strong>
                    <pre class="code-block">${escapeHtml(q.fullFormula)}</pre>
                </li>
            `;
        });
        html += `</ul>`;
    } else {
        html += `<p class="text-dim">Power Query bulunamadı.</p>`;
    }

    // Connections Section
    html += `<h4 class="mt-3"><i class="fa-solid fa-plug"></i> Veri Bağlantıları (${file.connections ? file.connections.length : 0})</h4>`;
    if (file.connections && file.connections.length > 0) {
        html += `<ul class="modal-list">`;
        file.connections.forEach(c => {
            html += `
                <li>
                    <strong>Bağlantı Adı:</strong> ${escapeHtml(c.name)}<br>
                    <strong>Bağlantı Cümlesi:</strong> <code>${escapeHtml(c.connectionString || 'Yok')}</code>
                </li>
            `;
        });
        html += `</ul>`;
    } else {
        html += `<p class="text-dim">Veri bağlantı cümlesinde IP bulunamadı.</p>`;
    }

    // VBA Section
    html += `<h4 class="mt-3"><i class="fa-solid fa-code"></i> VBA Makro Eşleşmeleri (${file.vbaMatches ? file.vbaMatches.length : 0})</h4>`;
    if (file.vbaMatches && file.vbaMatches.length > 0) {
        html += `<ul class="modal-list">`;
        file.vbaMatches.forEach(v => {
            html += `<li><strong>Modül:</strong> ${escapeHtml(v.module)} | <strong>Tespit Edilen IP:</strong> <code>${escapeHtml(v.ip)}</code></li>`;
        });
        html += `</ul>`;
    } else {
        html += `<p class="text-dim">VBA makrosunda eşleşen IP bulunamadı.</p>`;
    }

    html += `</div>`;
    modalBody.innerHTML = html;
    detailModal.style.display = 'flex';
}

// ------------------------------------------------------------------------------
// RULES MANAGEMENT
// ------------------------------------------------------------------------------
function addRuleRow(oldVal = '', newVal = '') {
    const div = document.createElement('div');
    div.className = 'rule-item';
    div.innerHTML = `
        <div class="rule-inputs">
            <div class="input-field">
                <small>Eski IP / Metin</small>
                <input type="text" class="rule-old" placeholder="Örn: 192.168.2.15" value="${escapeHtml(oldVal)}">
            </div>
            <div class="rule-arrow"><i class="fa-solid fa-arrow-right"></i></div>
            <div class="input-field">
                <small>Yeni IP / Metin</small>
                <input type="text" class="rule-new" placeholder="Örn: 10.0.0.100" value="${escapeHtml(newVal)}">
            </div>
        </div>
        <button class="btn-remove-rule" title="Kuralı Sil"><i class="fa-solid fa-trash"></i></button>
    `;

    div.querySelector('.btn-remove-rule').addEventListener('click', () => {
        div.remove();
    });

    rulesList.appendChild(div);
}

function setRuleOldIP(ip) {
    const firstRuleOld = rulesList.querySelector('.rule-old');
    if (firstRuleOld && (!firstRuleOld.value || firstRuleOld.value === '192.168.2.15')) {
        firstRuleOld.value = ip;
    } else {
        addRuleRow(ip, '');
    }
}

function getRules() {
    const rules = [];
    const rows = rulesList.querySelectorAll('.rule-item');
    rows.forEach(r => {
        const oldVal = r.querySelector('.rule-old').value.trim();
        const newVal = r.querySelector('.rule-new').value.trim();
        if (oldVal && newVal) {
            rules.push({ oldText: oldVal, newText: newVal });
        }
    });
    return rules;
}

// ------------------------------------------------------------------------------
// BACKUP OPERATION
// ------------------------------------------------------------------------------
async function runBackup() {
    const dir = folderPathInput.value.trim();
    if (!dir) return;

    btnBackup.disabled = true;
    btnBackup.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Yedekleniyor...`;
    appendLog(`Yedek alınıyor...`, 'info');

    try {
        const res = await fetch(`${API_BASE}/backup`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ directory: dir })
        });
        const data = await res.json();

        if (data.success) {
            appendLog(`Yedek başarıyla alındı! (${data.totalFiles} dosya) -> "${data.backupDirectory}"`, 'success');
            alert(`Yedekleme Tamamlandı!\nYedek Klasörü: ${data.backupDirectory}`);
        } else {
            appendLog(`Yedek hatası: ${data.error}`, 'error');
        }
    } catch (err) {
        appendLog(`Bağlantı hatası: ${err.message}`, 'error');
    } finally {
        btnBackup.disabled = false;
        btnBackup.innerHTML = `<i class="fa-solid fa-shield-halved"></i> Yedeğini Al`;
    }
}

// ------------------------------------------------------------------------------
// BULK UPDATE OPERATION
// ------------------------------------------------------------------------------
async function runUpdate() {
    const dir = folderPathInput.value.trim();
    const rules = getRules();

    if (!dir) {
        alert('Lütfen geçerli bir klasör yolu girin.');
        return;
    }

    if (rules.length === 0) {
        alert('Lütfen en az bir adet geçerli Eski IP -> Yeni IP değişim kuralı girin.');
        return;
    }

    const confirmMsg = `Toplu Güncelleme Başlatılsın mı?\n\n` +
        `Uygulanacak Kurallar:\n` +
        rules.map(r => `• ${r.oldText} ➔ ${r.newText}`).join('\n') + `\n\n` +
        `Hedef Klasör: ${dir}`;

    if (!confirm(confirmMsg)) return;

    btnExecuteUpdate.disabled = true;
    btnExecuteUpdate.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Güncelleniyor...`;

    appendLog(`TOPLU GÜNCELLEME BAŞLATILDI...`, 'warning');
    rules.forEach(r => appendLog(`Kural: "${r.oldText}" ➔ "${r.newText}"`, 'info'));
    startProgressPolling();

    const options = {
        updateQueries: chkQueries.checked,
        updateConnections: chkConnections.checked,
        updateVba: chkVba.checked,
        autoBackup: chkAutoBackup.checked,
        atomicBatch: chkAtomicBatch ? chkAtomicBatch.checked : true
    };

    try {
        const res = await fetch(`${API_BASE}/update`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                directory: dir,
                rules: rules,
                options: options
            })
        });
        const data = await res.json();

        if (data.success) {
            stopProgressPolling(`Güncelleme Başarıyla Tamamlandı!`);
            appendLog(`GÜNCELLEME TAMAMLANDI!`, 'success');
            appendLog(`Toplam İşlenen: ${data.totalFilesProcessed} | Güncellenen: ${data.updatedFilesCount} | Değişim Sayısı: ${data.totalReplacements}`, 'success');

            if (data.logs) {
                data.logs.forEach(l => {
                    if (l.status === 'Updated') {
                        appendLog(`✔ Güncellendi: ${l.fileName} (${l.changesMade} değişiklik)`, 'success');
                    } else if (l.status === 'No Changes') {
                        appendLog(`- Değişiklik Yok: ${l.fileName}`, 'info');
                    } else if (l.status === 'Error') {
                        appendLog(`✖ Hata: ${l.fileName}`, 'error');
                    }
                });
            }

            alert(`Güncelleme Tamamlandı!\n\nGüncellenen Dosya Sayısı: ${data.updatedFilesCount}\nToplam Değişiklik: ${data.totalReplacements}`);
            runScan(); // Re-scan to reflect fresh values
        } else {
            stopProgressPolling(`Güncelleme Hatayla Sonlandı!`);
            appendLog(`Hata: ${data.error}`, 'error');
            if (data.batchStatus === 'ROLLED_BACK') {
                appendLog(`GÜVENLİ İŞLEM: Hata nedeniyle yapılan tüm değişiklikler geri alındı (Rollback). Orijinal dosyalar korundu.`, 'warning');
                alert(`İşlem Sırasında Hata Oluştu!\n\nVeri güvenliği gereği yapılan tüm değişiklikler otomatik olarak geri alındı (Rollback).\nOrijinal dosyalarınız korunmuştur.\n\nHata Detayı: ${data.error}`);
            } else if (data.batchStatus === 'ROLLBACK_PARTIAL_FAILURE') {
                appendLog(`KRİTİK UYARI: Geri alma (Rollback) işlemi kısmen başarısız oldu! Bazı dosyalar yedekten geri yüklenemedi!`, 'error');
                if (data.failedRestoreFiles && data.failedRestoreFiles.length > 0) {
                    data.failedRestoreFiles.forEach(fr => {
                        appendLog(`Kritik Dosya: ${fr.OriginalPath} - Hata: ${fr.Error} - Yedek Konumu: ${fr.BackupPath}`, 'error');
                    });
                }
                alert(`KRİTİK HATA: Otomatik Geri Alma (Rollback) Kısmen Başarısız Oldu!\n\nYedek klasöründen manuel geri yükleme yapmanız gerekebilir.\n\nYedek Konumu: ${data.backupDirectory}\n\nHata: ${data.error}`);
            }
            if (data.logs) {
                data.logs.forEach(l => {
                    if (l.criticalRecovery) {
                        appendLog(`ACİL MANUEL KURTARMA GEREKLİ: Orijinal dosya: ${l.criticalRecovery.originalExpectedPath} -> Kurtarma dosyası: ${l.criticalRecovery.recoveryFilePath}`, 'error');
                        alert(`ACİL MANUEL KURTARMA GEREKİYOR!\n\nİki aşamalı kayıt sırasında beklenmeyen hata oluştu.\nOrijinal dosyanız silinmedi, şu isimle bekliyor:\n${l.criticalRecovery.recoveryFilePath}\n\nLütfen bu dosyanın adını orijinal adına çevirin.`);
                    }
                    if (l.status === 'Error') {
                        appendLog(`✖ Hatalı Dosya: ${l.fileName} - ${l.details ? l.details.join(', ') : ''}`, 'error');
                    }
                });
            }
        }
    } catch (err) {
        stopProgressPolling(`Bağlantı Hatası!`);
        appendLog(`Bağlantı hatası: ${err.message}`, 'error');
    } finally {
        btnExecuteUpdate.disabled = false;
        btnExecuteUpdate.innerHTML = `<i class="fa-solid fa-bolt"></i> Toplu Güncellemeyi Başlat`;
    }
}

function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}
