// ==============================================================================
// Excel SQL Connect Pro - Frontend Application Logic (Phase 3)
// ==============================================================================

const API_BASE = window.location.origin + '/api';

let currentScanData = null;
let activeOperationId = null;
let activeOpPollingInterval = null;
let serverWorkspace = '';
let userDesktop = '';
let userDownloads = '';
let currentBrowserDir = '';
let parentBrowserDir = '';
let selectedRestoreBackup = null;

// DOM Elements - Navigation & Status
const serverStatus = document.getElementById('serverStatus');
const navTabs = document.querySelectorAll('.nav-tab');
const tabContents = document.querySelectorAll('.tab-content');

// Controls
const folderPathInput = document.getElementById('folderPath');
const btnBrowse = document.getElementById('btnBrowse');
const btnScan = document.getElementById('btnScan');
const btnOpenFolder = document.getElementById('btnOpenFolder');
const btnBackup = document.getElementById('btnBackup');
const btnPreviewUpdate = document.getElementById('btnPreviewUpdate');
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

// Progress Section
const progressSection = document.getElementById('progressSection');
const progressDot = document.getElementById('progressDot');
const progressStatusText = document.getElementById('progressStatusText');
const progressStageBadge = document.getElementById('progressStageBadge');
const progressPercent = document.getElementById('progressPercent');
const progressBar = document.getElementById('progressBar');
const btnCancelOperation = document.getElementById('btnCancelOperation');

// Table & Log
const tableTitle = document.getElementById('tableTitle');
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

// Confirmation Modal
const confirmationModal = document.getElementById('confirmationModal');
const btnConfirmClose = document.getElementById('btnConfirmClose');
const btnCancelConfirm = document.getElementById('btnCancelConfirm');
const btnProceedUpdate = document.getElementById('btnProceedUpdate');
const chkConfirmRisks = document.getElementById('chkConfirmRisks');
const confirmTargetDir = document.getElementById('confirmTargetDir');
const confirmFileCount = document.getElementById('confirmFileCount');
const confirmRulesList = document.getElementById('confirmRulesList');
const confirmBackupBadge = document.getElementById('confirmBackupBadge');
const confirmRollbackBadge = document.getElementById('confirmRollbackBadge');

// Restore Modal
const restoreConfirmModal = document.getElementById('restoreConfirmModal');
const btnRestoreClose = document.getElementById('btnRestoreClose');
const btnCancelRestore = document.getElementById('btnCancelRestore');
const btnProceedRestore = document.getElementById('btnProceedRestore');
const restoreBackupDirText = document.getElementById('restoreBackupDirText');
const restoreTargetDirText = document.getElementById('restoreTargetDirText');

// History & Detail Modal
const historyTableBody = document.getElementById('historyTableBody');
const btnRefreshHistory = document.getElementById('btnRefreshHistory');
const historyDetailModal = document.getElementById('historyDetailModal');
const btnHistoryDetailClose = document.getElementById('btnHistoryDetailClose');
const historyDetailTitle = document.getElementById('historyDetailTitle');
const historyDetailBody = document.getElementById('historyDetailBody');

// Backups View
const backupsTableBody = document.getElementById('backupsTableBody');
const btnRefreshBackups = document.getElementById('btnRefreshBackups');

// Diagnostics View
const btnRunDiagnostics = document.getElementById('btnRunDiagnostics');
const diagExcelBody = document.getElementById('diagExcelBody');
const diagDiskBody = document.getElementById('diagDiskBody');
const diagProcBody = document.getElementById('diagProcBody');
const diagQueueBody = document.getElementById('diagQueueBody');

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
    checkActiveOperationOnLoad();
});

function bindEvents() {
    // Tab switching
    navTabs.forEach(tab => {
        tab.addEventListener('click', () => switchTab(tab.dataset.tab));
    });

    // Folder Browser
    btnBrowse.addEventListener('click', openFolderBrowserModal);
    btnFolderModalClose.addEventListener('click', () => { folderBrowserModal.style.display = 'none'; });
    btnPathUp.addEventListener('click', () => { if (parentBrowserDir) loadFolderTree(parentBrowserDir); });
    btnSelectCurrentFolder.addEventListener('click', selectFolderFromModal);

    btnShortcutDesktop.addEventListener('click', () => { if (userDesktop) loadFolderTree(userDesktop); });
    btnShortcutDownloads.addEventListener('click', () => { if (userDownloads) loadFolderTree(userDownloads); });
    btnShortcutProject.addEventListener('click', () => { if (serverWorkspace) loadFolderTree(serverWorkspace); });
    btnShortcutC.addEventListener('click', () => loadFolderTree('C:\\'));
    btnShortcutD.addEventListener('click', () => loadFolderTree('D:\\'));

    // Operations
    btnScan.addEventListener('click', runScan);
    btnOpenFolder.addEventListener('click', openFolder);
    btnBackup.addEventListener('click', runBackup);
    btnPreviewUpdate.addEventListener('click', runPreview);
    btnExecuteUpdate.addEventListener('click', openConfirmationModal);
    btnCancelOperation.addEventListener('click', cancelActiveOperation);
    btnAddRule.addEventListener('click', () => addRuleRow());
    btnClearLog.addEventListener('click', () => { logConsole.innerHTML = ''; });

    // Table Search
    tableSearch.addEventListener('input', filterTable);

    // Modals
    btnModalClose.addEventListener('click', () => { detailModal.style.display = 'none'; });
    btnConfirmClose.addEventListener('click', () => { confirmationModal.style.display = 'none'; });
    btnCancelConfirm.addEventListener('click', () => { confirmationModal.style.display = 'none'; });
    chkConfirmRisks.addEventListener('change', () => {
        btnProceedUpdate.disabled = !chkConfirmRisks.checked;
    });
    btnProceedUpdate.addEventListener('click', executeConfirmedUpdate);

    btnRestoreClose.addEventListener('click', () => { restoreConfirmModal.style.display = 'none'; });
    btnCancelRestore.addEventListener('click', () => { restoreConfirmModal.style.display = 'none'; });
    btnProceedRestore.addEventListener('click', executeConfirmedRestore);

    btnHistoryDetailClose.addEventListener('click', () => { historyDetailModal.style.display = 'none'; });

    // History & Backups & Diag buttons
    if (btnRefreshHistory) btnRefreshHistory.addEventListener('click', loadHistory);
    if (btnRefreshBackups) btnRefreshBackups.addEventListener('click', loadBackups);
    if (btnRunDiagnostics) btnRunDiagnostics.addEventListener('click', runDiagnostics);

    window.addEventListener('click', (e) => {
        if (e.target === detailModal) detailModal.style.display = 'none';
        if (e.target === folderBrowserModal) folderBrowserModal.style.display = 'none';
        if (e.target === confirmationModal) confirmationModal.style.display = 'none';
        if (e.target === restoreConfirmModal) restoreConfirmModal.style.display = 'none';
        if (e.target === historyDetailModal) historyDetailModal.style.display = 'none';
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
// TAB SWITCHING
// ------------------------------------------------------------------------------
function switchTab(targetId) {
    navTabs.forEach(t => {
        t.classList.toggle('active', t.dataset.tab === targetId);
    });
    tabContents.forEach(c => {
        if (c.id === targetId) {
            c.classList.remove('hidden');
        } else {
            c.classList.add('hidden');
        }
    });

    if (targetId === 'tabHistory') loadHistory();
    if (targetId === 'tabBackups') loadBackups();
    if (targetId === 'tabDiagnostics') runDiagnostics();
}

// ------------------------------------------------------------------------------
// ASYNC OPERATION POLLING & DISPATCH
// ------------------------------------------------------------------------------
async function checkActiveOperationOnLoad() {
    try {
        const res = await fetch(`${API_BASE}/operations/active`);
        const data = await res.json();
        if (data.active && data.operation) {
            appendLog(`Mevcut aktif işlem tespit edildi: ID ${data.operation.id} (${data.operation.type})`, 'info');
            startOperationPolling(data.operation.id, data.operation.type);
        }
    } catch (e) { }
}

function startOperationPolling(operationId, operationType) {
    activeOperationId = operationId;
    progressSection.classList.remove('hidden');
    progressDot.className = 'status-dot pulsing';
    btnCancelOperation.disabled = false;
    btnCancelOperation.style.display = 'inline-block';
    btnCancelOperation.innerHTML = `<i class="fa-solid fa-ban"></i> İptal Et`;

    if (activeOpPollingInterval) clearInterval(activeOpPollingInterval);

    activeOpPollingInterval = setInterval(async () => {
        try {
            const res = await fetch(`${API_BASE}/operations/${operationId}`);
            if (!res.ok) return;
            const data = await res.json();
            const op = data.operation;
            if (!op) return;

            updateProgressDisplay(op);

            const terminalStates = ['COMPLETED', 'FAILED', 'CANCELLED', 'ROLLED_BACK', 'ROLLBACK_PARTIAL_FAILURE', 'STALE_INTERRUPTED'];
            if (terminalStates.includes(op.status)) {
                clearInterval(activeOpPollingInterval);
                activeOpPollingInterval = null;
                handleOperationFinished(op);
            }
        } catch (err) {
            console.error("Polling hatası:", err);
        }
    }, 600);
}

function updateProgressDisplay(op) {
    const fileLabel = op.currentFile || 'Hazırlanıyor';
    progressStatusText.textContent = `${fileLabel} (${op.processedFiles || 0} / ${op.totalFiles || 0})`;
    progressPercent.textContent = `${op.progressPercent || 0}%`;
    progressBar.style.width = `${op.progressPercent || 0}%`;
    progressStageBadge.textContent = `Aşama: ${op.currentStage || op.status}`;

    if (op.status === 'CANCELLATION_REQUESTED') {
        btnCancelOperation.disabled = true;
        btnCancelOperation.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> İptal Ediliyor...`;
        progressStatusText.textContent = `İptal talep edildi, güvenli checkpoint bekleniyor...`;
    } else if (op.status === 'COMMITTING') {
        btnCancelOperation.disabled = true;
        btnCancelOperation.innerHTML = `<i class="fa-solid fa-lock"></i> Kilitlendi`;
        progressStageBadge.textContent = `Kalıcı Commit (Geri Alınamaz)`;
    } else if (op.status === 'ROLLING_BACK') {
        btnCancelOperation.disabled = true;
        btnCancelOperation.innerHTML = `<i class="fa-solid fa-rotate-left"></i> Geri Alınıyor`;
        progressStageBadge.textContent = `Rollback Uygulanıyor`;
    }
}

function handleOperationFinished(op) {
    progressDot.className = 'status-dot';
    btnCancelOperation.style.display = 'none';

    // Re-enable trigger buttons
    btnExecuteUpdate.disabled = false;
    btnExecuteUpdate.innerHTML = `<i class="fa-solid fa-bolt"></i> Toplu Güncellemeyi Başlat`;
    btnPreviewUpdate.disabled = false;
    btnPreviewUpdate.innerHTML = `<i class="fa-solid fa-eye"></i> Önizlemeyi Başlat (Dry Run)`;
    btnScan.disabled = false;
    btnScan.innerHTML = `<i class="fa-solid fa-magnifying-glass-chart"></i> Klasörü Tara`;

    const resData = op.resultData || {};

    if (op.status === 'COMPLETED') {
        progressBar.style.width = '100%';
        progressPercent.textContent = '100%';
        progressStatusText.textContent = 'İşlem Başarıyla Tamamlandı!';
        progressStageBadge.textContent = 'Tamamlandı';

        if (op.type === 'PREVIEW') {
            appendLog(`ÖNİZLEME (DRY RUN) TAMAMLANDI! Dosyalarda değişiklik yapılmadı.`, 'success');
            appendLog(`İncelenen Dosya: ${resData.totalFilesScanned || op.processedFiles} | Değişecek Dosya: ${resData.prospectiveUpdatedFilesCount || 0} | Potansiyel Değişiklik: ${resData.totalProspectiveReplacements || 0}`, 'info');
            tableTitle.innerHTML = `<i class="fa-solid fa-eye text-primary"></i> Önizleme (Dry Run) Sonuçları <span class="badge badge-warning">Değişiklik Yapılmadı</span>`;
            if (resData.files) {
                currentScanData = { files: resData.files, totalFiles: resData.files.length };
                renderTable(resData.files);
            }
            alert(`Önizleme (Dry Run) Tamamlandı!\n\nDosyalara DOKUNULMADI.\nDeğişecek Dosya Sayısı: ${resData.prospectiveUpdatedFilesCount || 0}\nToplam Potansiyel Değişiklik: ${resData.totalProspectiveReplacements || 0}`);
        } else if (op.type === 'UPDATE') {
            appendLog(`TOPLU GÜNCELLEME TAMAMLANDI!`, 'success');
            appendLog(`Toplam İşlenen: ${op.processedFiles} | Güncellenen: ${op.updatedFiles || resData.updatedFilesCount || 0}`, 'success');
            if (op.logs) {
                op.logs.forEach(l => {
                    if (l.status === 'Updated') {
                        appendLog(`✔ Güncellendi: ${l.fileName} (${l.changesMade} değişiklik)`, 'success');
                    } else if (l.status === 'No Changes') {
                        appendLog(`- Değişiklik Yok: ${l.fileName}`, 'info');
                    }
                });
            }
            alert(`Güncelleme Başarıyla Tamamlandı!\n\nGüncellenen Dosya: ${op.updatedFiles || resData.updatedFilesCount || 0}\nYedek Konumu: ${op.backupDirectory || 'Alınmadı'}`);
            runScan(); // Re-scan to show current values
        } else if (op.type === 'RESTORE') {
            appendLog(`GERİ YÜKLEME (RESTORE) TAMAMLANDI! Dosyalar orijinal hallerine döndürüldü.`, 'success');
            alert(`Geri Yükleme Tamamlandı!\nDosyalar doğrulanarak orijinal hallerine getirildi.`);
            runScan();
        }
    } else if (op.status === 'CANCELLED') {
        progressBar.style.width = '100%';
        progressStatusText.textContent = 'İşlem Kullanıcı Tarafından İptal Edildi.';
        progressStageBadge.textContent = 'İptal Edildi';
        appendLog(`[GÜVENLİ İPTAL] İşlem iptal edildi. Orijinal dosyalar güvenle korundu.`, 'warning');
        alert(`İşlem Güvenle İptal Edildi.\nOrijinal dosyalarınızda herhangi bir hasar oluşmadı.`);
    } else if (op.status === 'ROLLED_BACK') {
        progressBar.style.width = '100%';
        progressStatusText.textContent = 'Hata Nedeniyle Tüm Değişiklikler Geri Alındı (Rollback).';
        progressStageBadge.textContent = 'Geri Alındı';
        appendLog(`[ROLLBACK] Hata oluştuğu için yapılan tüm değişiklikler yedekten geri yüklendi. Orijinal dosyalar korundu.`, 'warning');
        alert(`İşlem Sırasında Hata Oluştu!\n\nVeri güvenliği mimarisi gereği yapılan tüm değişiklikler otomatik olarak geri alındı (Rollback).\nOrijinal dosyalarınız korunmuştur.\n\nHata: ${(op.errors || []).join('; ')}`);
    } else if (op.status === 'ROLLBACK_PARTIAL_FAILURE') {
        progressStatusText.textContent = 'KRİTİK UYARI: Geri alma işlemi kısmen başarısız oldu!';
        progressStageBadge.textContent = 'Kısmi Hata';
        appendLog(`[KRİTİK] Rollback sırasında bazı dosyalar geri yüklenemedi! Yedek klasörü: ${op.backupDirectory}`, 'error');
        alert(`KRİTİK UYARI: Rollback Kısmen Başarısız!\n\nYedek klasöründen manuel geri yükleme yapmanız gerekebilir.\nYedek: ${op.backupDirectory}`);
    } else {
        progressStatusText.textContent = `İşlem Başarısız: ${(op.errors || []).join('; ')}`;
        progressStageBadge.textContent = 'Başarısız';
        appendLog(`[HATA] İşlem başarısız oldu: ${(op.errors || []).join('; ')}`, 'error');
        alert(`İşlem Başarısız Oldu!\n\nHata Detayı: ${(op.errors || []).join('; ')}`);
    }
}

// ------------------------------------------------------------------------------
// CANCELLATION REQUEST
// ------------------------------------------------------------------------------
async function cancelActiveOperation() {
    if (!activeOperationId) return;

    btnCancelOperation.disabled = true;
    btnCancelOperation.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> İptal Talep Ediliyor...`;
    appendLog(`[OPERASYON] İptal talebi gönderiliyor: ${activeOperationId}...`, 'warning');

    try {
        const res = await fetch(`${API_BASE}/operations/${activeOperationId}/cancel`, {
            method: 'POST'
        });
        const data = await res.json();
        if (data.success) {
            appendLog(`[OPERASYON] İptal isteği kabul edildi. Güvenli checkpoint bekleniyor...`, 'info');
        } else {
            appendLog(`[OPERASYON] İptal reddedildi: ${data.message || data.error}`, 'error');
            btnCancelOperation.disabled = false;
            btnCancelOperation.innerHTML = `<i class="fa-solid fa-ban"></i> İptal Et`;
        }
    } catch (err) {
        appendLog(`[HATA] İptal isteği iletilemedi: ${err.message}`, 'error');
        btnCancelOperation.disabled = false;
        btnCancelOperation.innerHTML = `<i class="fa-solid fa-ban"></i> İptal Et`;
    }
}

// ------------------------------------------------------------------------------
// PREVIEW (DRY RUN) OPERATION
// ------------------------------------------------------------------------------
async function runPreview() {
    const dir = folderPathInput.value.trim();
    const rules = getRules();

    if (!dir) {
        alert('Lütfen geçerli bir klasör yolu girin.');
        return;
    }
    if (rules.length === 0) {
        alert('Lütfen en az bir adet IP / Metin değişim kuralı girin.');
        return;
    }

    btnPreviewUpdate.disabled = true;
    btnPreviewUpdate.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Önizleme Başlatılıyor...`;
    appendLog(`ÖNİZLEME (DRY RUN) BAŞLATILIYOR: Dosyalara DOKUNULMAZ, simülasyon yapılıyor...`, 'info');

    const options = {
        updateQueries: chkQueries.checked,
        updateConnections: chkConnections.checked,
        updateVba: chkVba.checked
    };

    try {
        const res = await fetch(`${API_BASE}/preview`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ directory: dir, rules: rules, options: options })
        });
        const data = await res.json();

        if (res.status === 202 && data.success) {
            appendLog(`Önizleme işlemi kuyruğa alındı (ID: ${data.operationId}).`, 'info');
            startOperationPolling(data.operationId, 'PREVIEW');
        } else if (res.status === 409) {
            alert(`Başka bir işlem şu anda çalışıyor (ID: ${data.activeOperationId}). Lütfen bitmesini veya iptal edilmesini bekleyin.`);
            btnPreviewUpdate.disabled = false;
            btnPreviewUpdate.innerHTML = `<i class="fa-solid fa-eye"></i> Önizlemeyi Başlat (Dry Run)`;
        } else {
            alert(`Hata: ${data.error}`);
            btnPreviewUpdate.disabled = false;
            btnPreviewUpdate.innerHTML = `<i class="fa-solid fa-eye"></i> Önizlemeyi Başlat (Dry Run)`;
        }
    } catch (err) {
        appendLog(`Bağlantı hatası: ${err.message}`, 'error');
        btnPreviewUpdate.disabled = false;
        btnPreviewUpdate.innerHTML = `<i class="fa-solid fa-eye"></i> Önizlemeyi Başlat (Dry Run)`;
    }
}

// ------------------------------------------------------------------------------
// UPDATE CONFIRMATION MODAL & EXECUTION
// ------------------------------------------------------------------------------
function openConfirmationModal() {
    const dir = folderPathInput.value.trim();
    const rules = getRules();

    if (!dir) {
        alert('Lütfen geçerli bir klasör yolu girin.');
        return;
    }
    if (rules.length === 0) {
        alert('Lütfen en az bir adet IP / Metin değişim kuralı girin.');
        return;
    }

    confirmTargetDir.textContent = dir;
    confirmFileCount.textContent = (currentScanData && currentScanData.files) ? currentScanData.files.length : 'Belirtilmemiş';

    confirmRulesList.innerHTML = '';
    rules.forEach(r => {
        const li = document.createElement('li');
        li.innerHTML = `<code>${escapeHtml(r.oldText)}</code> ➔ <code>${escapeHtml(r.newText)}</code>`;
        confirmRulesList.appendChild(li);
    });

    confirmBackupBadge.style.display = chkAutoBackup.checked ? 'inline-block' : 'none';
    confirmRollbackBadge.style.display = chkAtomicBatch.checked ? 'inline-block' : 'none';

    chkConfirmRisks.checked = false;
    btnProceedUpdate.disabled = true;

    confirmationModal.style.display = 'flex';
}

async function executeConfirmedUpdate() {
    confirmationModal.style.display = 'none';

    const dir = folderPathInput.value.trim();
    const rules = getRules();

    btnExecuteUpdate.disabled = true;
    btnExecuteUpdate.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Başlatılıyor...`;
    appendLog(`TOPLU GÜNCELLEME BAŞLATILIYOR...`, 'warning');

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
            body: JSON.stringify({ directory: dir, rules: rules, options: options })
        });
        const data = await res.json();

        if (res.status === 202 && data.success) {
            appendLog(`Güncelleme işlemi kuyruğa alındı (ID: ${data.operationId}).`, 'info');
            startOperationPolling(data.operationId, 'UPDATE');
        } else if (res.status === 409) {
            alert(`Sistem meşgul: Şu anda başka bir işlem çalışıyor (ID: ${data.activeOperationId}).`);
            btnExecuteUpdate.disabled = false;
            btnExecuteUpdate.innerHTML = `<i class="fa-solid fa-bolt"></i> Toplu Güncellemeyi Başlat`;
        } else {
            alert(`Hata: ${data.error}`);
            btnExecuteUpdate.disabled = false;
            btnExecuteUpdate.innerHTML = `<i class="fa-solid fa-bolt"></i> Toplu Güncellemeyi Başlat`;
        }
    } catch (err) {
        appendLog(`Bağlantı hatası: ${err.message}`, 'error');
        btnExecuteUpdate.disabled = false;
        btnExecuteUpdate.innerHTML = `<i class="fa-solid fa-bolt"></i> Toplu Güncellemeyi Başlat`;
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
    tableTitle.innerHTML = `<i class="fa-solid fa-table-list"></i> Tarama Sonuçları`;

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
            appendLog(`Tarama tamamlandı! Toplam ${data.totalFiles} dosya incelendi.`, 'success');
        } else {
            appendLog(`Tarama hatası: ${data.error}`, 'error');
            alert(`Tarama hatası: ${data.error}`);
        }
    } catch (err) {
        appendLog(`Bağlantı hatası: ${err.message}`, 'error');
    } finally {
        btnScan.disabled = false;
        btnScan.innerHTML = `<i class="fa-solid fa-magnifying-glass-chart"></i> Klasörü Tara`;
    }
}

function renderScanResults(data) {
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
        let statusLabel = f.status || 'Taranmış';
        if (statusLabel.startsWith('Error')) badgeClass = 'badge-danger';
        else if (statusLabel === 'Updated') badgeClass = 'badge-success';
        else if (statusLabel === 'Would Update') {
            badgeClass = 'badge-warning';
            statusLabel = `Değişecek (${f.prospectiveReplacements || 0})`;
        }

        tr.innerHTML = `
            <td><strong>${escapeHtml(f.fileName)}</strong></td>
            <td><span class="badge ${f.extension === '.xlsm' ? 'badge-warning' : 'badge-info'}">${(f.extension || '').toUpperCase()}</span></td>
            <td>${qCount > 0 ? `<span class="badge badge-info">${qCount} Sorgu</span>` : '<span class="text-dim">0</span>'}</td>
            <td>${cCount > 0 ? `<span class="badge badge-warning">${cCount} Bağlantı</span>` : '<span class="text-dim">0</span>'}</td>
            <td>${vCount > 0 ? `<span class="badge badge-success">${vCount} Makro</span>` : '<span class="text-dim">0</span>'}</td>
            <td><code class="text-primary">${escapeHtml(ipsStr)}</code></td>
            <td><span class="badge ${badgeClass}">${escapeHtml(statusLabel)}</span></td>
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
            <p><strong>SHA256:</strong> <code style="font-size:0.75rem">${escapeHtml(file.sha256 || 'Mevcut Değil')}</code></p>
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
// HISTORY TAB
// ------------------------------------------------------------------------------
async function loadHistory() {
    historyTableBody.innerHTML = `<tr><td colspan="8" class="text-center p-3"><i class="fa-solid fa-spinner fa-spin"></i> İşlem geçmişi yükleniyor...</td></tr>`;

    try {
        const res = await fetch(`${API_BASE}/history`);
        const data = await res.json();
        const list = data.history || [];

        if (list.length === 0) {
            historyTableBody.innerHTML = `
                <tr class="empty-row">
                    <td colspan="8">
                        <div class="empty-state">
                            <i class="fa-solid fa-inbox empty-icon"></i>
                            <p>Henüz kayıtlı bir geçmiş işlem bulunmuyor.</p>
                        </div>
                    </td>
                </tr>
            `;
            return;
        }

        historyTableBody.innerHTML = '';
        list.forEach(op => {
            const tr = document.createElement('tr');
            let badgeClass = 'badge-info';
            if (op.status === 'COMPLETED') badgeClass = 'badge-success';
            else if (op.status === 'CANCELLED') badgeClass = 'badge-warning';
            else if (op.status.includes('FAIL') || op.status.includes('ERR')) badgeClass = 'badge-danger';

            const startTimeStr = op.startTime ? new Date(op.startTime).toLocaleString('tr-TR') : '-';
            const durationSec = (op.startTime && op.endTime) 
                ? ((new Date(op.endTime) - new Date(op.startTime)) / 1000).toFixed(1) + 's' 
                : '-';

            tr.innerHTML = `
                <td><code>${escapeHtml(op.id.substring(0, 16))}...</code></td>
                <td>${escapeHtml(startTimeStr)}</td>
                <td><span class="badge badge-info">${escapeHtml(op.type)}</span></td>
                <td title="${escapeHtml(op.directory)}">${escapeHtml(op.directory ? op.directory.split('\\').pop() : '-')}</td>
                <td>${op.processedFiles || 0} / ${op.updatedFiles || 0}</td>
                <td>${durationSec}</td>
                <td><span class="badge ${badgeClass}">${escapeHtml(op.status)}</span></td>
                <td>
                    <button class="btn btn-outline-sm btn-hist-detail" data-id="${escapeHtml(op.id)}">
                        <i class="fa-solid fa-file-lines"></i> Rapor
                    </button>
                </td>
            `;

            tr.querySelector('.btn-hist-detail').addEventListener('click', () => showHistoryDetail(op));
            historyTableBody.appendChild(tr);
        });

    } catch (err) {
        historyTableBody.innerHTML = `<tr><td colspan="8" class="text-danger p-3">Geçmiş yüklenemedi: ${err.message}</td></tr>`;
    }
}

function showHistoryDetail(op) {
    historyDetailTitle.innerHTML = `<i class="fa-solid fa-file-lines"></i> İşlem Detayı: <code>${escapeHtml(op.id)}</code>`;
    
    let html = `
        <div class="confirm-summary-box mb-3">
            <div><strong>Tür:</strong> <span class="badge badge-info">${escapeHtml(op.type)}</span> | <strong>Durum:</strong> <strong>${escapeHtml(op.status)}</strong></div>
            <div><strong>Hedef Klasör:</strong> <code>${escapeHtml(op.directory)}</code></div>
            <div><strong>Başlangıç:</strong> ${escapeHtml(op.startTime)} | <strong>Bitiş:</strong> ${escapeHtml(op.endTime || '-')}</div>
            <div><strong>İşlenen / Güncellenen Dosya:</strong> ${op.processedFiles || 0} / ${op.updatedFiles || 0}</div>
            ${op.backupDirectory ? `<div><strong>Yedek Klasörü:</strong> <code>${escapeHtml(op.backupDirectory)}</code></div>` : ''}
        </div>
    `;

    if (op.errors && op.errors.length > 0) {
        html += `<h4 class="text-danger mb-1"><i class="fa-solid fa-triangle-exclamation"></i> Hatalar:</h4>`;
        html += `<ul class="modal-list mb-3">`;
        op.errors.forEach(e => { html += `<li class="text-danger">${escapeHtml(e)}</li>`; });
        html += `</ul>`;
    }

    if (op.logs && op.logs.length > 0) {
        html += `<h4><i class="fa-solid fa-list-check"></i> Dosya Günlüğü (${op.logs.length} Dosya):</h4>`;
        html += `<ul class="modal-list">`;
        op.logs.forEach(l => {
            html += `<li><strong>${escapeHtml(l.fileName)}</strong>: <span class="badge ${l.status === 'Updated' ? 'badge-success' : 'badge-info'}">${escapeHtml(l.status)}</span> (${l.changesMade || 0} Değişiklik)</li>`;
        });
        html += `</ul>`;
    }

    historyDetailBody.innerHTML = html;
    historyDetailModal.style.display = 'flex';
}

// ------------------------------------------------------------------------------
// BACKUPS & RESTORE TAB
// ------------------------------------------------------------------------------
async function loadBackups() {
    const dir = folderPathInput.value.trim();
    backupsTableBody.innerHTML = `<tr><td colspan="5" class="text-center p-3"><i class="fa-solid fa-spinner fa-spin"></i> Yedekler taranıyor...</td></tr>`;

    try {
        const res = await fetch(`${API_BASE}/backups?dir=${encodeURIComponent(dir)}`);
        const data = await res.json();
        const backups = data.backups || [];

        if (backups.length === 0) {
            backupsTableBody.innerHTML = `
                <tr class="empty-row">
                    <td colspan="5">
                        <div class="empty-state">
                            <i class="fa-solid fa-box-archive empty-icon"></i>
                            <p>Hedef klasörde henüz oluşturulmuş bir yedek arşivi bulunamadı.</p>
                        </div>
                    </td>
                </tr>
            `;
            return;
        }

        backupsTableBody.innerHTML = '';
        backups.forEach(b => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${escapeHtml(b.timestamp)}</strong></td>
                <td><code>${escapeHtml(b.folderName)}</code></td>
                <td><span class="badge badge-info">${b.fileCount} Dosya</span></td>
                <td>${escapeHtml(b.operation || 'AutoBackup')}</td>
                <td>
                    <button class="btn btn-danger-sm btn-restore-item">
                        <i class="fa-solid fa-rotate-left"></i> Bu Yedeğe Geri Dön
                    </button>
                </td>
            `;

            tr.querySelector('.btn-restore-item').addEventListener('click', () => {
                selectedRestoreBackup = b.backupDirectory;
                restoreBackupDirText.textContent = b.backupDirectory;
                restoreTargetDirText.textContent = dir;
                restoreConfirmModal.style.display = 'flex';
            });

            backupsTableBody.appendChild(tr);
        });

    } catch (err) {
        backupsTableBody.innerHTML = `<tr><td colspan="5" class="text-danger p-3">Yedekler yüklenemedi: ${err.message}</td></tr>`;
    }
}

async function executeConfirmedRestore() {
    restoreConfirmModal.style.display = 'none';
    if (!selectedRestoreBackup) return;

    const dir = folderPathInput.value.trim();
    appendLog(`[RESTORE] Doğrulanmış yedekten geri yükleme başlatılıyor: ${selectedRestoreBackup}...`, 'warning');
    switchTab('tabDashboard');

    try {
        const res = await fetch(`${API_BASE}/restore`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ directory: dir, backupDir: selectedRestoreBackup })
        });
        const data = await res.json();

        if (res.status === 202 && data.success) {
            appendLog(`Restore işlemi kuyruğa alındı (ID: ${data.operationId}).`, 'info');
            startOperationPolling(data.operationId, 'RESTORE');
        } else {
            alert(`Restore başlatılamadı: ${data.error}`);
        }
    } catch (err) {
        appendLog(`Restore bağlantı hatası: ${err.message}`, 'error');
    }
}

// ------------------------------------------------------------------------------
// SYSTEM DIAGNOSTICS TAB
// ------------------------------------------------------------------------------
async function runDiagnostics() {
    btnRunDiagnostics.disabled = true;
    btnRunDiagnostics.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Kontrol Ediliyor...`;

    try {
        const res = await fetch(`${API_BASE}/diagnostics`);
        const diag = await res.json();

        // 1. Excel COM
        let excelHtml = `
            <div class="diag-row"><span>Excel Kurulu:</span> <span class="${diag.ExcelComAvailable ? 'diag-status-ok' : 'diag-status-err'}">${diag.ExcelComAvailable ? 'Evet' : 'Hayır'}</span></div>
            <div class="diag-row"><span>Excel Sürümü:</span> <span>${escapeHtml(diag.ExcelVersion || 'Bilinmiyor')}</span></div>
            <div class="diag-row"><span>İşlemci Mimarisi:</span> <span>${escapeHtml(diag.ExcelBitness || 'Bilinmiyor')}</span></div>
            <div class="diag-row"><span>COM Yanıt Süresi:</span> <span>${diag.ExcelComProbeElapsedMs || 0} ms</span></div>
        `;
        diagExcelBody.innerHTML = excelHtml;

        // 2. Disk Permissions
        let diskHtml = `
            <div class="diag-row"><span>Çalışma Dizini:</span> <span class="${diag.WorkingDirWritable ? 'diag-status-ok' : 'diag-status-err'}">${diag.WorkingDirWritable ? 'Yazılabilir' : 'Kilitli'}</span></div>
            <div class="diag-row"><span>Yedek Dizini:</span> <span class="${diag.BackupDirWritable ? 'diag-status-ok' : 'diag-status-err'}">${diag.BackupDirWritable ? 'Yazılabilir' : 'Kilitli'}</span></div>
            <div class="diag-row"><span>Temp Dizini:</span> <span class="${diag.TempDirWritable ? 'diag-status-ok' : 'diag-status-err'}">${diag.TempDirWritable ? 'Yazılabilir' : 'Kilitli'}</span></div>
        `;
        diagDiskBody.innerHTML = diskHtml;

        // 3. Process Monitor
        let pidsStr = (diag.ActiveExcelPids && diag.ActiveExcelPids.length > 0) ? diag.ActiveExcelPids.join(', ') : 'Yok';
        let procHtml = `
            <div class="diag-row"><span>Açık Excel Süreçleri:</span> <span class="badge badge-info">${diag.ActiveExcelProcessCount || 0} Adet</span></div>
            <div class="diag-row"><span>Mevcut PID'ler:</span> <code>${escapeHtml(pidsStr)}</code></div>
            <div class="diag-row"><span>Kullanıcı Oturumu Koruması:</span> <span class="diag-status-ok">Devrede</span></div>
        `;
        diagProcBody.innerHTML = procHtml;

        // 4. Queue & State
        let queueHtml = `
            <div class="diag-row"><span>Aktif İşlem Var mı:</span> <span class="${diag.ActiveOperationId ? 'diag-status-warn' : 'diag-status-ok'}">${diag.ActiveOperationId ? 'Evet (' + diag.ActiveOperationId.substring(0, 8) + '...)' : 'Boşta'}</span></div>
            <div class="diag-row"><span>Kayıtlı İşlem Sayısı:</span> <span>${diag.ActiveOperationsCount || 0}</span></div>
            <div class="diag-row"><span>Mutex Kilit Durumu:</span> <span class="diag-status-ok">Normal</span></div>
        `;
        diagQueueBody.innerHTML = queueHtml;

    } catch (err) {
        console.error("Diagnostics error:", err);
    } finally {
        btnRunDiagnostics.disabled = false;
        btnRunDiagnostics.innerHTML = `<i class="fa-solid fa-heart-pulse"></i> Sağlık Kontrolünü Çalıştır`;
    }
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
    appendLog(`Manuel yedek alınıyor...`, 'info');

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
// INTERACTIVE FOLDER BROWSER
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

            subfoldersList.innerHTML = '';

            if (data.drives && data.drives.length > 0) {
                data.drives.forEach(d => {
                    const card = document.createElement('div');
                    card.className = 'subfolder-card';
                    card.innerHTML = `<i class="fa-solid fa-hard-drive"></i> ${escapeHtml(d.name)}`;
                    card.addEventListener('click', () => loadFolderTree(d.path));
                    subfoldersList.appendChild(card);
                });
            }

            if (data.subFolders && data.subFolders.length > 0) {
                data.subFolders.forEach(sub => {
                    const card = document.createElement('div');
                    card.className = 'subfolder-card';
                    card.innerHTML = `<i class="fa-solid fa-folder"></i> ${escapeHtml(sub.name)}`;
                    card.addEventListener('click', () => loadFolderTree(sub.path));
                    subfoldersList.appendChild(card);
                });
            }

            if ((!data.subFolders || data.subFolders.length === 0) && (!data.drives || data.drives.length === 0)) {
                subfoldersList.innerHTML = `<div class="subfolder-card-placeholder text-dim">Alt klasör bulunamadı.</div>`;
            }
        } else {
            subfoldersList.innerHTML = `<div class="subfolder-card-placeholder text-danger">Klasör okunamadı: ${data.error}</div>`;
        }
    } catch (err) {
        subfoldersList.innerHTML = `<div class="subfolder-card-placeholder text-danger">Hata: ${err.message}</div>`;
    }
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

function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}
