// JavaScript Shipping Data Pengguna
// Integrasi Firebase, pencarian, pagination, dan ekspor Excel dipertahankan.
(function clock() {
    const timeEl = document.getElementById('clock-time');
    const dateEl = document.getElementById('clock-date');
    if (!timeEl || !dateEl) return;
    const now = new Date();
    timeEl.textContent = now.toLocaleTimeString('id-ID', { hour12: false });
    dateEl.textContent = now.toLocaleDateString('id-ID', {
        weekday: 'short', day: '2-digit', month: 'short', year: 'numeric'
    });
    setTimeout(clock, 1000);
})();

import {initializeApp} from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js';
            import {getFirestore, collection, getDocs} from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js';

            // Konfigurasi Firebase mengikuti aplikasi utama yang diberikan.
            const firebaseConfig = {
                apiKey: 'AIzaSyAdjD2FiCU-uP5eo8JuKyP_Gc_QNGpAPK8',
                authDomain: 'inspection-data-11f7d.firebaseapp.com',
                projectId: 'inspection-data-11f7d',
                storageBucket: 'inspection-data-11f7d.firebasestorage.app',
                messagingSenderId: '1082215196163',
                appId: '1:1082215196163:web:fdaa360881782ea3ad1036',
                measurementId: 'G-KBV92YDP1K'
            };
            const COLLECTION_NAME = 'shippingData';
            const app = initializeApp(firebaseConfig);
            const db = getFirestore(app);
            const fields = [
                {key: 'PO Date', label: 'PO Date', type: 'date'}, {key: 'Shipper', label: 'Shipper'}, {key: 'SO', label: 'SO'}, {key: 'Vendor', label: 'Vendor'}, {key: 'PO', label: 'PO'},
                {key: 'Container Number', label: 'Container Number'}, {key: 'ECRD', label: 'ECRD', type: 'date'},
                {key: 'ETD', label: 'ETD', type: 'date'}, {key: 'ETA', label: 'ETA', type: 'date'},
                {key: 'Payment Status', label: 'Payment Status', type: 'text'}, {key: 'Payment Date', label: 'Payment Date', type: 'date'}
            ];
            let records = [];
            let currentPage = 1;
            let pageSize = 10;
            const $ = id => document.getElementById(id);
            const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]));
            function setStatus(id, message, kind = '') {const el = $(id); el.textContent = message; el.className = 'status' + (kind ? ' ' + kind : '');}
            function normalizeHeader(s) {return String(s ?? '').trim().toLowerCase().replace(/[_\s-]+/g, '').replace(/[^a-z0-9]/g, '');}
            function toText(v) {
                if (v === null || v === undefined) return '';
                if (v instanceof Date && !isNaN(v)) return `${v.getMonth() + 1}/${v.getDate()}/${v.getFullYear()}`;
                return String(v).trim();
            }
            function excelDateToText(value, cell) {
                if (value === null || value === undefined || value === '') return '';
                if (cell && cell.t === 'n' && (cell.z || '').match(/[dmy]/i)) {
                    const d = XLSX.SSF.parse_date_code(value); if (d) return `${d.m}/${d.d}/${d.y}`;
                }
                if (typeof value === 'number' && value > 20000 && value < 80000) {
                    const d = XLSX.SSF.parse_date_code(value); if (d) return `${d.m}/${d.d}/${d.y}`;
                }
                return toText(value);
            }
            function dateForInput(value) {
                const s = String(value ?? '').trim();
                if (!s) return '';
                if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
                let m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
                if (m) {
                    const mm = m[1].padStart(2, '0'), dd = m[2].padStart(2, '0'), yyyy = m[3];
                    return `${yyyy}-${mm}-${dd}`;
                }
                const d = new Date(s);
                if (!Number.isNaN(d.getTime())) return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
                return '';
            }
            function getVisibleRecords() {
                const q = $('searchInput').value.trim().toLowerCase();
                const selectedField = $('searchColumn').value || 'Shipper';
                return records.filter(r => !q || String(r.data[selectedField] ?? '').toLowerCase().includes(q));
            }
            function renderTable() {
                const shown = getVisibleRecords();
                const total = shown.length;
                const totalPages = Math.max(1, Math.ceil(total / pageSize));
                currentPage = Math.min(Math.max(1, currentPage), totalPages);
                const startIndex = (currentPage - 1) * pageSize;
                const pageRecords = shown.slice(startIndex, startIndex + pageSize);
                $('rowCount').textContent = `${total} dari ${records.length} data`;
                if (!pageRecords.length) {
                    $('tableBody').innerHTML = `<tr><td colspan="11" class="empty">${records.length ? 'Tidak ada data yang cocok dengan pencarian.' : 'Belum ada data yang tersedia.'}</td></tr>`;
                } else {
                    $('tableBody').innerHTML = pageRecords.map(record => `<tr>${fields.map(f => {
                        const val = record.data[f.key] ?? '';
                        if (f.key === 'Payment Status') {
                            const normalized = String(val).trim().toLowerCase();
                            const status = normalized === 'ok' ? 'Ok' : (normalized === 'queue' || normalized === 'queueue' || normalized === 'queque' ? 'Queue' : String(val || ''));
                            const statusClass = status === 'Ok' ? 'payment-status-ok' : (status === 'Queue' ? 'payment-status-queue' : '');
                            return `<td><span class="payment-status-select ${statusClass}">${esc(status)}</span></td>`;
                        }
                        return `<td>${esc(val)}</td>`;
                    }).join('')}</tr>`).join('');
                }
                const first = total === 0 ? 0 : startIndex + 1;
                const last = Math.min(startIndex + pageSize, total);
                $('paginationInfo').textContent = `Menampilkan ${first}–${last} dari ${total} data`;
                let controls = `<button class="page-btn" type="button" data-page="${currentPage - 1}" ${currentPage === 1 ? 'disabled' : ''}>‹ Sebelumnya</button>`;
                const maxVisible = 5;
                let pageStart = Math.max(1, currentPage - 2);
                let pageEnd = Math.min(totalPages, pageStart + maxVisible - 1);
                pageStart = Math.max(1, pageEnd - maxVisible + 1);
                if (pageStart > 1) controls += `<button class="page-btn" type="button" data-page="1">1</button>${pageStart > 2 ? '<span>…</span>' : ''}`;
                for (let page = pageStart; page <= pageEnd; page++) controls += `<button class="page-btn ${page === currentPage ? 'active' : ''}" type="button" data-page="${page}" ${page === currentPage ? 'aria-current="page"' : ''}>${page}</button>`;
                if (pageEnd < totalPages) controls += `${pageEnd < totalPages - 1 ? '<span>…</span>' : ''}<button class="page-btn" type="button" data-page="${totalPages}">${totalPages}</button>`;
                controls += `<button class="page-btn" type="button" data-page="${currentPage + 1}" ${currentPage === totalPages ? 'disabled' : ''}>Berikutnya ›</button>`;
                $('paginationButtons').innerHTML = controls;
            }
            function dateSortValue(value) {
                const s = String(value ?? '').trim();
                if (!s) return Number.POSITIVE_INFINITY;
                let match = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
                if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])).getTime();
                match = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
                if (match) return new Date(Number(match[3]), Number(match[1]) - 1, Number(match[2])).getTime();
                const parsed = new Date(s);
                return Number.isNaN(parsed.getTime()) ? Number.POSITIVE_INFINITY : parsed.getTime();
            }
            async function loadData() {
                $('refreshBtn').disabled = true;
                try {
                    const snap = await getDocs(collection(db, COLLECTION_NAME));
                    records = snap.docs.map(d => ({id: d.id, data: d.data()}));
                    records.sort((a, b) => {
                        const dateDiff = dateSortValue(a.data['PO Date']) - dateSortValue(b.data['PO Date']);
                        return Number.isNaN(dateDiff) || dateDiff === 0
                            ? String(a.data.SO || '').localeCompare(String(b.data.SO || ''))
                            : dateDiff;
                    });
                    renderTable(); $('connectionPill').textContent = 'Firebase terhubung'; $('connectionPill').style.background = '#e8f7ee'; $('connectionPill').style.color = '#067647';
                    setStatus('tableStatus', 'Data berhasil dimuat.', 'success');
                } catch (err) {
                    console.error(err); $('connectionPill').textContent = 'Koneksi bermasalah'; $('connectionPill').style.background = '#fff0f0'; $('connectionPill').style.color = '#b42318';
                    setStatus('tableStatus', `Gagal membaca Firestore: ${err.message}. Periksa konfigurasi dan Firestore Rules.`, 'error');
                    $('tableBody').innerHTML = '<tr><td colspan="11" class="empty">Data tidak dapat dimuat dari Firestore.</td></tr>';
                } finally {$('refreshBtn').disabled = false;}
            }
            $('downloadExcelBtn').addEventListener('click', () => {
                const shown = getVisibleRecords();
                if (!shown.length) {
                    setStatus('tableStatus', 'Tidak ada data yang cocok untuk di-download.', 'error');
                    return;
                }
                if (!window.XLSX) {
                    setStatus('tableStatus', 'Library Excel tidak berhasil dimuat. Periksa koneksi internet.', 'error');
                    return;
                }
                const exportRows = shown.map(record => {
                    const row = {};
                    fields.forEach(field => {
                        let value = record.data[field.key] ?? '';
                        if (field.type === 'date' && value) {
                            const iso = dateForInput(value);
                            if (iso) {
                                const [yyyy, mm, dd] = iso.split('-');
                                value = `${mm}/${dd}/${yyyy}`;
                            }
                        }
                        row[field.label] = value;
                    });
                    return row;
                });
                const worksheet = XLSX.utils.json_to_sheet(exportRows, {header: fields.map(field => field.label)});
                worksheet['!cols'] = [{wch: 16}, {wch: 14}, {wch: 20}, {wch: 18}, {wch: 22}, {wch: 14}, {wch: 14}, {wch: 14}, {wch: 18}, {wch: 16}];
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, 'Shipping Data');
                const searchTerm = $('searchInput').value.trim();
                const safeTerm = searchTerm.replace(/[^a-z0-9_-]/gi, '_').slice(0, 30);
                const filename = searchTerm ? `Shipping_Data_${safeTerm}.xlsx` : 'Shipping_Data.xlsx';
                XLSX.writeFile(workbook, filename);
                setStatus('tableStatus', `Berhasil menyiapkan ${shown.length} baris data untuk Excel.`, 'success');
            });
            $('searchInput').addEventListener('input', () => {currentPage = 1; renderTable();});
            $('searchColumn').addEventListener('change', () => {
                const selected = $('searchColumn').value;
                $('searchInput').placeholder = `Cari berdasarkan ${selected}…`;
                currentPage = 1;
                renderTable();
            });
            $('pageSizeSelect').addEventListener('change', () => {pageSize = Number($('pageSizeSelect').value) || 10; currentPage = 1; renderTable();});
            $('paginationButtons').addEventListener('click', (event) => {
                const button = event.target.closest('button[data-page]');
                if (!button || button.disabled) return;
                currentPage = Number(button.dataset.page) || 1;
                renderTable();
            });
            $('refreshBtn').addEventListener('click', loadData);
            loadData();
