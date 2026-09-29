const fs = require('fs');
let html = fs.readFileSync('index.html', 'utf-8');

// 1. Inject nav button
const navAnchor = `id="nav-pengaturan"><i`;
const navInsert = `<button class="nav-link-custom" onclick="switchTab('point')" id="nav-point"><i class="bi bi-star-fill"></i>Point</button>\n                `;
html = html.replace(navAnchor, navInsert + '<button class="nav-link-custom" onclick="switchTab(\'pengaturan\')" ' + navAnchor);

// 2. Inject switchTab logic
const switchAnchor = `if (t === 'jurnal') loadJurnalMenu();`;
const switchInsert = `if (t === 'point') loadPointMenu();\n        `;
html = html.replace(switchAnchor, switchAnchor + '\n        ' + switchInsert);

// 3. Inject tab pane (Let's insert it before <div class="tab-pane fade" id="pills-pengaturan">)
const paneAnchor = `<div class="tab-pane fade" id="pills-pengaturan">`;
const paneInsert = `
        <!-- TAB POINT REWARD & PELANGGARAN -->
        <div class="tab-pane fade" id="pills-point">
            <div class="d-flex justify-content-between align-items-center mb-3">
                <h5 class="fw-bold m-0"><i class="bi bi-star-fill text-warning me-2"></i>Data Prestasi & Pelanggaran</h5>
                <button class="btn btn-primary btn-sm rounded-pill px-3 shadow-sm" onclick="showPointModal()">
                    <i class="bi bi-plus-lg me-1"></i> Input Poin
                </button>
            </div>
            
            <div class="card p-3 shadow-sm border-0 bg-white">
                <div class="table-responsive">
                    <table class="table table-hover align-middle table-nowrap custom-table" id="point-table">
                        <thead class="table-light text-secondary">
                            <tr>
                                <th>Tanggal</th>
                                <th>Siswa</th>
                                <th>Kelas</th>
                                <th>Tipe</th>
                                <th>Poin</th>
                                <th>Aksi</th>
                            </tr>
                        </thead>
                        <tbody id="point-tbody">
                            <tr><td colspan="6" class="text-center text-muted py-4"><div class="spinner-border text-primary spinner-border-sm" role="status"></div>  Memuat data...</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
`;
html = html.replace(paneAnchor, paneInsert + '\n        ' + paneAnchor);

// 4. Inject Modals (after jurnal modal)
const modalAnchor = `<!-- Modal Presensi Siswa -->`;
const modalInsert = `
<!-- Modal Input Point -->
<div class="modal fade" id="pointModal" tabindex="-1">
    <div class="modal-dialog modal-dialog-centered">
        <div class="modal-content rounded-4 border-0 shadow">
            <div class="modal-header bg-primary text-white rounded-top-4">
                <h5 class="modal-title fw-bold"><i class="bi bi-star-fill me-2"></i>Input Poin Siswa</h5>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
            </div>
            <div class="modal-body p-4">
                <form id="pointForm">
                    <input type="hidden" id="point-id">
                    
                    <div class="mb-3">
                        <label class="form-label text-secondary small fw-bold">Tanggal</label>
                        <input type="date" class="form-control rounded-3" id="point-tgl" required>
                    </div>

                    <div class="mb-3">
                        <label class="form-label text-secondary small fw-bold">Kelas</label>
                        <select class="form-select rounded-3" id="point-kelas" required onchange="populatePointSiswa()"></select>
                    </div>

                    <div class="mb-3">
                        <label class="form-label text-secondary small fw-bold">Nama Siswa</label>
                        <select class="form-select rounded-3" id="point-nama" required></select>
                    </div>

                    <div class="mb-3">
                        <label class="form-label text-secondary small fw-bold">Tipe Poin</label>
                        <select class="form-select rounded-3" id="point-tipe" required onchange="populatePointKategori()">
                            <option value="">-- Pilih Tipe --</option>
                            <option value="Prestasi">Prestasi (Reward)</option>
                            <option value="Pelanggaran">Pelanggaran (Punishment)</option>
                        </select>
                    </div>

                    <div class="mb-3">
                        <label class="form-label text-secondary small fw-bold">Kategori / Ketentuan</label>
                        <select class="form-select rounded-3" id="point-kategori" required></select>
                    </div>
                </form>
            </div>
            <div class="modal-footer border-0">
                <button type="button" class="btn btn-light rounded-pill px-4" data-bs-dismiss="modal">Batal</button>
                <button type="button" class="btn btn-primary rounded-pill px-4 fw-bold" onclick="simpanPoint()"><i class="bi bi-save me-1"></i>Simpan</button>
            </div>
        </div>
    </div>
</div>
`;
html = html.replace(modalAnchor, modalInsert + '\n' + modalAnchor);

// 5. Inject Logic via JS
// Add JS at the end before closing script tag
const jsAnchor = `// ================== PENGATURAN SEKOLAH ==================`;
const jsInsert = `
// ================== POINT REWARD & PELANGGARAN ==================
let pointData = [];
const pelanggaranList = [
    { nama: 'Terlambat Datang', poin: -5 },
    { nama: 'Atribut Tidak Lengkap', poin: -5 },
    { nama: 'Membawa HP Tanpa Izin', poin: -10 },
    { nama: 'Bolos Sekolah', poin: -20 },
    { nama: 'Berkelahi', poin: -50 }
];
const prestasiList = [
    { nama: 'Juara Kelas', poin: 20 },
    { nama: 'Juara Lomba Tingkat Sekolah', poin: 30 },
    { nama: 'Juara Lomba Tingkat Kota/Kab', poin: 50 },
    { nama: 'Mengharumkan Nama Sekolah', poin: 100 }
];

async function loadPointMenu() {
    try {
        const res = await callAPI('getPoint', {});
        if(res && res.status === 'success'){
            pointData = res.data;
            renderPointTable(pointData);
        } else throw new Error(res?.message || 'Error muat Point');
    } catch(err) {
        Swal.fire('Error', err.message, 'error');
    }
}

function renderPointTable(data) {
    const tbody = document.getElementById('point-tbody');
    if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted">Belum ada data prestasi/pelanggaran.</td></tr>';
        return;
    }
    tbody.innerHTML = '';
    data.reverse().forEach(d => {
        const badgeClass = d.tipe === 'Prestasi' ? 'bg-success' : 'bg-danger';
        const poinColor = d.tipe === 'Prestasi' ? 'text-success fw-bold' : 'text-danger fw-bold';
        
        tbody.innerHTML += \`
            <tr>
                <td class="text-secondary" style="font-size:0.9rem;">\${d.tanggal}</td>
                <td class="fw-semibold">\${d.nama}</td>
                <td><span class="badge bg-secondary rounded-pill">\${d.kelas}</span></td>
                <td><span class="badge \${badgeClass}">\${d.tipe}</span></td>
                <td class="\${poinColor}">\${d.poin}</td>
                <td>
                    <button class="btn btn-sm btn-outline-danger rounded-pill" onclick="hapusPoint('\${d.id}')">
                        <i class="bi bi-trash"></i>
                    </button>
                </td>
            </tr>
        \`;
    });
}

function showPointModal() {
    document.getElementById('pointForm').reset();
    document.getElementById('point-id').value = '';
    document.getElementById('point-tgl').valueAsDate = new Date();
    
    // Populate Kelas
    const kSelect = document.getElementById('point-kelas');
    kSelect.innerHTML = '<option value="">-- Pilih Kelas --</option>';
    let availableClasses = [...new Set(siswaData.map(s => s.kelas.trim()))].sort();
    availableClasses.forEach(c => kSelect.innerHTML += \`<option value="\${c}">\${c}</option>\`);
    
    document.getElementById('point-nama').innerHTML = '<option value="">-- Pilih Siswa --</option>';
    document.getElementById('point-kategori').innerHTML = '<option value="">-- Kategori --</option>';
    
    new bootstrap.Modal(document.getElementById('pointModal')).show();
}

function populatePointSiswa() {
    const kelas = document.getElementById('point-kelas').value;
    const nSelect = document.getElementById('point-nama');
    nSelect.innerHTML = '<option value="">-- Pilih Siswa --</option>';
    
    if(!kelas) return;
    
    let filtered = siswaData.filter(s => s.kelas.trim() === kelas).sort((a,b) => a.nama.localeCompare(b.nama));
    filtered.forEach(s => {
        nSelect.innerHTML += \`<option value="\${s.nama}">\${s.nama}</option>\`;
    });
}

function populatePointKategori() {
    const tipe = document.getElementById('point-tipe').value;
    const catSelect = document.getElementById('point-kategori');
    catSelect.innerHTML = '<option value="">-- Pilih Kategori --</option>';
    
    let list = tipe === 'Prestasi' ? prestasiList : (tipe === 'Pelanggaran' ? pelanggaranList : []);
    
    list.forEach(k => {
        catSelect.innerHTML += \`<option value="\${k.nama} (\${k.poin})">\${k.nama} (Poin: \${k.poin})</option>\`;
    });
}

async function simpanPoint() {
    const form = document.getElementById('pointForm');
    if (!form.checkValidity()) { form.reportValidity(); return; }

    const id = document.getElementById('point-id').value;
    const tgl = document.getElementById('point-tgl').value;
    const kls = document.getElementById('point-kelas').value;
    const nama = document.getElementById('point-nama').value;
    const tipe = document.getElementById('point-tipe').value;
    const katVal = document.getElementById('point-kategori').value;
    
    if(!tgl || !kls || !nama || !tipe || !katVal) {
        Swal.fire('Peringatan', 'Harap lengkapi semua form!', 'warning');
        return;
    }

    // Extract Point Info from the category selection string "Category Name (-10)"
    let poin = 0;
    let desc = katVal;
    let match = katVal.match(/\(([-]?[0-9]+)\)/);
    if(match) {
        poin = parseFloat(match[1]);
        desc = katVal.replace(/\(.*\)/, '').trim();
    }
    
    let pLoad = { action: 'simpanPoint', id: id, tanggal: tgl, kelas: kls, nama: nama, tipe: tipe, kategori: desc, poin: poin };
    
    Swal.fire({title: 'Menyimpan...', allowOutsideClick: false});
    Swal.showLoading();
    
    try {
        const res = await callAPI('simpanPoint', pLoad);
        if(res && res.status === 'success'){
            bootstrap.Modal.getInstance(document.getElementById('pointModal'))?.hide();
            Swal.fire('Sukses', 'Data poin berhasil disimpan!', 'success');
            loadPointMenu();
        } else throw new Error(res?.message || 'Gagal menyimpan.');
    } catch (e) {
        Swal.fire('Error', e.message, 'error');
    }
}

async function hapusPoint(id) {
    Swal.fire({
        title: 'Hapus data poin ini?',
        icon: 'warning',
        showCancelButton: true,
        confirmButtonText: 'Ya, Hapus!'
    }).then(async (result) => {
        if (result.isConfirmed) {
            Swal.fire({title: 'Menghapus...', allowOutsideClick: false}); Swal.showLoading();
            try {
                let res = await callAPI('hapusPoint', { action: 'hapusPoint', id: id });
                if(res && res.status === 'success'){
                    Swal.fire('Terhapus', 'Data poin berhasil dihapus.', 'success');
                    loadPointMenu();
                } else throw new Error(res?.message || "Gagal Hapus");
            } catch(e) { Swal.fire('Error', e.message, 'error'); }
        }
    });
}
`;
html = html.replace(jsAnchor, jsInsert + '\n\n' + jsAnchor);

fs.writeFileSync('index.html', html, 'utf-8');
console.log('index.html patched!');
