const fs = require('fs');
let html = fs.readFileSync('admin.html', 'utf8');

const planningSection = `
                    <!-- Planning Saldo Section -->
                    <div class="glass-card" style="margin-bottom: 2rem;">
                        <h2 style="font-size: 1.25rem; font-weight: 600; margin-bottom: 1rem;"><i class="ri-pie-chart-2-line" style="color: var(--primary-color);"></i> Perencanaan Saldo</h2>
                        <div style="display: flex; gap: 1rem; flex-wrap: wrap; align-items: flex-end; margin-bottom: 1.5rem;">
                            <div style="flex: 1; min-width: 200px;">
                                <label style="display: block; margin-bottom: 0.5rem; font-size: 0.9rem; color: var(--text-secondary);">Nama Planning</label>
                                <input type="text" id="plan-name" class="form-control" placeholder="Contoh: Beli Laptop">
                            </div>
                            <div style="flex: 1; min-width: 150px;">
                                <label style="display: block; margin-bottom: 0.5rem; font-size: 0.9rem; color: var(--text-secondary);">Kategori</label>
                                <input type="text" id="plan-kategori" list="kategori-list" class="form-control" placeholder="Contoh: Renov, TOP">
                                <datalist id="kategori-list">
                                    <option value="Renov">
                                    <option value="TOP">
                                    <option value="Lainnya">
                                </datalist>
                            </div>
                            <div style="flex: 1; min-width: 200px;">
                                <label style="display: block; margin-bottom: 0.5rem; font-size: 0.9rem; color: var(--text-secondary);">Budget (Rp)</label>
                                <input type="number" id="plan-budget" class="form-control" placeholder="Contoh: 15000000">
                            </div>
                            <div>
                                <button onclick="submitPlanning()" id="btn-submit-plan" class="btn btn-primary"><i class="ri-add-line"></i> Tambah Planning</button>
                            </div>
                        </div>
                        
                        <div class="table-responsive">
                            <table class="data-table">
                                <thead>
                                    <tr>
                                        <th>Tanggal</th>
                                        <th>Planning</th>
                                        <th>Kategori</th>
                                        <th>Budget</th>
                                        <th style="width: 80px;">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody id="planning-body">
                                    <tr><td colspan="5" style="text-align: center; padding: 2rem; color: var(--text-muted);"><i class="ri-loader-4-line ri-spin"></i> Memuat data planning...</td></tr>
                                </tbody>
                            </table>
                        </div>
                        
                        <div style="margin-top: 2rem; padding-top: 1.5rem; border-top: 1px solid var(--border-color);">
                            <h3 style="font-size: 1.1rem; font-weight: 600; margin-bottom: 1rem;"><i class="ri-wallet-3-line" style="color: var(--primary-color);"></i> Rincian Saldo Akhir Berdasarkan Planning</h3>
                            <div id="saldo-breakdown" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 1rem;">
                                <!-- Will be filled by JS -->
                            </div>
                        </div>
                    </div>
`;

// Insert the planning section right after the dashboard-stats
if (!html.includes('<!-- Planning Saldo Section -->')) {
    html = html.replace(/(<div class="dashboard-stats"[\s\S]*?<\/div>\s*<\/div>)/, `$1\n\n${planningSection}`);
}

const planningJS = `
        let globalPlanningData = [];
        let globalSaldoAkhir = 0;

        async function fetchPlanningData() {
            try {
                const response = await fetch(\`\${SCRIPT_URL}?action=get_planning\`);
                const result = await response.json();
                if (result.success) {
                    globalPlanningData = result.data;
                    renderPlanningTable();
                    updateSaldoBreakdown();
                }
            } catch (error) {
                console.error('Error fetching planning:', error);
                document.getElementById('planning-body').innerHTML = '<tr><td colspan="5" style="text-align: center; color: var(--danger-color);">Gagal memuat planning.</td></tr>';
            }
        }

        function renderPlanningTable() {
            const tbody = document.getElementById('planning-body');
            if (globalPlanningData.length === 0) {
                tbody.innerHTML = '<tr><td colspan="5" style="text-align: center; color: var(--text-muted); padding: 2rem;">Belum ada data planning.</td></tr>';
                return;
            }
            
            let html = '';
            globalPlanningData.forEach(plan => {
                const date = new Date(plan.timestamp).toLocaleDateString('id-ID');
                html += \`
                    <tr>
                        <td>\${date}</td>
                        <td style="font-weight: 500;">\${plan.planning}</td>
                        <td><span class="badge badge-primary">\${plan.kategori}</span></td>
                        <td style="font-weight: 600;">\${formatCurrency(plan.budget)}</td>
                        <td>
                            <button onclick="deletePlanning('\${plan.id}')" class="btn btn-sm" style="background: rgba(239,68,68,0.1); color: #ef4444; border: none; padding: 0.3rem 0.6rem;">
                                <i class="ri-delete-bin-line"></i>
                            </button>
                        </td>
                    </tr>
                \`;
            });
            tbody.innerHTML = html;
        }

        async function submitPlanning() {
            const name = document.getElementById('plan-name').value.trim();
            const kategori = document.getElementById('plan-kategori').value.trim();
            const budget = parseFloat(document.getElementById('plan-budget').value);
            
            if (!name || !kategori || isNaN(budget) || budget <= 0) {
                alert('Harap isi semua field dengan benar!');
                return;
            }
            
            const btn = document.getElementById('btn-submit-plan');
            btn.innerHTML = '<i class="ri-loader-4-line ri-spin"></i> Menyimpan...';
            btn.disabled = true;
            
            try {
                const response = await fetch(SCRIPT_URL, {
                    method: 'POST',
                    body: JSON.stringify({
                        action: 'save_planning',
                        planning: name,
                        kategori: kategori,
                        budget: budget
                    })
                });
                
                const result = await response.json();
                if (result.success) {
                    document.getElementById('plan-name').value = '';
                    document.getElementById('plan-kategori').value = '';
                    document.getElementById('plan-budget').value = '';
                    await fetchPlanningData();
                } else {
                    alert('Gagal menyimpan: ' + result.message);
                }
            } catch (error) {
                alert('Terjadi kesalahan jaringan.');
            } finally {
                btn.innerHTML = '<i class="ri-add-line"></i> Tambah Planning';
                btn.disabled = false;
            }
        }
        
        async function deletePlanning(id) {
            if (!confirm('Apakah Anda yakin ingin menghapus planning ini? Saldo akhir akan dikembalikan.')) return;
            
            showLoader();
            try {
                const response = await fetch(SCRIPT_URL, {
                    method: 'POST',
                    body: JSON.stringify({
                        action: 'delete_planning',
                        id: id
                    })
                });
                
                const result = await response.json();
                if (result.success) {
                    await fetchPlanningData();
                } else {
                    alert('Gagal menghapus: ' + result.message);
                }
            } catch (error) {
                alert('Terjadi kesalahan jaringan.');
            } finally {
                hideLoader();
            }
        }

        function updateSaldoBreakdown() {
            const container = document.getElementById('saldo-breakdown');
            if (globalSaldoAkhir === 0) {
                container.innerHTML = '<div style="color: var(--text-muted);">Saldo akhir belum termuat atau bernilai 0.</div>';
                return;
            }
            
            let totalBudget = 0;
            let categories = {};
            
            globalPlanningData.forEach(plan => {
                totalBudget += plan.budget;
                if (!categories[plan.kategori]) categories[plan.kategori] = 0;
                categories[plan.kategori] += plan.budget;
            });
            
            let sisaTersedia = globalSaldoAkhir - totalBudget;
            
            let cardsHtml = \`
                <div class="stat-card" style="background: rgba(14, 165, 233, 0.05); border: 1px solid rgba(14, 165, 233, 0.2);">
                    <div style="font-size: 0.9rem; color: var(--text-secondary); margin-bottom: 0.2rem;">Total Saldo Akhir</div>
                    <div style="font-size: 1.2rem; font-weight: 700; color: #0ea5e9;">\${formatCurrency(globalSaldoAkhir)}</div>
                </div>
            \`;
            
            for (let cat in categories) {
                cardsHtml += \`
                    <div class="stat-card" style="background: rgba(245, 158, 11, 0.05); border: 1px solid rgba(245, 158, 11, 0.2);">
                        <div style="font-size: 0.9rem; color: var(--text-secondary); margin-bottom: 0.2rem;">Saldo \${cat}</div>
                        <div style="font-size: 1.2rem; font-weight: 700; color: #f59e0b;">\${formatCurrency(categories[cat])}</div>
                    </div>
                \`;
            }
            
            cardsHtml += \`
                <div class="stat-card" style="background: \${sisaTersedia >= 0 ? 'rgba(16, 185, 129, 0.05)' : 'rgba(239, 68, 68, 0.05)'}; border: 1px solid \${sisaTersedia >= 0 ? 'rgba(16, 185, 129, 0.2)' : 'rgba(239, 68, 68, 0.2)'};">
                    <div style="font-size: 0.9rem; color: var(--text-secondary); margin-bottom: 0.2rem;">Sisa Bebas/Unallocated</div>
                    <div style="font-size: 1.2rem; font-weight: 700; color: \${sisaTersedia >= 0 ? '#10b981' : '#ef4444'};">\${formatCurrency(sisaTersedia)}</div>
                </div>
            \`;
            
            container.innerHTML = cardsHtml;
        }
`;

// Insert planningJS before closing script tag
if (!html.includes('globalPlanningData')) {
    html = html.replace(/<\/script>\s*<\/body>/, `\n${planningJS}\n</script>\n</body>`);
}

// Update fetchReportData to include fetchPlanningData
const fetchReportDataRegex = /(async function fetchReportData\(\) \{[\s\S]*?await Promise\.all\(\[[\s\S]*?fetchJagoData\(\),[\s\S]*?fetchNeracaData\(\))/;
if (html.match(fetchReportDataRegex) && !html.includes('fetchPlanningData(),')) {
    html = html.replace(fetchReportDataRegex, `$1,\n                    fetchPlanningData()`);
}

// Update fetchNeracaData to store globalSaldoAkhir and call updateSaldoBreakdown
const fetchNeracaDataRegex = /(if \(document.getElementById\('stat-sisa-saldo'\)\) document.getElementById\('stat-sisa-saldo'\).innerText = formatRupiah\(saldoAkhir \|\| 0\);)/;
if (html.match(fetchNeracaDataRegex) && !html.includes('globalSaldoAkhir = saldoAkhir')) {
    html = html.replace(fetchNeracaDataRegex, `$1\n                    globalSaldoAkhir = saldoAkhir || 0;\n                    updateSaldoBreakdown();`);
}

fs.writeFileSync('admin.html', html);
console.log('admin.html updated successfully');
