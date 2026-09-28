const fs = require('fs');
let html = fs.readFileSync('admin.html', 'utf8');

// 1. Replace the 4th stat-card
const oldCardRegex = /<div class="stat-card">\s*<div class="stat-icon purple"><i class="ri-wallet-3-line"><\/i><\/div>\s*<div class="stat-details">\s*<h3 id="stat-sisa-saldo">Rp 0<\/h3>\s*<p>Saldo Akhir<\/p>\s*<\/div>\s*<\/div>/;

const newCard = `<div class="stat-card" style="padding: 0; display: flex; flex-direction: column; overflow: hidden; background: white; border: 1px solid var(--border-color); border-radius: 12px; grid-column: span 1;">
                            <!-- Top box: Sisa Saldo -->
                            <div style="padding: 1.2rem; text-align: center; border-bottom: 1px solid var(--border-color); background: rgba(14, 165, 233, 0.05);">
                                <p style="margin: 0 0 0.3rem 0; font-size: 0.9rem; color: var(--text-secondary);">Sisa Saldo</p>
                                <h3 id="stat-sisa-saldo" style="margin: 0; font-size: 1.4rem; color: #0ea5e9; font-weight: 700;">Rp 0</h3>
                            </div>
                            
                            <!-- Middle box: Saldo exclude -->
                            <div style="padding: 0.8rem; text-align: center; border-bottom: 1px solid var(--border-color); background: rgba(16, 185, 129, 0.05);">
                                <p style="margin: 0 0 0.2rem 0; font-size: 0.8rem; color: var(--text-secondary);">Saldo exclude TOP dan Renov</p>
                                <h4 id="stat-saldo-exclude" style="margin: 0; font-size: 1.1rem; color: #10b981; font-weight: 700;">Rp 0</h4>
                            </div>
                            
                            <!-- Bottom boxes: TOP and Renov -->
                            <div style="display: flex; background: rgba(245, 158, 11, 0.05);">
                                <div style="flex: 1; padding: 0.8rem; text-align: center; border-right: 1px solid var(--border-color);">
                                    <p style="margin: 0 0 0.2rem 0; font-size: 0.75rem; color: var(--text-secondary);">Saldo TOP</p>
                                    <h4 id="stat-saldo-top" style="margin: 0; font-size: 1rem; color: #f59e0b; font-weight: 700;">Rp 0</h4>
                                </div>
                                <div style="flex: 1; padding: 0.8rem; text-align: center;">
                                    <p style="margin: 0 0 0.2rem 0; font-size: 0.75rem; color: var(--text-secondary);">Saldo Renov</p>
                                    <h4 id="stat-saldo-renov" style="margin: 0; font-size: 1rem; color: #f59e0b; font-weight: 700;">Rp 0</h4>
                                </div>
                            </div>
                        </div>`;

html = html.replace(oldCardRegex, newCard);

// 2. Remove the old "Rincian Saldo Akhir" from the Planning section
const oldRincianRegex = /<div style="margin-top: 2rem; padding-top: 1\.5rem; border-top: 1px solid var\(--border-color\);">\s*<h3 style="font-size: 1\.1rem; font-weight: 600; margin-bottom: 1rem;"><i class="ri-wallet-3-line" style="color: var\(--primary-color\);"><\/i> Rincian Saldo Akhir Berdasarkan Planning<\/h3>\s*<div id="saldo-breakdown" style="display: grid; grid-template-columns: repeat\(auto-fit, minmax\(220px, 1fr\)\); gap: 1rem;">\s*<!-- Will be filled by JS -->\s*<\/div>\s*<\/div>/;
html = html.replace(oldRincianRegex, '');

// 3. Update the updateSaldoBreakdown JS function
const oldJSFunctionRegex = /function updateSaldoBreakdown\(\) \{[\s\S]*?container\.innerHTML = cardsHtml;\s*\}/;
const newJSFunction = `function updateSaldoBreakdown() {
            if (globalSaldoAkhir === 0) return;
            
            let totalBudget = 0;
            let saldoTop = 0;
            let saldoRenov = 0;
            
            globalPlanningData.forEach(plan => {
                totalBudget += plan.budget;
                const kat = plan.kategori.toLowerCase();
                if (kat.includes('top')) saldoTop += plan.budget;
                else if (kat.includes('renov')) saldoRenov += plan.budget;
            });
            
            let sisaTersedia = globalSaldoAkhir - totalBudget;
            
            if(document.getElementById('stat-saldo-exclude')) {
                const elExclude = document.getElementById('stat-saldo-exclude');
                elExclude.innerText = formatRupiah(sisaTersedia);
                if(sisaTersedia < 0) {
                    elExclude.style.color = '#ef4444';
                    elExclude.parentElement.style.background = 'rgba(239, 68, 68, 0.05)';
                } else {
                    elExclude.style.color = '#10b981';
                    elExclude.parentElement.style.background = 'rgba(16, 185, 129, 0.05)';
                }
            }
            if(document.getElementById('stat-saldo-top')) {
                document.getElementById('stat-saldo-top').innerText = formatRupiah(saldoTop);
            }
            if(document.getElementById('stat-saldo-renov')) {
                document.getElementById('stat-saldo-renov').innerText = formatRupiah(saldoRenov);
            }
        }`;

html = html.replace(oldJSFunctionRegex, newJSFunction);

fs.writeFileSync('admin.html', html);
console.log("Updated UI for Saldo Akhir");
