const fs = require('fs');
let html = fs.readFileSync('admin.html', 'utf8');

// The planning section starts with: <!-- Planning Saldo Section -->
// And ends with: </div> exactly before <!-- End of dashboard-view --> OR before whatever comes next.
// Wait, let's just extract it exactly.
const startTag = '<!-- Planning Saldo Section -->';
const endTag = '</div>\n\n                    \n                    <div class="glass-card" style="margin-bottom: 2rem;">'; 
// That's too risky. 

const sectionRegex = /\s*<!-- Planning Saldo Section -->[\s\S]*?<div style="margin-top: 2rem; padding-top: 1\.5rem; border-top: 1px solid var\(--border-color\);">[\s\S]*?<\/div>\s*<\/div>\s*<\/div>\s*/;

let match = html.match(sectionRegex);
if (match) {
    let sectionCode = match[0];
    html = html.replace(sectionRegex, '\n'); // Remove from old location
    
    // Insert before <!-- End of report-view -->
    html = html.replace(/\s*(<\/div>\s*<!-- End of report-view -->)/, `\n\n${sectionCode}\n$1`);
    
    fs.writeFileSync('admin.html', html);
    console.log("Moved successfully.");
} else {
    console.log("Could not find section with regex.");
}
