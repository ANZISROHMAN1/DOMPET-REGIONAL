const fs = require('fs');
let html = fs.readFileSync('admin.html', 'utf8');

const sectionRegex = /\s*<!-- Planning Saldo Section -->[\s\S]*?<div style="margin-top: 2rem; padding-top: 1\.5rem; border-top: 1px solid var\(--border-color\);">[\s\S]*?<\/div>\s*<\/div>\s*<\/div>\s*/;

let match = html.match(sectionRegex);
if (match) {
    let sectionCode = match[0];
    html = html.replace(sectionRegex, '\n\n'); // Remove from old location
    
    // Find AI Insight end and insert after it
    html = html.replace(/(<\/div>\s*<!-- Overview Cards -->)/, `\n\n${sectionCode}\n$1`);
    
    fs.writeFileSync('admin.html', html);
    console.log("Moved successfully.");
} else {
    console.log("Could not find section with regex.");
}
