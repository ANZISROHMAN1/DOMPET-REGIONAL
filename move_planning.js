const fs = require('fs');
let html = fs.readFileSync('admin.html', 'utf8');

const lines = html.split('\n');

// Find the start and end of the Planning Section
let startIdx = -1;
let endIdx = -1;
for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes('<!-- Planning Saldo Section -->')) {
        startIdx = i;
    }
    if (startIdx !== -1 && i > startIdx && lines[i].includes('<!-- End of dashboard-view -->')) {
        // We know it ends exactly at the div before it, but let's just search for the exact end div of planning
        // The planning section is 50 lines long.
        break;
    }
}

// Since I have the exact line numbers from view_file (160 to 209 in 1-based index)
// Wait, the line numbers might shift if the file was modified, but I just checked it.
// Let's do it safely by string extraction
const startStr = '                    <!-- Planning Saldo Section -->';
const endStr = '                    </div>'; // This closes the glass-card of Planning Saldo
// But there are many </div>. Let's extract exactly using line numbers 159-209.

const planningLines = lines.slice(159, 210); // 160 to 210 is index 159 to 209 (51 lines)
// Let's verify
if (planningLines[0].includes('<!-- Planning Saldo Section -->') && planningLines[50].includes('</div>')) {
    console.log("Found planning lines correctly.");
    
    // Remove from original position
    lines.splice(159, 51);
    
    // Now find the Neraca Keuangan section to insert after
    let insertIdx = -1;
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('<!-- End of report-view -->')) {
            insertIdx = i;
            break;
        }
    }
    
    if (insertIdx !== -1) {
        // Insert right before <!-- End of report-view -->
        lines.splice(insertIdx, 0, ...planningLines);
        fs.writeFileSync('admin.html', lines.join('\n'));
        console.log("Moved successfully.");
    } else {
        console.log("Could not find insert index.");
    }
} else {
    console.log("Mismatch in planning lines extraction:", planningLines[0], planningLines[50]);
}

