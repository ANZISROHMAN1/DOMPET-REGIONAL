const fs = require('fs');
let html = fs.readFileSync('admin.html', 'utf8');

// 1. Extract Planning Section
const planningRegex = /\s*<!-- Planning Saldo Section -->[\s\S]*?<\/div>\s*<\/div>\s*<\/div>/;
// Wait, the div closing tags:
// <div class="glass-card"...>
// ...
// </div>
// It's just one main div. Let's make sure we capture it correctly.
