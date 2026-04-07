const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

const skillsDir = path.join(__dirname, '..', 'skills');
const dirs = fs.readdirSync(skillsDir).filter(f => fs.statSync(path.join(skillsDir, f)).isDirectory());

console.log('=========================================');
console.log('   Agent Skills Specification Validator   ');
console.log('=========================================\n');

let failedCount = 0;
let passedCount = 0;

for (const dir of dirs) {
    process.stdout.write(`Validating: ${dir}... `);
    try {
        execSync(`npx --yes skills-ref validate "./skills/${dir}"`, { 
            stdio: 'pipe', 
            cwd: path.join(__dirname, '..') 
        });
        console.log('✅ PASS');
        passedCount++;
    } catch (e) {
        console.log('❌ FAIL');
        console.error(`\x1b[31mError in ${dir}:\x1b[0m\n${e.stderr || e.stdout || e.message}`);
        failedCount++;
    }
}

console.log('\n-----------------------------------------');
console.log(`Summary: ${passedCount} passed, ${failedCount} failed.`);
console.log('-----------------------------------------');

if (failedCount > 0) {
    process.exit(1);
} else {
    process.exit(0);
}
