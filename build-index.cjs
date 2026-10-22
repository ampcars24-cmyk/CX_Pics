const fs = require('node:fs');
const path = require('node:path');

const here = __dirname;
const sourcePath = path.join(here, '..', 'review-router', 'index.html');
const templatePath = path.join(here, 'index.template.html');
const outputPath = path.join(here, 'index.html');
const source = fs.readFileSync(sourcePath, 'utf8');
const match = source.match(/window\.CARS24_LOCATIONS\s*=\s*(\[[\s\S]*?\])\s*;\s*<\/script>/);
if (!match) throw new Error('Could not find the embedded location list in review-router/index.html');
const locations = JSON.parse(match[1]);
const buyHubs = locations.filter(h => h.code === 'C1' && h.intent === 'buy' && Number.isFinite(h.lat) && Number.isFinite(h.lng))
  .map(({ storeCode, code, intent, name, address, city, state, lat, lng }) => ({ storeCode, code, intent, name, address, city, state, lat, lng }));
if (!buyHubs.length) throw new Error('No valid C1 buy hubs with coordinates were found.');
const template = fs.readFileSync(templatePath, 'utf8');
if (!template.includes('/*__BUY_HUBS__*/')) throw new Error('Buy hub data placeholder is missing.');
fs.writeFileSync(outputPath, template.replace('/*__BUY_HUBS__*/', JSON.stringify(buyHubs)), 'utf8');
console.log(`Wrote ${path.relative(process.cwd(), outputPath)} with ${buyHubs.length} C1 buy hubs.`);
