// Bundles the Outlook add-in and the preview page, and writes the add-in
// manifest for the URL in src/config.js. The output is committed so GitHub
// Pages can serve it without a build step.
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { build } from 'esbuild';

const { baseUrl } = createRequire(import.meta.url)('./src/config.js');
const manifest = readFileSync('src/manifest.xml', 'utf8')
  .replaceAll('{{baseUrl}}', baseUrl)
  .replaceAll('{{origin}}', new URL(baseUrl).origin);
writeFileSync('addin/manifest.xml', manifest);

const common = { bundle: true, minify: true, legalComments: 'none', logLevel: 'info', format: 'iife' };

await Promise.all([
  // Classic Outlook's JavaScript-only runtime needs one plain script.
  build({ ...common, entryPoints: ['src/launchevent.js'], outfile: 'addin/launchevent.js', target: 'es2017' }),
  build({ ...common, entryPoints: ['src/preview.js'], outfile: 'preview.js', target: 'es2020' })
]);
