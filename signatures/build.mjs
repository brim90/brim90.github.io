// Bundles the Outlook add-in and the preview page. The output is committed so
// GitHub Pages can serve it without a build step.
import { build } from 'esbuild';

const common = { bundle: true, minify: true, legalComments: 'none', logLevel: 'info', format: 'iife' };

await Promise.all([
  // Classic Outlook's JavaScript-only runtime needs one plain script.
  build({ ...common, entryPoints: ['src/launchevent.js'], outfile: 'addin/launchevent.js', target: 'es2017' }),
  build({ ...common, entryPoints: ['src/preview.js'], outfile: 'preview.js', target: 'es2020' })
]);
