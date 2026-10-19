// Bundles the admin console, the sign-in redirect page and the Outlook add-in.
// Output is committed so GitHub Pages can serve it without a build step.
import { build } from 'esbuild';

const common = { bundle: true, minify: true, sourcemap: false, legalComments: 'none', logLevel: 'info' };

await Promise.all([
  build({ ...common, entryPoints: ['src/admin.js'], outfile: 'dist/admin.js', format: 'iife', target: 'es2020' }),
  build({ ...common, entryPoints: ['src/redirect.js'], outfile: 'dist/redirect.js', format: 'iife', target: 'es2020' }),
  // Classic Outlook's JavaScript-only runtime needs a single plain script.
  build({ ...common, entryPoints: ['src/launchevent.js'], outfile: 'addin/launchevent.js', format: 'iife', target: 'es2017' })
]);
