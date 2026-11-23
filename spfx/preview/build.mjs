/**
 * Builds the standalone preview: the real web part bundled for the browser,
 * with SCSS compiled by the sass package and scoped through esbuild's
 * CSS-modules loader (mirrors what the SPFx rig does with css-loader).
 *
 *   node preview/build.mjs
 */
import * as esbuild from 'esbuild';
import sass from 'sass';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const distDir = path.join(here, 'dist');

/**
 * The component imports @microsoft/sp-http for live data calls. Those calls
 * never run in preview mode, and the package only resolves inside the real
 * SPFx runtime, so the harness swaps it for a tiny stub.
 */
const spHttpStubPlugin = {
  name: 'sp-http-stub',
  setup(build) {
    build.onResolve({ filter: /^@microsoft\/sp-http$/ }, () => ({
      path: 'sp-http-stub',
      namespace: 'sp-http-stub'
    }));
    build.onLoad({ filter: /.*/, namespace: 'sp-http-stub' }, () => ({
      contents: [
        'export const SPHttpClient = { configurations: { v1: {} } };',
        'export class SPHttpClientResponse {}',
        'export class MSGraphClientV3 {}'
      ].join('\n'),
      loader: 'js'
    }));
  }
};

/** Compile *.module.scss -> scoped CSS (esbuild "local-css" loader). */
const scssModulePlugin = {
  name: 'scss-module',
  setup(build) {
    build.onLoad({ filter: /\.module\.scss$/ }, (args) => {
      const result = sass.renderSync({ file: args.path });
      return { contents: result.css.toString('utf8'), loader: 'local-css' };
    });
  }
};

await esbuild.build({
  entryPoints: [path.join(here, 'entry.tsx')],
  bundle: true,
  format: 'iife',
  target: ['es2019'],
  jsx: 'automatic',
  outfile: path.join(distDir, 'bundle.js'),
  plugins: [spHttpStubPlugin, scssModulePlugin],
  define: { 'process.env.NODE_ENV': '"production"' },
  logLevel: 'info'
});

const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>SML Cosmos Intranet — Preview</title>
<link rel="stylesheet" href="./bundle.css" />
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0;
    background: #f2f3f7;
    font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif;
    color: #1b1010;
  }
  .spbar {
    display: flex;
    align-items: center;
    gap: 14px;
    height: 48px;
    padding: 0 20px;
    background: #ffffff;
    border-bottom: 1px solid #e3e5ec;
    position: sticky;
    top: 0;
    z-index: 10;
  }
  .spbar .suite {
    width: 26px; height: 26px; border-radius: 5px;
    background: linear-gradient(135deg, #dc2626, #86efac);
    flex: 0 0 auto;
  }
  .spbar .site { font-size: 15px; font-weight: 600; }
  .spbar .crumb { font-size: 13px; color: #6d5f5f; }
  .spbar .spacer { flex: 1 1 auto; }
  .spbar .chip {
    font-size: 12px; font-weight: 700; color: #166534;
    background: #dcfce7; border: 1px solid #86efac;
    padding: 4px 10px; border-radius: 999px;
  }
  .shell { max-width: 1180px; margin: 0 auto; padding: 26px 22px 64px; }
  .hint {
    margin-top: 26px; font-size: 12px; color: #6d5f5f; text-align: center; line-height: 1.7;
  }
  .hint code { background: #ffffff; border: 1px solid #e3e5ec; padding: 1px 6px; border-radius: 6px; }
</style>
</head>
<body>
  <div class="spbar">
    <span class="suite"></span>
    <span class="site">SML Cosmos</span>
    <span class="crumb">Home</span>
    <span class="spacer"></span>
    <span class="chip">Local preview — not deployed</span>
  </div>
  <div class="shell">
    <div id="preview-root"></div>
    <p class="hint">
      This is the real web part running locally with demo data.<br />
      Preview a video: <code>http://localhost:4399/?video=https://youtu.be/VIDEO_ID</code>
      &nbsp;·&nbsp; Preview a logo: <code>http://localhost:4399/?logo=https://.../logo.png</code>
    </p>
  </div>
  <script src="./bundle.js"></script>
</body>
</html>
`;

fs.writeFileSync(path.join(distDir, 'index.html'), html, 'utf8');
console.log(`Preview written to ${distDir}/index.html`);
