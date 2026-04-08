import { spawnSync } from 'node:child_process';
import { existsSync, statSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';

const require = createRequire(import.meta.url);

const packageJsonPath = require.resolve('date-holidays/package.json');
const packageDir = path.dirname(packageJsonPath);
const holidaysJsonPath = path.join(packageDir, 'data', 'holidays.json');
const shrinkScriptPath = path.join(packageDir, 'scripts', 'holidays2json.cjs');

if (!existsSync(holidaysJsonPath) || !existsSync(shrinkScriptPath)) {
  console.warn('[shrink-date-holidays] date-holidays package files were not found. Skipping.');
  process.exit(0);
}

const originalSize = statSync(holidaysJsonPath).size;
const alreadySlimmed = originalSize < 100_000;

if (alreadySlimmed) {
  console.log(`[shrink-date-holidays] already slimmed (${originalSize} bytes).`);
  process.exit(0);
}

const result = spawnSync(process.execPath, [shrinkScriptPath, '--pick', 'KR', '--min'], {
  cwd: packageDir,
  stdio: 'inherit',
});

if (result.status !== 0) {
  process.exit(result.status ?? 1);
}

const shrunkSize = statSync(holidaysJsonPath).size;
console.log(`[shrink-date-holidays] reduced holidays.json from ${originalSize} to ${shrunkSize} bytes.`);
