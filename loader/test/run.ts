// Runs every test: each test/*.test.ts file, then the real-file cases in test/cases/. Usage: npm test
import { spawnSync } from 'child_process';
import { readdirSync } from 'fs';

const dir = new URL('.', import.meta.url).pathname;
const files = [
  ...readdirSync(dir)
    .filter(f => f.endsWith('.test.ts'))
    .sort(),
  'cases.ts',
];
const failed: string[] = [];
for (const f of files) {
  const t = Date.now();
  const r = spawnSync(process.execPath, ['--import', 'tsx', dir + f], { encoding: 'utf8' });
  const secs = ((Date.now() - t) / 1000).toFixed(1);
  if (r.status === 0) {
    console.log(`✓ ${f} (${secs}s)`);
    if (f === 'cases.ts')
      console.log(
        r.stdout
          .trim()
          .split('\n')
          .map(l => '    ' + l)
          .join('\n'),
      );
  } else {
    failed.push(f);
    console.log(
      `✗ ${f} (${secs}s)\n${(r.stdout + r.stderr)
        .trim()
        .split('\n')
        .slice(-40)
        .map(l => '    ' + l)
        .join('\n')}`,
    );
  }
}
console.log(
  failed.length
    ? `\n${failed.length} of ${files.length} test files failed: ${failed.join(', ')}`
    : `\nAll ${files.length} test files passed.`,
);
process.exit(failed.length ? 1 : 0);
