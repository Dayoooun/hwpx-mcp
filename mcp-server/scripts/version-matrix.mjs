#!/usr/bin/env node
/**
 * 버전 매트릭스 — 같은 종단간 시나리오를 여러 서버 버전에 돌려 표로 낸다.
 *
 *   npm run test:versions                 # 기본: 게시된 0.3.0, 0.3.3 + local
 *   npm run test:versions -- 0.3.3 local  # 원하는 것만
 *
 * 쓰임:
 *   - 리뷰어가 "0.3.0 에서 5개, 0.3.3 에서 10개 통과"라고 알려 주면, 같은 표를
 *     우리 쪽에서 만들어 대조한다.
 *   - 새 버전을 내기 전에 local 이 게시된 최신 버전보다 나빠진 칸이 없는지 본다
 *     (나빠진 칸이 있으면 exit 1).
 *
 * 각 버전은 tests/e2e 를 HWPX_MCP_SERVER=<target> 으로 한 번 돌린다.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const args = process.argv.slice(2);
const targets = (args.length ? args : ['0.3.0', '0.3.3', 'local'])
  .map(t => (t === 'local' || t.startsWith('npm:') ? t : `npm:${t}`));

if (targets.includes('local')) {
  const b = spawnSync('npm', ['run', 'build'], { stdio: 'inherit' });
  if (b.status !== 0) process.exit(b.status ?? 1);
}

const results = new Map(); // test title -> { target: 'pass' | 'fail' }
for (const target of targets) {
  const report = path.join(os.tmpdir(), `hwpx-matrix-${target.replace(/[^a-z0-9.]/gi, '_')}.json`);
  const r = spawnSync('npx', ['vitest', 'run', 'tests/e2e', '--reporter=json', `--outputFile=${report}`], {
    env: { ...process.env, HWPX_MCP_SERVER: target },
    stdio: ['ignore', 'ignore', 'inherit'],
  });
  if (!fs.existsSync(report)) {
    console.error(`[${target}] no report (exit ${r.status})`);
    continue;
  }
  const json = JSON.parse(fs.readFileSync(report, 'utf8'));
  for (const file of json.testResults) {
    for (const t of file.assertionResults) {
      const title = t.title;
      if (!results.has(title)) results.set(title, {});
      results.get(title)[target] = t.status === 'passed' ? 'pass' : 'fail';
    }
  }
}

const col = Math.max(...[...results.keys()].map(k => k.length), 10);
const header = ['시나리오'.padEnd(col), ...targets.map(t => t.replace('npm:', '').padStart(7))].join(' | ');
console.log('\n' + header);
console.log('-'.repeat(header.length));
const passCount = Object.fromEntries(targets.map(t => [t, 0]));
for (const [title, row] of results) {
  console.log([title.padEnd(col), ...targets.map(t => {
    if (row[t] === 'pass') passCount[t]++;
    return (row[t] === 'pass' ? 'O' : row[t] === 'fail' ? 'X' : '-').padStart(7);
  })].join(' | '));
}
console.log('-'.repeat(header.length));
console.log(['통과'.padEnd(col), ...targets.map(t => `${passCount[t]}/${results.size}`.padStart(7))].join(' | '));

// local 이 비교 대상 중 가장 최근 게시 버전보다 나빠진 칸이 있으면 실패.
const published = targets.filter(t => t.startsWith('npm:'));
const latest = published[published.length - 1];
if (targets.includes('local') && latest) {
  const worse = [...results].filter(([, row]) => row[latest] === 'pass' && row.local !== 'pass').map(([t]) => t);
  if (worse.length) {
    console.error(`\nlocal 이 ${latest.slice(4)} 보다 나빠진 시나리오:\n  ${worse.join('\n  ')}`);
    process.exit(1);
  }
}
