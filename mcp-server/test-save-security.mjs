import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

const server = fileURLToPath(new URL('./dist/index.js', import.meta.url));

async function withServer(run) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'hwpx-save-security-'));
  const client = new Client({ name: 'save-security-test', version: '1.0.0' });
  try {
    await client.connect(new StdioClientTransport({ command: process.execPath, args: [server], cwd: directory }));
    const call = async (name, args) => {
      const result = await client.callTool({ name, arguments: args });
      assert.equal(result.isError, undefined);
      return JSON.parse(result.content[0].text);
    };
    const created = await call('create_document', { title: 'Security regression' });
    assert.ok(created.doc_id);
    await run(directory, call, created.doc_id);
  } finally {
    await client.close();
    await fs.rm(directory, { recursive: true, force: true });
  }
}

for (const suffix of ['.tmp', '.bak']) {
  test(`save does not follow an existing ${suffix} symlink`, async () => {
    await withServer(async (directory, call, docId) => {
      const output = path.join(directory, 'document.hwpx');
      const victim = path.join(directory, 'unrelated.txt');
      await fs.writeFile(victim, 'DO NOT CHANGE');
      if (suffix === '.bak') await fs.writeFile(output, 'ORIGINAL DOCUMENT');
      await fs.symlink(victim, output + suffix);
      const result = await call('save_document', { doc_id: docId, output_path: output });
      assert.equal(await fs.readFile(victim, 'utf8'), 'DO NOT CHANGE');
      if (suffix === '.bak') {
        assert.match(result.error, /Backup destination must be a regular file/);
        assert.equal(await fs.readFile(output, 'utf8'), 'ORIGINAL DOCUMENT');
      } else {
        assert.equal(result.error, undefined);
        const opened = await call('open_document', { file_path: output });
        assert.ok(opened.doc_id);
      }
      assert.equal((await fs.lstat(output + suffix)).isSymbolicLink(), true);
      assert.equal((await fs.readdir(directory)).some(name => name.startsWith('.hwpx-save-')), false);
    });
  });
}

test('repeated saves preserve a regular backup and produce readable documents', async () => {
  await withServer(async (directory, call, docId) => {
    const output = path.join(directory, 'document.hwpx');
    for (let attempt = 0; attempt < 3; attempt++) {
      const before = attempt ? await fs.readFile(output) : null;
      const result = await call('save_document', { doc_id: docId, output_path: output });
      assert.equal(result.error, undefined);
      if (before) assert.deepEqual(await fs.readFile(output + '.bak'), before);
      const opened = await call('open_document', { file_path: output });
      assert.ok(opened.doc_id);
    }
  });
});

test('failed rename preserves the existing destination and removes staging files', async () => {
  await withServer(async (directory, call, docId) => {
    const output = path.join(directory, 'existing-directory');
    await fs.mkdir(output);
    await fs.writeFile(path.join(output, 'original.txt'), 'ORIGINAL');
    const result = await call('save_document', {
      doc_id: docId, output_path: output, create_backup: false,
    });
    assert.match(result.error, /Save failed; original document preserved/);
    assert.equal(await fs.readFile(path.join(output, 'original.txt'), 'utf8'), 'ORIGINAL');
    assert.equal((await fs.readdir(directory)).some(name => name.startsWith('.hwpx-save-')), false);
  });
});

test('backup replacement does not overwrite a hard-linked unrelated file', async () => {
  await withServer(async (directory, call, docId) => {
    const output = path.join(directory, 'document.hwpx');
    const victim = path.join(directory, 'unrelated.txt');
    await fs.writeFile(output, 'ORIGINAL DOCUMENT');
    await fs.writeFile(victim, 'DO NOT CHANGE');
    await fs.link(victim, output + '.bak');
    const result = await call('save_document', { doc_id: docId, output_path: output });
    assert.equal(result.error, undefined);
    assert.equal(await fs.readFile(victim, 'utf8'), 'DO NOT CHANGE');
    assert.equal(await fs.readFile(output + '.bak', 'utf8'), 'ORIGINAL DOCUMENT');
  });
});

test('save preserves a pre-existing regular .tmp file', async () => {
  await withServer(async (directory, call, docId) => {
    const output = path.join(directory, 'document.hwpx');
    await fs.writeFile(output + '.tmp', 'UNRELATED TEMP FILE');
    const result = await call('save_document', { doc_id: docId, output_path: output });
    assert.equal(result.error, undefined);
    assert.equal(await fs.readFile(output + '.tmp', 'utf8'), 'UNRELATED TEMP FILE');
  });
});

for (const useRelativePath of [false, true]) {
  test(`existing ${useRelativePath ? 'relative' : 'absolute'} paths outside the working directory remain usable`, async () => {
    await withServer(async (directory, call, docId) => {
      const external = await fs.mkdtemp(path.join(os.tmpdir(), 'hwpx-user-documents-'));
      try {
        const output = path.join(external, '사업계획서 최종본.hwpx');
        const requestedPath = useRelativePath ? path.relative(directory, output) : output;
        const saved = await call('save_document', { doc_id: docId, output_path: requestedPath });
        assert.equal(saved.error, undefined);
        const opened = await call('open_document', { file_path: requestedPath });
        assert.ok(opened.doc_id);
        const resaved = await call('save_document', { doc_id: opened.doc_id });
        assert.equal(resaved.error, undefined);
        assert.ok((await fs.stat(output + '.bak')).size > 0);
        for (const tool of ['export_to_text', 'export_to_html']) {
          const exportedPath = path.join(external, tool);
          const exported = await call(tool, {
            doc_id: opened.doc_id,
            output_path: useRelativePath ? path.relative(directory, exportedPath) : exportedPath,
          });
          assert.equal(exported.error, undefined);
          assert.equal((await fs.stat(exportedPath)).isFile(), true);
        }
      } finally {
        await fs.rm(external, { recursive: true, force: true });
      }
    });
  });
}
