// Edição de uma lista de referência: termos de texto ou expressões regulares, importação e teste.
import { get, post, put } from '../api.js';
import { html, render as paint, icon, openDialog, confirmDialog, toast, fmtNum, plural, debounce } from '../ui.js';
import { go } from '../nav.js';

const PAGE_SIZE = 100;

/** Mesma normalização do servidor: sem acentos, minúsculas e espaços simples. */
function fold(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

const termKey = (t) => `${t.type}|${t.type === 'text' ? fold(t.value) : t.value}|${t.validator || ''}`;

function firstCsvCell(line, sep) {
  if (!line.startsWith('"')) {
    const idx = line.indexOf(sep);
    return idx === -1 ? line : line.slice(0, idx);
  }
  let out = '';
  for (let i = 1; i < line.length; i++) {
    if (line[i] === '"') {
      if (line[i + 1] === '"') {
        out += '"';
        i++;
      } else break;
    } else out += line[i];
  }
  return out;
}

/** Lê um .txt (um termo por linha) ou .csv (primeira coluna), em UTF-8 ou Windows-1252. */
async function readTermsFile(file) {
  const buf = await file.arrayBuffer();
  let text;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(buf);
  } catch {
    text = new TextDecoder('windows-1252').decode(buf);
  }
  let lines = text.replace(/^﻿/, '').split(/\r?\n/);
  if (/\.csv$/i.test(file.name)) {
    const first = lines.find((l) => l.trim()) || '';
    const sep = (first.match(/;/g) || []).length >= (first.match(/,/g) || []).length ? ';' : ',';
    lines = lines.map((l) => firstCsvCell(l, sep));
  }
  return lines.map((l) => l.trim()).filter(Boolean);
}

export async function render(root, { params, ctx }) {
  const id = params[0];
  const list = id ? await get(`/api/lists/${id}`) : { name: '', description: '', terms: [] };
  const validators = ctx.info.validators || {};
  const presets = ctx.info.presets || [];
  let page = 1;
  let filter = '';
  let dirty = false;

  const markDirty = () => {
    dirty = true;
  };

  const visibleTerms = () => {
    const f = fold(filter);
    const rows = list.terms.map((t, i) => [t, i]);
    return f ? rows.filter(([t]) => fold(`${t.value} ${t.label || ''}`).includes(f)) : rows;
  };

  const termRow = ([t, i]) => html`<tr data-index="${i}">
    <td>
      <select data-field="type" aria-label="Tipo do termo ${i + 1}">
        <option value="text" ${t.type === 'text' ? 'selected' : ''}>Texto</option>
        <option value="regex" ${t.type === 'regex' ? 'selected' : ''}>Expressão regular</option>
      </select>
    </td>
    <td><input type="text" data-field="value" class="${t.type === 'regex' ? 'mono' : ''}" value="${t.value}" aria-label="Termo ${i + 1}" /></td>
    <td>
      ${t.type === 'text'
        ? html`<label class="check small"><input type="checkbox" data-field="wholeWord" ${t.wholeWord ? 'checked' : ''} /><span>Palavra inteira</span></label>`
        : html`<select data-field="validator" aria-label="Validação do termo ${i + 1}">
            <option value="">Sem validação</option>
            ${Object.entries(validators).map(([key, label]) => html`<option value="${key}" ${t.validator === key ? 'selected' : ''}>${label}</option>`)}
          </select>`}
    </td>
    <td>
      ${t.type === 'regex'
        ? html`<input type="text" data-field="label" value="${t.label || ''}" placeholder="Ex.: CPF" aria-label="Nome no relatório do termo ${i + 1}" />`
        : html`<span class="muted small">o próprio termo</span>`}
    </td>
    <td class="actions"><button type="button" class="icon-btn danger" data-action="remove" aria-label="Remover termo ${i + 1}" title="Remover">${icon('trash')}</button></td>
  </tr>`;

  const drawTerms = () => {
    const box = root.querySelector('[data-terms]');
    const rows = visibleTerms();
    const pages = Math.max(1, Math.ceil(rows.length / PAGE_SIZE));
    page = Math.min(page, pages);
    const slice = rows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);
    root.querySelector('[data-term-count]').textContent = `(${fmtNum(list.terms.length)})`;
    paint(
      box,
      list.terms.length === 0
        ? html`<div class="empty">Nenhum termo ainda. Adicione termos, importe um arquivo ou use um modelo pronto.</div>`
        : html`<div class="table-wrap">
              <table class="data terms-table">
                <thead><tr><th>Tipo</th><th>Termo ou expressão</th><th>Opções</th><th>Nome no relatório</th><th><span class="sr-only">Remover</span></th></tr></thead>
                <tbody>${slice.map(termRow)}</tbody>
              </table>
            </div>
            ${rows.length === 0 ? html`<div class="empty">Nenhum termo corresponde ao filtro.</div>` : ''}
            ${pages > 1
              ? html`<div class="pager">
                  <span class="muted small">${fmtNum(rows.length)} termo(s) · página ${page} de ${pages}</span>
                  <div class="inline">
                    <button type="button" class="btn small" data-action="page" data-page="${page - 1}" ${page <= 1 ? 'disabled' : ''}>Anterior</button>
                    <button type="button" class="btn small" data-action="page" data-page="${page + 1}" ${page >= pages ? 'disabled' : ''}>Próxima</button>
                  </div>
                </div>`
              : ''}`,
    );
  };

  const draw = () => {
    paint(
      root,
      html`<div class="page-head">
          <div>
            <h1>${id ? 'Editar lista de referência' : 'Nova lista de referência'}</h1>
            <div class="sub">Termos de texto ignoram maiúsculas e acentos ("salario" encontra "SALÁRIO"). Expressões regulares seguem a sintaxe do JavaScript.</div>
          </div>
          <div class="actions">
            <button type="button" class="btn" data-action="back">Voltar</button>
            <button type="button" class="btn primary" data-action="save">${icon('check')} Salvar lista</button>
          </div>
        </div>

        <section class="card">
          <div class="form-grid">
            <label class="field"><span>Nome da lista</span><input type="text" name="name" maxlength="200" value="${list.name}" placeholder="Ex.: Dados pessoais (LGPD)" /></label>
            <label class="field"><span>Descrição (opcional)</span><input type="text" name="description" maxlength="1000" value="${list.description || ''}" /></label>
          </div>
        </section>

        <section class="card">
          <div class="card-head">
            <h2>Termos <span class="muted" data-term-count></span></h2>
            <div class="inline">
              <button type="button" class="btn small" data-action="add">${icon('plus')} Termo</button>
              <button type="button" class="btn small" data-action="bulk">${icon('list')} Adicionar vários</button>
              <button type="button" class="btn small" data-action="import">${icon('upload')} Importar .txt/.csv</button>
              <input type="file" accept=".txt,.csv,text/plain,text/csv" data-file hidden />
              <select data-preset aria-label="Adicionar modelo pronto">
                <option value="">Adicionar modelo pronto…</option>
                ${presets.map((p) => html`<option value="${p.id}">${p.label} — ${p.description}</option>`)}
              </select>
              <input type="search" data-filter placeholder="Filtrar termos" aria-label="Filtrar termos" value="${filter}" />
            </div>
          </div>
          <div data-terms></div>
        </section>

        <section class="card test-panel">
          <h2>Testar os termos</h2>
          <p class="muted small">Cole um trecho de texto para ver o que seria encontrado (inclui alterações ainda não salvas).</p>
          <textarea data-test-text rows="5" placeholder="Ex.: O salário de João da Silva, CPF 529.982.247-25, foi pago."></textarea>
          <div class="inline"><button type="button" class="btn" data-action="test">${icon('flask')} Testar</button></div>
          <div data-test-result></div>
        </section>`,
    );
    drawTerms();
  };

  const addTerms = (values, { type = 'text', wholeWord = false } = {}) => {
    const keys = new Set(list.terms.map(termKey));
    let added = 0;
    for (const value of values) {
      const term = { type, value, wholeWord: type === 'text' && wholeWord, validator: null, label: '' };
      if (!value || keys.has(termKey(term))) continue;
      keys.add(termKey(term));
      list.terms.push(term);
      added++;
    }
    if (added) markDirty();
    return { added, skipped: values.length - added };
  };

  const bulkDialog = async (prefill = [], title = 'Adicionar vários termos') => {
    const result = await openDialog({
      title,
      wide: true,
      submitLabel: 'Adicionar',
      body: html`<div class="form-grid">
        <label class="field full">
          <span>Termos (um por linha)</span>
          <textarea name="terms" rows="10">${prefill.join('\n')}</textarea>
          <small>Linhas repetidas ou já existentes na lista são ignoradas.</small>
        </label>
        <label class="field">
          <span>Tipo</span>
          <select name="type"><option value="text">Texto</option><option value="regex">Expressão regular</option></select>
        </label>
        <label class="check"><input type="checkbox" name="wholeWord" /><span>Palavra inteira (só para texto): "ana" não encontra "banana"</span></label>
      </div>`,
      onSubmit: (form) => {
        const values = form.elements.terms.value
          .split(/\r?\n/)
          .map((v) => v.trim())
          .filter(Boolean);
        if (values.length === 0) throw new Error('Informe ao menos um termo.');
        return addTerms(values, { type: form.elements.type.value, wholeWord: form.elements.wholeWord.checked });
      },
    });
    if (result) {
      filter = '';
      root.querySelector('[data-filter]').value = '';
      page = Math.ceil(list.terms.length / PAGE_SIZE) || 1;
      drawTerms();
      toast(`${plural(result.added, 'termo adicionado', 'termos adicionados')}${result.skipped ? ` (${fmtNum(result.skipped)} repetido(s) ignorado(s))` : ''}.`, 'success');
    }
  };

  const save = async () => {
    list.name = root.querySelector('[name="name"]').value.trim();
    list.description = root.querySelector('[name="description"]').value.trim();
    if (!list.name) {
      toast('Informe o nome da lista.', 'error');
      root.querySelector('[name="name"]').focus();
      return;
    }
    const body = { name: list.name, description: list.description, terms: list.terms.filter((t) => String(t.value).trim()) };
    try {
      const saved = id ? await put(`/api/lists/${id}`, body) : await post('/api/lists', body);
      dirty = false;
      toast(`Lista salva com ${plural(saved.terms.length, 'termo', 'termos')}.`, 'success');
      if (!id) {
        go(`/listas/${saved.id}`);
        return;
      }
      list.terms = saved.terms;
      drawTerms();
    } catch (err) {
      toast(err.message, 'error');
    }
  };

  const runTest = async () => {
    const box = root.querySelector('[data-test-result]');
    const text = root.querySelector('[data-test-text]').value;
    if (!text.trim()) {
      paint(box, html`<p class="muted small">Digite um texto para testar.</p>`);
      return;
    }
    try {
      const { matches } = await post('/api/lists/test', { terms: list.terms.filter((t) => String(t.value).trim()), text });
      paint(
        box,
        matches.length === 0
          ? html`<p class="muted">Nenhum termo encontrado no texto.</p>`
          : matches.map(
              (m) => html`<div class="match">
                <div class="match-head"><span class="chip"><b>${m.term}</b></span><span class="muted small">${plural(m.count, 'ocorrência', 'ocorrências')}</span></div>
                ${m.samples.map(
                  (s) => html`<div class="sample">${s.where ? html`<span class="where">${s.where}</span>` : ''}${s.before}<mark>${s.match}</mark>${s.after}</div>`,
                )}
              </div>`,
            ),
      );
    } catch (err) {
      paint(box, html`<div class="alert error">${icon('alert')}<div>${err.message}</div></div>`);
    }
  };

  const onClick = async (event) => {
    const el = event.target.closest('[data-action]');
    if (!el) return;
    const action = el.dataset.action;
    if (action === 'save') return save();
    if (action === 'back') {
      if (dirty && !(await confirmDialog('Há alterações não salvas. Sair mesmo assim?', { confirmLabel: 'Sair sem salvar' }))) return;
      dirty = false;
      return go('/listas');
    }
    if (action === 'add') {
      list.terms.push({ type: 'text', value: '', wholeWord: false, validator: null, label: '' });
      markDirty();
      filter = '';
      root.querySelector('[data-filter]').value = '';
      page = Math.ceil(list.terms.length / PAGE_SIZE);
      drawTerms();
      root.querySelector(`tr[data-index="${list.terms.length - 1}"] [data-field="value"]`)?.focus();
      return;
    }
    if (action === 'bulk') return bulkDialog();
    if (action === 'import') return root.querySelector('[data-file]').click();
    if (action === 'remove') {
      const index = Number(el.closest('tr').dataset.index);
      list.terms.splice(index, 1);
      markDirty();
      drawTerms();
      return;
    }
    if (action === 'page') {
      page = Number(el.dataset.page);
      drawTerms();
      return;
    }
    if (action === 'test') return runTest();
  };

  const updateField = (event) => {
    const field = event.target.dataset.field;
    const row = event.target.closest('tr[data-index]');
    if (!field || !row) return;
    const term = list.terms[Number(row.dataset.index)];
    if (!term) return;
    markDirty();
    if (field === 'type') {
      term.type = event.target.value;
      term.wholeWord = false;
      term.validator = null;
      drawTerms();
    } else if (field === 'wholeWord') {
      term.wholeWord = event.target.checked;
    } else if (field === 'validator') {
      term.validator = event.target.value || null;
    } else {
      term[field] = event.target.value;
    }
  };

  const onChange = async (event) => {
    if (event.target.matches('[data-file]')) {
      const file = event.target.files[0];
      event.target.value = '';
      if (!file) return;
      try {
        const values = await readTermsFile(file);
        if (values.length === 0) return toast('O arquivo não tem termos.', 'error');
        await bulkDialog(values, `Importar ${file.name}`);
      } catch (err) {
        toast(`Não foi possível ler o arquivo: ${err.message}`, 'error');
      }
      return;
    }
    if (event.target.matches('[data-preset]')) {
      const preset = presets.find((p) => p.id === event.target.value);
      event.target.value = '';
      if (!preset) return;
      const term = { type: 'regex', value: preset.value, validator: preset.validator, label: preset.label, wholeWord: false };
      if (list.terms.some((t) => termKey(t) === termKey(term))) return toast(`O modelo ${preset.label} já está na lista.`);
      list.terms.push(term);
      markDirty();
      page = Math.ceil(list.terms.length / PAGE_SIZE);
      drawTerms();
      toast(`Modelo ${preset.label} adicionado.`, 'success');
      return;
    }
    updateField(event);
  };

  const onInput = (event) => {
    if (event.target.matches('[data-filter]')) return filterTerms(event.target.value);
    if (event.target.matches('[name="name"], [name="description"]')) return markDirty();
    if (event.target.dataset.field === 'value' || event.target.dataset.field === 'label') updateField(event);
  };

  const filterTerms = debounce((value) => {
    filter = value;
    page = 1;
    drawTerms();
  }, 200);

  const beforeUnload = (event) => {
    if (dirty) event.preventDefault();
  };

  draw();
  root.addEventListener('click', onClick);
  root.addEventListener('change', onChange);
  root.addEventListener('input', onInput);
  window.addEventListener('beforeunload', beforeUnload);
  return () => {
    root.removeEventListener('click', onClick);
    root.removeEventListener('change', onChange);
    root.removeEventListener('input', onInput);
    window.removeEventListener('beforeunload', beforeUnload);
  };
}
