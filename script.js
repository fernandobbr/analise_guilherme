/* ══════════════════════════════════════════════════════════════════
   DASHBOARD INDUSTRIAL — FILTROS + BUSCA POR COLABORADOR
   Fonte: dados_dashboard.json (gerado de Dados1.xlsx)
   ══════════════════════════════════════════════════════════════════ */
'use strict';

/* ─────────────────────────────────────────────────────────────────
   CONSTANTES
   ───────────────────────────────────────────────────────────────── */
const SEMANA   = 22;
const PERIODO  = { ini:'25/05/2026', fim:'29/05/2026', dias:'SEGUNDA A SEXTA' };
const HE_PREV_H = 5;
const DIAS_MAP  = ['Segunda','Terca','Quarta','Quinta','Sexta'];
const DIAS_ABREV= {0:'SEG',1:'TER',2:'QUA',3:'QUI',4:'SEX'};
const DIAS_LABEL= {Segunda:'Segunda',Terca:'Terça',Quarta:'Quarta',Quinta:'Quinta',Sexta:'Sexta'};
const DATAS_MAP = {Segunda:'25/05',Terca:'26/05',Quarta:'27/05',Quinta:'28/05',Sexta:'29/05'};

/* ─────────────────────────────────────────────────────────────────
   ESTADO GLOBAL
   ───────────────────────────────────────────────────────────────── */
let RAW    = null;   /* dados carregados do JSON */
let STATE  = { setor:'', turno:'' };
let METRICS = null;  /* métricas do último renderAll */
let searchTimer = null;

/* ─────────────────────────────────────────────────────────────────
   LOAD DATA — lê DADOS_XLSX injetado por dados.js (sem fetch/server)
   ───────────────────────────────────────────────────────────────── */
function loadData() {
  if (typeof DADOS_XLSX === 'undefined') {
    document.getElementById('dashboard').innerHTML =
      '<div class="filter-loading">Erro: arquivo dados.js não encontrado na pasta do projeto.</div>';
    return;
  }
  RAW = DADOS_XLSX;
  init();
}

/* ─────────────────────────────────────────────────────────────────
   INIT — preenche selects e renderiza
   ───────────────────────────────────────────────────────────────── */
function init() {
  populateSelects();
  bindEvents();
  renderAll();
}

function populateSelects() {
  const setores = [...new Set(RAW.colabs.map(c => c.setor).filter(Boolean))].sort();
  const turnos  = [...new Set(RAW.colabs.map(c => c.turno).filter(Boolean))].sort();

  const selSetor = document.getElementById('filter-setor');
  setores.forEach(s => {
    const o = document.createElement('option');
    o.value = s; o.textContent = s;
    selSetor.appendChild(o);
  });

  const selTurno = document.getElementById('filter-turno');
  turnos.forEach(t => {
    const o = document.createElement('option');
    o.value = t; o.textContent = t;
    selTurno.appendChild(o);
  });
}

/* ─────────────────────────────────────────────────────────────────
   EVENTS
   ───────────────────────────────────────────────────────────────── */
function bindEvents() {
  document.getElementById('filter-setor').addEventListener('change', e => {
    STATE.setor = e.target.value;
    updateActiveTags();
    renderAll();
  });

  document.getElementById('filter-turno').addEventListener('change', e => {
    STATE.turno = e.target.value;
    updateActiveTags();
    renderAll();
  });

  document.getElementById('filter-reset').addEventListener('click', resetFilters);

  /* Search */
  const input = document.getElementById('search-colab');
  const clearBtn = document.getElementById('search-clear');

  input.addEventListener('input', () => {
    clearBtn.hidden = !input.value;
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => doSearch(input.value.trim()), 250);
  });

  input.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeSearch();
  });

  clearBtn.addEventListener('click', () => {
    input.value = '';
    clearBtn.hidden = true;
    closeSearch();
    input.focus();
  });

  /* Fecha dropdown ao clicar fora */
  document.addEventListener('click', e => {
    if (!e.target.closest('.filter-bar__search-wrap')) closeSearch();
  });

  /* Filtro HE */
  document.getElementById('he-search').addEventListener('input', () => {
    if (METRICS) renderHETable(METRICS);
  });

  /* Upload planilha */
  document.getElementById('btn-upload').addEventListener('click', () => {
    document.getElementById('xlsx-input').click();
  });

  document.getElementById('xlsx-input').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const labelText = document.getElementById('upload-label-text');
    labelText.textContent = 'Processando…';
    const reader = new FileReader();
    reader.onerror = function() {
      labelText.textContent = 'Carregar planilha';
      alert('Não foi possível ler o arquivo.');
    };
    reader.onload = function(ev) {
      try {
        const data = parseXlsxData(ev.target.result);
        console.log('[upload] colabs:', data.colabs.length, '| registros:', data.registros.length, '| HE:', data.heSabado.length);
        reloadFromData(data);
        labelText.textContent = `✓ ${data.colabs.length} colab. / ${data.registros.length} reg.`;
        setTimeout(() => { labelText.textContent = 'Carregar planilha'; }, 5000);
      } catch(err) {
        console.error('[upload] erro:', err);
        labelText.textContent = 'Erro — ver console';
        alert('Erro ao processar planilha:\n\n' + err.message);
        setTimeout(() => { labelText.textContent = 'Carregar planilha'; }, 4000);
      }
      e.target.value = '';
    };
    reader.readAsArrayBuffer(file);
  });

  /* Modal */
  document.getElementById('modal-close').addEventListener('click', closeModal);
  document.getElementById('modal-overlay').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeModal();
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeModal();
  });
}

function resetFilters() {
  STATE.setor = '';
  STATE.turno = '';
  document.getElementById('filter-setor').value = '';
  document.getElementById('filter-turno').value = '';
  document.getElementById('search-colab').value = '';
  document.getElementById('search-clear').hidden = true;
  closeSearch();
  updateActiveTags();
  renderAll();
}

function updateActiveTags() {
  const wrap = document.getElementById('filter-active-tags');
  wrap.innerHTML = '';
  if (STATE.setor) wrap.appendChild(makeTag(STATE.setor, () => {
    STATE.setor = '';
    document.getElementById('filter-setor').value = '';
    updateActiveTags(); renderAll();
  }));
  if (STATE.turno) wrap.appendChild(makeTag(STATE.turno, () => {
    STATE.turno = '';
    document.getElementById('filter-turno').value = '';
    updateActiveTags(); renderAll();
  }));
}

function makeTag(label, onRemove) {
  const tag = document.createElement('span');
  tag.className = 'filter-tag';
  tag.innerHTML = `${label} <button class="filter-tag__remove" title="Remover filtro">✕</button>`;
  tag.querySelector('button').addEventListener('click', onRemove);
  return tag;
}

/* ─────────────────────────────────────────────────────────────────
   FILTER LOGIC
   ───────────────────────────────────────────────────────────────── */
function getFilteredColabs() {
  return RAW.colabs.filter(c => {
    if (STATE.setor && c.setor !== STATE.setor) return false;
    if (STATE.turno && c.turno !== STATE.turno) return false;
    return true;
  });
}

function getRegistrosForMats(mats) {
  const matSet = new Set(mats);
  return RAW.registros.filter(r => matSet.has(r.m));
}

/* ─────────────────────────────────────────────────────────────────
   AFASTAMENTO FORMAL — situações que NÃO contam como ausência
   (afastamento temporário, licença, férias, aposentadoria)
   ───────────────────────────────────────────────────────────────── */
const AFAST_FORMAL_KEYS = [
  'AFASTAMENTO', 'LICEN', 'APOSENTADORIA', 'INVALIDEZ',
  'FERIAS', 'FÉRIAS', 'FERIA',
];
function isAfastFormal(motivo) {
  const m = (motivo || '').toUpperCase();
  return AFAST_FORMAL_KEYS.some(k => m.includes(k));
}

/* ─────────────────────────────────────────────────────────────────
   COMPUTE METRICS — dinâmico por filtro
   ───────────────────────────────────────────────────────────────── */
function computeMetrics(colabs, registros) {
  /* ── Por dia ── */
  const byDay = {};
  DIAS_MAP.forEach(d => byDay[d] = {
    dia: DIAS_LABEL[d], abrev: DIAS_ABREV[DIAS_MAP.indexOf(d)],
    data: DATAS_MAP[d],
    colabs:    new Set(),
    presentes: new Set(),
    hPrev:0, hTrab:0, hPerd:0,
    faltas:0, atestados:0, atrasos:0,
    faltasMat:    new Set(),   /* falta dia inteiro  (hTrab = 0) */
    meioPerMat:   new Set(),   /* falta meio período (hTrab > 0) */
    atestadosMat: new Set(),
  });

  registros.forEach(r => {
    const day = byDay[r.d];
    if (!day) return;
    if (r.p > 0) day.colabs.add(r.m);
    if (r.t > 0) day.presentes.add(r.m);
    day.hPrev += r.p;
    day.hTrab += r.t;
    day.hPerd += r.x;
    const mot = (r.o || '').toUpperCase();

    if (mot.includes('FALTA') && !mot.includes('JUSTIF') && !isAfastFormal(r.o)) {
      if (r.t === 0) {
        day.faltasMat.add(r.m);    /* dia inteiro */
      } else {
        day.meioPerMat.add(r.m);   /* meio período — trabalhou mas faltou parte */
      }
    }
    if ((mot.includes('ATESTADO') || mot.includes('DECLARACAO')) && !isAfastFormal(r.o)) {
      day.atestadosMat.add(r.m);
    }
    if (mot.includes('ATRASO') || mot.includes('ANTECIPADA')) day.atrasos++;
  });

  const porDia = DIAS_MAP.map(d => {
    const bd        = byDay[d];
    const faltas    = bd.faltasMat.size;
    const meioPer   = bd.meioPerMat.size;
    const atestados = bd.atestadosMat.size;
    /* TOTAL AUSÊNCIAS = faltas + meio período + atestados */
    const aus = faltas + meioPer + atestados;
    return {
      ...bd,
      ideal:     bd.colabs.size,
      presentes: bd.presentes.size,
      faltas,
      meioPer,
      atestados,
      aus,
    };
  });

  /* ── Totais ── */
  const hPrev  = porDia.reduce((s,d) => s + d.hPrev, 0);
  const hTrab  = porDia.reduce((s,d) => s + d.hTrab, 0);
  const hPerd  = porDia.reduce((s,d) => s + d.hPerd, 0);
  const efic   = hPrev > 0 ? hTrab / hPrev : 0;
  const avgPct = efic;
  const avgPres= porDia.reduce((s,d) => s + d.presentes, 0) / porDia.length;

  /* Headcount por status — aba Dados_Colaboradores filtrada por setor+turno
     Ideal = Ativos + Vagas Abertas (exclui Afastados conforme regra de negócio) */
  const countAtivo = colabs.filter(c => c.status === 'Ativo').length;
  const countVagas = colabs.filter(c => c.status === 'Vaga Aberta').length;
  const countAfas  = colabs.filter(c => c.status === 'Afastado').length;
  const avgIdeal   = countAtivo + countVagas;

  /* ── HE Sábado — filtrado por setor/turno ── */
  const heFiltrada = RAW.heSabado.filter(w => {
    if (STATE.setor && w.setor !== STATE.setor) return false;
    if (STATE.turno && w.turno !== STATE.turno) return false;
    return true;
  });
  const heConv  = heFiltrada.length;
  const heComp  = heFiltrada.filter(w => w.s).length;
  const heAus   = heConv - heComp;
  const hePrevH = heConv * HE_PREV_H;
  const heTrabH = heFiltrada.reduce((s, w) => s + w.h, 0) / 3600;
  const heAprov = hePrevH > 0 ? heTrabH / hePrevH : 0;

  return {
    porDia, avgIdeal, avgPres, avgPct,
    countAtivo, countVagas, countAfas,
    hPrev, hTrab, hPerd, efic,
    heConv, heComp, heAus, hePrevH, heTrabH, heAprov,
    heFiltrada,
    totalColabs: colabs.length,
  };
}

/* ─────────────────────────────────────────────────────────────────
   FORMATADORES
   ───────────────────────────────────────────────────────────────── */
const F = {
  pct : v => (v * 100).toFixed(1).replace('.', ',') + '%',
  h   : v => v.toFixed(1).replace('.', ',') + 'h',
  hInt: v => Math.round(v).toLocaleString('pt-BR') + 'h',
  hMM : sec => {
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
  },
  dec : v => Number(v.toFixed(1)).toLocaleString('pt-BR'),
};

function pBadge(pct) {
  const c = pct >= .95 ? 'g' : pct >= .92 ? 'y' : 'r';
  return `<span class="pbadge pbadge--${c}">${F.pct(pct)}</span>`;
}

/* ─────────────────────────────────────────────────────────────────
   RENDER ALL — ponto de entrada principal
   ───────────────────────────────────────────────────────────────── */
const CONTENT_SECTIONS = ['.kpi-strip', '.ops-grid', '.chart-grid'];

function setContentVisible(visible) {
  const msg = document.getElementById('no-filter-msg');
  msg.hidden = visible;
  CONTENT_SECTIONS.forEach(sel => {
    const el = document.querySelector(sel);
    if (el) el.hidden = !visible;
  });
}

function renderAll() {
  if (!STATE.setor && !STATE.turno) {
    setContentVisible(false);
    renderHeader();
    return;
  }

  setContentVisible(true);

  const colabs     = getFilteredColabs();
  const mats       = colabs.map(c => c.mat).filter(Boolean);
  const registros  = getRegistrosForMats(mats);
  const m          = computeMetrics(colabs, registros);
  METRICS          = m;

  const setorLabel = STATE.setor || 'Todos os Setores';
  const turnoLabel = STATE.turno || 'Todos os Turnos';

  renderHeader();
  renderKPIs(m);
  renderDailyTable(m);
  renderHorasProd(m);
  renderHESummary(m);
  renderBarChart(m);
  renderDonut(m);
  renderHETable(m);
  renderInsights(m, setorLabel, turnoLabel);
}

/* ─────────────────────────────────────────────────────────────────
   RENDER: HEADER
   ───────────────────────────────────────────────────────────────── */
function renderHeader() {
  const setor = STATE.setor || 'TODOS OS SETORES';
  const turno = STATE.turno;

  /* Setor: mostra filtro ativo ou "TODOS OS SETORES" */
  document.getElementById('hdr-setor').textContent = setor.toUpperCase();

  /* Rótulo muda conforme filtros */
  document.getElementById('hdr-setor-lbl').textContent =
    STATE.setor ? 'SETOR:' : 'VISUALIZAÇÃO:';

  /* Linha de turno — visível apenas quando filtro de turno está ativo */
  const turnoLine = document.getElementById('hdr-turno-line');
  if (turno) {
    document.getElementById('hdr-turno').textContent = turno.toUpperCase();
    turnoLine.style.display = '';
  } else {
    turnoLine.style.display = 'none';
  }

  document.getElementById('hdr-supervisor').textContent = '_______________';
  document.getElementById('hdr-periodo').textContent    = `${PERIODO.ini} a ${PERIODO.fim}`;
  document.getElementById('hdr-semana').textContent     = `SEMANA ${SEMANA} · ${PERIODO.dias}`;
}

/* ─────────────────────────────────────────────────────────────────
   RENDER: KPI STRIP
   ───────────────────────────────────────────────────────────────── */
function renderKPIs(m) {
  /* Card 1 — Colaboradores Ideais (Ativos + Vagas, excl. Afastados) */
  document.getElementById('kv-ideal').textContent = m.avgIdeal;

  /* Card 2 — Status: Ativos | Vagas Abertas | Afastados */
  document.getElementById('kv-status-ativo').textContent = m.countAtivo;
  document.getElementById('kv-status-vagas').textContent = m.countVagas;
  document.getElementById('kv-status-afas').textContent  = m.countAfas;

  const circ = 2 * Math.PI * 32;
  const fill = m.avgPct * circ;
  const arc  = document.getElementById('ring-arc');
  arc.setAttribute('stroke-dasharray', `${fill.toFixed(2)} ${circ.toFixed(2)}`);
  const cc = m.avgPct >= .95 ? '' : m.avgPct >= .92 ? 'ring-fill--yellow' : 'ring-fill--red';
  arc.setAttribute('class', `ring-fill ${cc}`.trim());
  document.getElementById('kv-presenca-pct').textContent = F.pct(m.avgPct);

  document.getElementById('kv-htrab').textContent = F.h(m.hTrab);
  document.getElementById('kv-hperd').textContent = F.h(m.hPerd);
  document.getElementById('kv-he').textContent    = F.h(m.heTrabH);
}

/* ─────────────────────────────────────────────────────────────────
   RENDER: COMPARATIVO DIÁRIO
   ───────────────────────────────────────────────────────────────── */
function renderDailyTable(m) {
  document.getElementById('daily-tbody').innerHTML =
    m.porDia.map(d => {
      const pct = d.hPrev > 0 ? d.hTrab / d.hPrev : 0;
      return `
        <tr class="dtable__row--clickable" data-dia="${d.dia}" title="Clique para ver detalhes de ${d.dia}">
          <td class="ta-l">${d.dia}<span class="row-detail-hint">▶</span></td>
          <td>${m.avgIdeal}</td>
          <td><strong>${d.presentes}</strong></td>
          <td>${d.faltas    > 0 ? `<span class="abs-badge abs-badge--red">${d.faltas}</span>`       : d.faltas}</td>
          <td>${d.meioPer   > 0 ? `<span class="abs-badge abs-badge--orange">${d.meioPer}</span>`  : d.meioPer}</td>
          <td>${d.atestados > 0 ? `<span class="abs-badge abs-badge--orange">${d.atestados}</span>`: d.atestados}</td>
          <td>${d.aus       > 0 ? `<span class="abs-badge abs-badge--gray">${d.aus}</span>`        : d.aus}</td>
          <td>${pBadge(pct)}</td>
        </tr>`;
    }).join('');

  /* Delegação de eventos — clique em qualquer linha do comparativo */
  document.getElementById('daily-tbody').onclick = e => {
    const tr = e.target.closest('tr[data-dia]');
    if (tr) openDayDetail(tr.dataset.dia, m);
  };

  const n       = m.porDia.length;
  const aPres    = F.dec(m.porDia.reduce((s,d)=>s+d.presentes,0)/n);
  const aFalt    = F.dec(m.porDia.reduce((s,d)=>s+d.faltas,0)/n);
  const aMeio    = F.dec(m.porDia.reduce((s,d)=>s+d.meioPer,0)/n);
  const aAtes    = F.dec(m.porDia.reduce((s,d)=>s+d.atestados,0)/n);
  const aAus     = F.dec(m.porDia.reduce((s,d)=>s+d.aus,0)/n);

  document.getElementById('daily-tfoot').innerHTML = `
    <tr>
      <td class="ta-l">MÉDIA SEMANAL</td>
      <td>${m.avgIdeal}</td><td>${aPres}</td>
      <td>${aFalt}</td><td>${aMeio}</td><td>${aAtes}</td>
      <td>${aAus}</td><td>${F.pct(m.avgPct)}</td>
    </tr>`;
}

/* ─────────────────────────────────────────────────────────────────
   MODAL DETALHE DO DIA — ausências e ocorrências
   ───────────────────────────────────────────────────────────────── */
function openDayDetail(dia, m) {
  /* Matriculas do filtro atual */
  const colabs    = getFilteredColabs();
  const matSet    = new Set(colabs.map(c => c.mat).filter(Boolean));
  const colabMap  = Object.fromEntries(colabs.map(c => [c.mat, c]));

  /* Registros do dia selecionado */
  const regsDay   = RAW.registros.filter(r => r.d === dia && matSet.has(r.m));

  /* Dados do dia nos metrics */
  const diaData   = m.porDia.find(d => d.dia === dia || d.dia === DIAS_LABEL[Object.keys(DIAS_LABEL).find(k => DIAS_LABEL[k] === dia)]);
  const dataFmt   = diaData ? diaData.data : '';

  /* Usa isAfastFormal() definida globalmente */

  /* Categorias — exclui afastados formais e licenças */
  const ausentes = regsDay.filter(r =>
    r.t === 0 && r.p > 0 && !isAfastFormal(r.o)
  );
  const ocorrencias = regsDay.filter(r => {
    const m = (r.o || '').toUpperCase();
    return r.o && r.o !== '' && r.o !== '-'
      && !m.includes('HORA EXTRA') && !m.includes('COMPENSA')
      && !isAfastFormal(r.o);
  });

  /* Remove duplicatas — prefere ocorrência sobre ausência pura */
  const vistos = new Set();
  const linhas = [];

  ocorrencias.forEach(r => {
    if (!vistos.has(r.m)) {
      vistos.add(r.m);
      linhas.push(r);
    }
  });
  ausentes.forEach(r => {
    if (!vistos.has(r.m)) {
      vistos.add(r.m);
      linhas.push(r);
    }
  });

  linhas.sort((a, b) => {
    const ma = (a.o||'').toUpperCase(), mb = (b.o||'').toUpperCase();
    if (ma.includes('FALTA') && !mb.includes('FALTA')) return -1;
    if (mb.includes('FALTA') && !ma.includes('FALTA')) return 1;
    return 0;
  });

  /* Título */
  document.getElementById('modal-title').textContent =
    `AUSÊNCIAS E OCORRÊNCIAS — ${dia.toUpperCase()}`;
  document.getElementById('modal-meta').innerHTML = `
    <span class="modal-box__meta-chip">Data: <strong>${dataFmt}</strong></span>
    <span class="modal-box__meta-chip">Ideal: <strong>${m.avgIdeal}</strong></span>
    <span class="modal-box__meta-chip">Presentes: <strong>${diaData ? diaData.presentes : '—'}</strong></span>
    <span class="modal-box__meta-chip">Ausências: <strong>${diaData ? m.avgIdeal - diaData.presentes : '—'}</strong></span>`;

  if (!linhas.length) {
    document.getElementById('modal-body').innerHTML =
      '<p style="text-align:center;color:#6a7f9a;padding:24px 0">Nenhuma ocorrência ou ausência registrada neste dia.</p>';
  } else {
    const rows = linhas.map((r, i) => {
      const c   = colabMap[r.m] || {};
      const mot = (r.o || '').replace(/\[LM\]/g, '').trim() || 'Ausência sem motivo';
      const motU= mot.toUpperCase();
      let tipo  = 'other', tipoLabel = 'OCORRÊNCIA';
      if (motU.includes('FALTA'))                              { tipo = 'falta';    tipoLabel = 'FALTA'; }
      else if (motU.includes('ATESTADO') || motU.includes('DECLARACAO')) { tipo = 'atestado'; tipoLabel = 'ATESTADO'; }
      else if (motU.includes('ATRASO') || motU.includes('ANTECIPADA'))   { tipo = 'atraso';   tipoLabel = 'ATRASO'; }
      else if (motU.includes('AFASTAMENTO') || motU.includes('LICENCA')) { tipo = 'afas';     tipoLabel = 'AFASTAMENTO'; }
      else if (r.t === 0 && r.p > 0)                          { tipo = 'falta';    tipoLabel = 'AUSENTE'; }

      const hPerd = r.x > 0 ? F.h(r.x) : '—';
      return `
        <tr>
          <td class="col-n">${i+1}</td>
          <td class="ta-l" style="font-family:var(--f-sans);font-size:.72rem">${c.nome || r.m}</td>
          <td style="font-family:var(--f-mono);font-size:.7rem;color:var(--text3)">${r.m}</td>
          <td><span class="day-badge day-badge--${tipo}">${tipoLabel}</span></td>
          <td style="font-family:var(--f-sans);font-size:.68rem;color:var(--text2);max-width:160px;
                     white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${mot}">${mot}</td>
          <td>${hPerd}</td>
        </tr>`;
    }).join('');

    document.getElementById('modal-body').innerHTML = `
      <table class="modal-table">
        <thead>
          <tr>
            <th class="col-n">Nº</th>
            <th class="ta-l">NOME</th>
            <th>MATRÍCULA</th>
            <th>TIPO</th>
            <th class="ta-l">MOTIVO / OCORRÊNCIA</th>
            <th>H. PERDIDAS</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
        <tfoot>
          <tr>
            <td colspan="2" class="ta-l">TOTAL</td>
            <td colspan="2">${linhas.length} registro(s)</td>
            <td>—</td>
            <td>${F.h(linhas.reduce((s,r)=>s+r.x,0))}</td>
          </tr>
        </tfoot>
      </table>`;
  }

  document.getElementById('modal-overlay').hidden = false;
  document.getElementById('modal-close').focus();
}

/* ─────────────────────────────────────────────────────────────────
   RENDER: HORAS PRODUÇÃO
   ───────────────────────────────────────────────────────────────── */
function renderHorasProd(m) {
  document.getElementById('kv-hprev').textContent  = F.h(m.hPrev);
  document.getElementById('kv-htrab2').textContent = F.h(m.hTrab);
  document.getElementById('kv-hperd2').textContent = F.h(m.hPerd);
  document.getElementById('kv-efic').textContent   = F.pct(m.efic);
}

/* ─────────────────────────────────────────────────────────────────
   RENDER: HE SÁBADO
   ───────────────────────────────────────────────────────────────── */
function renderHESummary(m) {
  document.getElementById('kv-he-conv').textContent  = m.heConv;
  document.getElementById('kv-he-comp').textContent  = m.heComp;
  document.getElementById('kv-he-aus').textContent   = m.heAus;
  document.getElementById('kv-he-hprev').textContent = F.hInt(m.hePrevH);
  document.getElementById('kv-he-hreal').textContent = F.h(m.heTrabH);
  document.getElementById('kv-aprov').textContent    = F.pct(m.heAprov);
}

/* ─────────────────────────────────────────────────────────────────
   RENDER: BAR CHART
   ───────────────────────────────────────────────────────────────── */
function renderBarChart(m) {
  const W=380, H=260, padL=36, padR=10, padT=20, padB=30;
  const cW=W-padL-padR, cH=H-padT-padB;

  const pcts  = m.porDia.map(d => d.hPrev>0 ? d.hTrab/d.hPrev : 0);
  const minVal= Math.max(0, Math.min(...pcts) - 0.04);
  const maxVal= Math.max(...pcts) + 0.06;          /* sem cap em 100% */
  const yRange= maxVal - minVal;
  const yPx   = v => padT + cH * (1 - (v-minVal)/yRange);
  const yH    = v => cH * ((v-minVal)/yRange);

  /* Rótulos do eixo Y */
  const step   = yRange < 0.06 ? 0.01 : yRange < 0.15 ? 0.02 : 0.05;
  const gStart = Math.ceil(minVal / step) * step;
  const yLabels = [];
  for (let y = gStart; y <= maxVal + 0.001; y = Math.round((y+step)*1000)/1000) {
    const py = yPx(y);
    yLabels.push(`
      <text x="${padL-5}" y="${(py+3.5).toFixed(1)}" text-anchor="end"
            font-family="IBM Plex Mono,monospace" font-size="9" fill="#52637a">
        ${(y*100).toFixed(0)}%
      </text>`);
  }


  const slotW = cW / m.porDia.length;
  const barW  = Math.min(slotW * 0.6, 54);

  const bars = m.porDia.map((d, i) => {
    const pct = d.hPrev>0 ? d.hTrab/d.hPrev : 0;
    const cx  = padL + i*slotW + slotW/2;
    const bh  = Math.max(yH(pct), 1);
    const by  = yPx(pct);
    const lblY = Math.max(by - 5, padT + 10); /* sempre dentro do SVG */
    return `
      <rect x="${(cx-barW/2).toFixed(1)}" y="${by.toFixed(1)}"
            width="${barW}" height="${bh.toFixed(1)}" fill="#2196f3" rx="2"/>
      <text x="${cx.toFixed(1)}" y="${lblY.toFixed(1)}" text-anchor="middle"
            font-family="IBM Plex Mono,monospace" font-size="10" font-weight="700" fill="#1a3a6b">
        ${F.pct(pct)}
      </text>
      <text x="${cx.toFixed(1)}" y="${(padT+cH+16).toFixed(1)}" text-anchor="middle"
            font-family="IBM Plex Sans,sans-serif" font-size="10" font-weight="700"
            fill="#3a5070" letter-spacing="1">${d.abrev}</text>`;
  }).join('');

  /* Eixos em L — preto */
  const axes = `
    <line x1="${padL}" y1="${padT}" x2="${padL}" y2="${padT+cH}" stroke="#111" stroke-width="1.8"/>
    <line x1="${padL}" y1="${padT+cH}" x2="${W-padR}" y2="${padT+cH}" stroke="#111" stroke-width="1.8"/>`;

  document.getElementById('bar-chart').innerHTML = `
    <svg viewBox="0 0 ${W} ${H}" width="100%" height="${H}"
         role="img" aria-label="Presença diária">
      ${yLabels.join('')}${bars}${axes}
    </svg>`;
}

/* ─────────────────────────────────────────────────────────────────
   RENDER: DONUT
   ───────────────────────────────────────────────────────────────── */
function arcPath(cx,cy,R,r,a1,a2) {
  const [c,s]=[Math.cos,Math.sin];
  const p=(a,rd)=>[cx+rd*c(a),cy+rd*s(a)];
  const[ox,oy]=p(a1,R),[tx,ty]=p(a2,R);
  const[ix,iy]=p(a1,r),[jx,jy]=p(a2,r);
  const lg=(a2-a1)>Math.PI?1:0;
  return `M${ox} ${oy} A${R} ${R} 0 ${lg} 1 ${tx} ${ty} L${jx} ${jy} A${r} ${r} 0 ${lg} 0 ${ix} ${iy}Z`;
}

function renderDonut(m) {
  const total = m.hTrab + m.hPerd + m.heTrabH;
  if (total === 0) { document.getElementById('donut-wrap').innerHTML=''; return; }

  const slices = [
    { label:'Horas Trabalhadas', val:m.hTrab,   color:'#2196f3', fmt:F.h(m.hTrab)   },
    { label:'Horas Perdidas',    val:m.hPerd,   color:'#e53935', fmt:F.h(m.hPerd)   },
    { label:'HE Sábado',         val:m.heTrabH, color:'#8e24aa', fmt:F.h(m.heTrabH) },
  ];

  const S=190, cx=S/2, cy=S/2, R=84, r=46;
  let angle=-Math.PI/2, paths='', labels='';
  slices.forEach(sl => {
    const sweep=(sl.val/total)*2*Math.PI;
    paths+=`<path d="${arcPath(cx,cy,R,r,angle,angle+sweep)}"
                 fill="${sl.color}" stroke="white" stroke-width="1.5" opacity="0.93"/>`;
    const pv=sl.val/total;
    if(pv>.02){
      const mid=angle+sweep/2, lr=r+(R-r)/2;
      const fs = pv>.15 ? 10 : pv>.07 ? 9 : 8;
      labels+=`<text x="${(cx+lr*Math.cos(mid)).toFixed(1)}"
                     y="${(cy+lr*Math.sin(mid)).toFixed(1)}"
                     text-anchor="middle" dominant-baseline="middle"
                     font-family="IBM Plex Mono,monospace"
                     font-size="${fs}" font-weight="700" fill="#fff">
                 ${(pv*100).toFixed(1).replace('.',',')}%
               </text>`;
    }
    angle+=sweep;
  });

  document.getElementById('donut-wrap').innerHTML=`
    <svg viewBox="0 0 ${S} ${S}" width="${S}" height="${S}">${paths}${labels}</svg>`;
  document.getElementById('donut-legend').innerHTML=
    slices.map(sl=>`
      <div class="dl-item">
        <div class="dl-dot" style="background:${sl.color}"></div>
        <div>
          <div class="dl-label">${sl.label} <span class="dl-sub">(${sl.fmt})</span></div>
          <div class="dl-value">${(sl.val/total*100).toFixed(1).replace('.',',')}%</div>
        </div>
      </div>`).join('');
}

/* ─────────────────────────────────────────────────────────────────
   RENDER: HE TABLE
   ───────────────────────────────────────────────────────────────── */
function renderHETable(m) {
  const term  = document.getElementById('he-search').value.trim();
  const termN = normalize(term);

  const lista = m.heFiltrada.filter(w => {
    if (!term) return true;
    if (normalize(w.n).includes(termN)) return true;
    if (String(w.mat ?? '').includes(term)) return true;
    if (normalize(w.setor ?? '').includes(termN)) return true;
    const situacao = w.s ? 'compareceu' : 'ausente';
    if (situacao.includes(termN)) return true;
    return false;
  });

  document.getElementById('he-tbody').innerHTML = lista.length
    ? lista.map((w,i)=>`
      <tr>
        <td class="col-n">${i+1}</td>
        <td style="font-family:var(--f-mono);font-size:.68rem">${w.mat ?? '—'}</td>
        <td class="ta-l nome-cell">${w.n}</td>
        <td class="ta-l" style="font-size:.68rem;color:var(--text2)">${w.setor ?? '—'}</td>
        <td>${w.s?'<span class="sbadge sbadge--s">COMPARECEU</span>'
               :'<span class="sbadge sbadge--n">AUSENTE</span>'}</td>
        <td>${w.h>0 ? F.hMM(w.h) : '—'}</td>
      </tr>`).join('')
    : '<tr><td colspan="6" style="text-align:center;color:var(--text3);padding:16px">Nenhum resultado encontrado.</td></tr>';

  document.getElementById('he-foot').innerHTML=`
    <div class="he-foot__cell">
      <div class="he-foot__lbl">TOTAL COMPARECERAM</div>
      <div class="he-foot__val he-foot__val--green">${m.heComp}</div>
    </div>
    <div class="he-foot__cell">
      <div class="he-foot__lbl">TOTAL AUSENTES</div>
      <div class="he-foot__val he-foot__val--red">${m.heAus}</div>
    </div>
    <div class="he-foot__cell">
      <div class="he-foot__lbl">TOTAL HORAS REALIZADAS</div>
      <div class="he-foot__val">${F.h(m.heTrabH)}</div>
    </div>`;
}

/* ─────────────────────────────────────────────────────────────────
   RENDER: INSIGHTS
   ───────────────────────────────────────────────────────────────── */
function renderInsights(m, setorLabel, turnoLabel) {
  const best  = m.porDia.reduce((a,b)=>(b.hTrab/b.hPrev)>(a.hTrab/a.hPrev)?b:a);
  const worst = m.porDia.reduce((a,b)=>(b.hTrab/b.hPrev)<(a.hTrab/a.hPrev)?b:a);
  const totF  = m.porDia.reduce((s,d)=>s+d.faltas,0);
  const totA  = m.porDia.reduce((s,d)=>s+d.atestados,0);
  const totAt = m.porDia.reduce((s,d)=>s+d.atrasos,0);

  document.getElementById('analysis-text').innerHTML = [
    `A presença média da semana foi de ${F.pct(m.avgPct)} — ${setorLabel}${turnoLabel !== 'Todos os Turnos' ? ' / ' + turnoLabel : ''}.`,
    `Melhor desempenho na ${best.dia} (${F.pct(best.hTrab/best.hPrev)}), menor na ${worst.dia} (${F.pct(worst.hTrab/worst.hPrev)}).`,
    `Registradas ${totF} falta(s), ${totA} atestado(s) e ${totAt} atraso(s), resultando em ${F.h(m.hPerd)} de horas perdidas.`,
    `Eficiência da semana: ${F.pct(m.efic)}.`,
  ].map(f => `<li>${f}</li>`).join('');
}

/* ─────────────────────────────────────────────────────────────────
   BUSCA DE COLABORADORES
   ───────────────────────────────────────────────────────────────── */
function normalize(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase();
}

function doSearch(term) {
  const dropdown = document.getElementById('search-dropdown');
  if (!term || term.length < 2) { closeSearch(); return; }

  const norm = normalize(term);
  const results = RAW.colabs.filter(c => {
    if (!c.nome && !c.mat) return false;
    if (normalize(c.nome).includes(norm)) return true;
    if (String(c.mat).startsWith(term)) return true;
    return false;
  }).slice(0, 10);

  if (!results.length) {
    dropdown.innerHTML = `<div class="search-no-result">Nenhum colaborador encontrado</div>`;
  } else {
    dropdown.innerHTML = results.map(c => `
      <div class="search-item" tabindex="0" role="option"
           data-mat="${c.mat}" data-nome="${c.nome}">
        <span class="search-item__mat">${c.mat || 'S/N'}</span>
        <span class="search-item__nome">${c.nome}</span>
        <span class="search-item__setor">${c.turno || ''}</span>
      </div>`).join('');

    dropdown.querySelectorAll('.search-item').forEach(item => {
      item.addEventListener('click', () => {
        const mat = parseInt(item.dataset.mat);
        openColabModal(mat);
        closeSearch();
        document.getElementById('search-colab').value = item.dataset.nome;
      });
      item.addEventListener('keydown', e => {
        if (e.key === 'Enter') item.click();
      });
    });
  }

  dropdown.hidden = false;
}

function closeSearch() {
  const dropdown = document.getElementById('search-dropdown');
  dropdown.hidden = true;
  dropdown.innerHTML = '';
}

/* ─────────────────────────────────────────────────────────────────
   MODAL COLABORADOR — detalhe individual
   ───────────────────────────────────────────────────────────────── */
function openColabModal(mat) {
  const colab = RAW.colabs.find(c => c.mat === mat);
  if (!colab) return;

  const regs  = RAW.registros.filter(r => r.m === mat);
  const overlay = document.getElementById('modal-overlay');

  /* Header */
  document.getElementById('modal-title').textContent = colab.nome;
  document.getElementById('modal-meta').innerHTML = `
    <span class="modal-box__meta-chip">Matrícula: <strong>${colab.mat}</strong></span>
    <span class="modal-box__meta-chip">Setor: <strong>${colab.setor}</strong></span>
    <span class="modal-box__meta-chip">Turno: <strong>${colab.turno}</strong></span>
    <span class="modal-box__meta-chip">Status: <strong>${colab.status}</strong></span>`;

  /* Tabela por dia */
  if (!regs.length) {
    document.getElementById('modal-body').innerHTML =
      '<p style="text-align:center;color:#6a7f9a;padding:20px">Sem registros para este colaborador no período.</p>';
  } else {
    const totalPrev = regs.reduce((s,r)=>s+r.p,0);
    const totalTrab = regs.reduce((s,r)=>s+r.t,0);
    const totalPerd = regs.reduce((s,r)=>s+r.x,0);

    const rows = DIAS_MAP.map(dk => {
      const r = regs.find(r => r.d === dk);
      if (!r) return `<tr><td>${DIAS_LABEL[dk]}</td><td colspan="5" style="color:#aaa;text-align:center">—</td></tr>`;
      const isAlert = r.x > 0 || (r.o && r.o !== '-' && !r.o.includes('HORA EXTRA'));
      const motClass = isAlert ? 'modal-motivo modal-motivo--alert' : 'modal-motivo';
      const mot = r.o && r.o !== '-' ? r.o.replace(/\[LM\]/g,'').trim() : '—';
      return `
        <tr>
          <td>${DIAS_LABEL[dk]}</td>
          <td>${DATAS_MAP[dk]}</td>
          <td>${r.p > 0 ? F.h(r.p) : '—'}</td>
          <td>${r.t > 0 ? F.h(r.t) : '—'}</td>
          <td>${r.x > 0 ? F.h(r.x) : '—'}</td>
          <td><span class="${motClass}" title="${mot}">${mot}</span></td>
        </tr>`;
    }).join('');

    const efic = totalPrev > 0 ? (totalTrab/totalPrev*100).toFixed(1).replace('.',',')+'%' : '—';

    document.getElementById('modal-body').innerHTML = `
      <table class="modal-table">
        <thead>
          <tr>
            <th>DIA</th><th>DATA</th>
            <th>H. PREV.</th><th>H. TRAB.</th><th>H. PERD.</th>
            <th>MOTIVO / OCORRÊNCIA</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
        <tfoot>
          <tr>
            <td>TOTAL</td><td>—</td>
            <td>${F.h(totalPrev)}</td>
            <td>${F.h(totalTrab)}</td>
            <td>${totalPerd>0?F.h(totalPerd):'—'}</td>
            <td>Eficiência: ${efic}</td>
          </tr>
        </tfoot>
      </table>`;
  }

  overlay.hidden = false;
  document.getElementById('modal-close').focus();
}

function closeModal() {
  document.getElementById('modal-overlay').hidden = true;
}

/* ─────────────────────────────────────────────────────────────────
   XLSX UPLOAD — parsing e recarga de dados
   ───────────────────────────────────────────────────────────────── */
function cleanNome(s) {
  if (!s) return null;
  s = String(s).trim();
  const idx = s.indexOf(' - ');
  return idx >= 0 ? s.slice(0, idx).trim() : s;
}

function xlTimeToH(v) {
  if (v == null) return 0.0;
  if (v instanceof Date) {
    return Math.round((v.getUTCHours() + v.getUTCMinutes()/60 + v.getUTCSeconds()/3600) * 10000) / 10000;
  }
  if (typeof v === 'number') {
    const frac = v - Math.floor(v);
    return Math.round(frac * 24 * 10000) / 10000;
  }
  return 0.0;
}

/* Extrai o dia da semana (0=Dom … 6=Sáb) de um valor de data do Excel:
   pode chegar como Date (cellDates), número serial ou string dd/mm/yyyy */
function xlDateWD(v) {
  if (!v) return null;
  if (v instanceof Date) return v.getUTCDay();
  if (typeof v === 'number' && v > 0) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
    return d.getUTCDay();
  }
  if (typeof v === 'string') {
    const p = v.split(/[\/\-\.]/);
    if (p.length === 3) {
      const d = new Date(Date.UTC(+p[2], +p[1] - 1, +p[0]));
      if (!isNaN(d)) return d.getUTCDay();
    }
  }
  return null;
}

function parseXlsxData(buffer) {
  if (typeof XLSX === 'undefined') throw new Error('Biblioteca XLSX não carregada. Verifique a conexão com a internet.');

  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const abas = wb.SheetNames.join(', ');
  console.log('[xlsx] abas encontradas:', abas);

  /* ── 1. Colaboradores ── */
  const ws_c = wb.Sheets['Dados_Colaboradores'];
  if (!ws_c) throw new Error(`Aba "Dados_Colaboradores" não encontrada.\nAbas disponíveis: ${abas}`);
  const rows_c = XLSX.utils.sheet_to_json(ws_c, { header: 1, defval: null });
  const colabs = [];
  const mat_lookup = {};
  for (let i = 1; i < rows_c.length; i++) {
    const r = rows_c[i];
    if (!r[1] && !r[0]) continue;
    if (!r[2]) continue;
    const mat   = r[0] != null ? parseInt(r[0]) : null;
    const nome  = cleanNome(r[1]);
    const setor = r[2] != null ? String(r[2]).trim() : null;
    const turno = r[3] != null ? String(r[3]).trim() : null;
    const status= r[4] != null ? String(r[4]).trim() : null;
    colabs.push({ mat, nome, setor, turno, status });
    if (mat) mat_lookup[mat] = { setor, turno };
  }

  /* ── 2. Registros diários (Seg–Sex) ── */
  const DAY_JS = {1:'Segunda', 2:'Terca', 3:'Quarta', 4:'Quinta', 5:'Sexta'};
  const ws_d = wb.Sheets['dados'];
  if (!ws_d) throw new Error(`Aba "dados" não encontrada.\nAbas disponíveis: ${abas}`);
  const rows_d = XLSX.utils.sheet_to_json(ws_d, { header: 1, cellDates: true, defval: null });
  const registros = [];
  for (let i = 1; i < rows_d.length; i++) {
    const r = rows_d[i];
    if (!r[0] || !r[1]) continue;
    const wd = xlDateWD(r[0]);
    if (wd === null || !DAY_JS[wd]) continue;
    let o = r[7] != null ? String(r[7]).trim() : '';
    if (o === '-' || o === 'None') o = '';
    registros.push({
      m: parseInt(r[1]),
      d: DAY_JS[wd],
      p: xlTimeToH(r[2]),
      t: xlTimeToH(r[3]),
      x: r[8] != null ? Math.round(parseFloat(r[8]) * 10000) / 10000 : 0.0,
      o,
    });
  }

  /* ── 3. HE Sábado (semana mais recente) ── */
  const ws_he = wb.Sheets['HE_SABADO'];
  if (!ws_he) throw new Error(`Aba "HE_SABADO" não encontrada.\nAbas disponíveis: ${abas}`);
  const rows_he_raw = XLSX.utils.sheet_to_json(ws_he, { header: 1, cellDates: true, defval: null });
  const rows_he = rows_he_raw.slice(1).filter(r => r[0] != null);

  /* semana mais recente — compara por timestamp se for Date, senão por valor */
  const toNum = v => v instanceof Date ? v.getTime() : Number(v);
  const sem_max = rows_he.length
    ? rows_he.reduce((mx, r) => toNum(r[0]) > toNum(mx) ? r[0] : mx, rows_he[0][0])
    : null;

  const he_sabado = [];
  for (const r of rows_he) {
    if (toNum(r[0]) !== toNum(sem_max)) continue;
    const mat  = r[1] != null ? parseInt(r[1]) : null;
    const nome = r[2] != null ? String(r[2]).trim() : '';
    const s    = r[4] != null ? String(r[4]).toUpperCase().trim() === 'SIM' : false;
    const h    = r[6] != null ? Math.round(xlTimeToH(r[6]) * 3600) : 0;
    const info = mat_lookup[mat] || {};
    he_sabado.push({ mat, n: nome, setor: info.setor || null, turno: info.turno || null, s, h });
  }

  return { colabs, registros, heSabado: he_sabado };
}

function reloadFromData(newData) {
  const prevSetor = STATE.setor;
  const prevTurno = STATE.turno;

  RAW = newData;

  const novosSetores = new Set(newData.colabs.map(c => c.setor).filter(Boolean));
  const novosTurnos  = new Set(newData.colabs.map(c => c.turno).filter(Boolean));

  STATE = {
    setor: novosSetores.has(prevSetor) ? prevSetor : '',
    turno: novosTurnos.has(prevTurno)  ? prevTurno  : '',
  };

  document.getElementById('filter-setor').innerHTML = '<option value="">Todos os Setores</option>';
  document.getElementById('filter-turno').innerHTML = '<option value="">Todos os Turnos</option>';
  document.getElementById('search-colab').value = '';
  document.getElementById('search-clear').hidden = true;
  closeSearch();
  populateSelects();
  document.getElementById('filter-setor').value = STATE.setor;
  document.getElementById('filter-turno').value = STATE.turno;
  updateActiveTags();
  renderAll();
}

/* ─────────────────────────────────────────────────────────────────
   BOOT
   ───────────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', loadData);

