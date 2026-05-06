import React, { useEffect, useMemo, useState } from 'react';
import ReactDOM from 'react-dom/client';
import {
  Area, AreaChart, Bar, BarChart, CartesianGrid, Cell, Line, LineChart,
  Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis
} from 'recharts';
import * as XLSX from 'xlsx';
import { api } from './api.js';
import './styles.css';

const COLORS = ['#6d71f0', '#8a8ef5', '#c4c6ff', '#30d173', '#ffb84d', '#ff8078', '#a5a1b3'];
const today = () => new Date().toISOString().slice(0, 10);
const monthKey = d => String(d || '').slice(0, 7);
const money = v => 'R$ ' + Number(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const brDate = d => d ? d.split('-').reverse().join('/') : '-';

const METRICS = [
  { v: 'expense', label: 'Saídas' },
  { v: 'income', label: 'Entradas' },
  { v: 'balance', label: 'Saldo' },
  { v: 'count', label: 'Quantidade' }
];
const GROUPS = [
  { v: 'category', label: 'Categoria' },
  { v: 'group', label: 'Grupo' },
  { v: 'account', label: 'Cartão/Conta' },
  { v: 'month', label: 'Mês' },
  { v: 'type', label: 'Tipo' }
];
const TYPES = [
  { v: 'kpi', label: 'KPI' },
  { v: 'bar', label: 'Barras' },
  { v: 'line', label: 'Linha' },
  { v: 'area', label: 'Área' },
  { v: 'pie', label: 'Pizza' }
];
const SIZES = [1, 2, 3, 4, 6, 8, 12];

function metricValue(row, metric) {
  const value = Number(row.amount) || 0;
  if (metric === 'expense') return row.type === 'expense' ? value : 0;
  if (metric === 'income') return row.type === 'income' ? value : 0;
  if (metric === 'balance') return row.type === 'income' ? value : -value;
  if (metric === 'count') return 1;
  return value;
}

function groupName(row, groupBy) {
  if (groupBy === 'category') return row.category || 'Sem categoria';
  if (groupBy === 'group') return row.group || 'Sem grupo';
  if (groupBy === 'account') return row.account || 'Sem cartão/conta';
  if (groupBy === 'month') return monthKey(row.date);
  if (groupBy === 'type') return row.type === 'income' ? 'Entrada' : 'Saída';
  return 'Total';
}

function aggregate(rows, metric) {
  return rows.reduce((sum, r) => sum + metricValue(r, metric), 0);
}

function grouped(rows, metric, groupBy) {
  const map = new Map();
  for (const row of rows) {
    const name = groupName(row, groupBy);
    map.set(name, (map.get(name) || 0) + metricValue(row, metric));
  }
  return [...map.entries()].map(([name, value]) => ({ name, value }));
}

function colorForCategory(categories, name, fallback = '#6d71f0') {
  return categories.find(c => c.name === name)?.color || fallback;
}

function Login({ onLogin }) {
  const [form, setForm] = useState({ code: '', password: '' });
  const [error, setError] = useState('');
  async function submit(e) {
    e.preventDefault();
    try {
      const data = await api.post('/login', form);
      api.setAuth(data.token, data.user);
      onLogin(data.user);
    } catch (err) {
      setError(err.message);
    }
  }
  return (
    <main className="login-bg">
      <form className="glass login-card" onSubmit={submit}>
        <div className="login-brand">Finans</div>
        <div className="login-sub">Finanças pessoais</div>
        <div className="field"><label className="label">Código</label><input className="input" value={form.code} onChange={e => setForm({ ...form, code: e.target.value })} /></div>
        <div className="field"><label className="label">Senha</label><input className="input" type="password" value={form.password} onChange={e => setForm({ ...form, password: e.target.value })} /></div>
        {error && <div className="error-msg">{error}</div>}
        <button className="btn accent">Entrar</button>
      </form>
    </main>
  );
}

function App() {
  const [user, setUser] = useState(api.user);
  const [tab, setTab] = useState('dashboard');
  const readOnly = user?.role === 'user';
  useEffect(() => applySavedTheme(), []);
  if (!user) return <Login onLogin={setUser} />;
  function logout() {
    api.setAuth(null, null);
    setUser(null);
  }
  return (
    <>
      <nav className="navbar">
        <div className="brand"><div className="brand-icon"><span /><span /><span /></div>Finans</div>
        <div className="nav-links">
          {(readOnly ? [
            ['dashboard', 'Dashboards']
          ] : [
            ['dashboard', 'Dashboards'],
            ['transacoes', 'Transações'],
            ['dividas', 'Dívidas'],
            ['dividasFixas', 'Dívidas fixas'],
            ['mercado', 'Mercado'],
            ['config', 'Configurações'],
            ['usuarios', 'Usuários']
          ]).map(([key, label]) => <button key={key} className={`nav-link ${tab === key ? 'active' : ''}`} onClick={() => setTab(key)}>{label}</button>)}
        </div>
        <div className="nav-user"><span>{user.name}</span><button onClick={logout}>Sair</button></div>
      </nav>
      <main className="container">
        {tab === 'dashboard' && <Dashboard readOnly={readOnly} />}
        {!readOnly && tab === 'transacoes' && <Transactions />}
        {!readOnly && tab === 'dividas' && <Debts />}
        {!readOnly && tab === 'dividasFixas' && <FixedDebts />}
        {!readOnly && tab === 'mercado' && <Market />}
        {!readOnly && tab === 'config' && <Settings />}
        {!readOnly && tab === 'usuarios' && <Users />}
      </main>
    </>
  );
}

function Dashboard({ readOnly = false }) {
  const [rows, setRows] = useState([]);
  const [categories, setCategories] = useState([]);
  const [groups, setGroups] = useState([]);
  const [accounts, setAccounts] = useState([]);
  const [dashboards, setDashboards] = useState([]);
  const [active, setActive] = useState(null);
  const [widgets, setWidgets] = useState([]);
  const [editing, setEditing] = useState(null);
  const [showWidget, setShowWidget] = useState(false);
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [filters, setFilters] = useState({ month: '', week: '', from: '', to: '', type: '', category: '', group: '', account: '' });
  const filteredRows = useMemo(() => filterRows(rows, filters), [rows, filters]);
  const accountNames = useMemo(() => [...new Set([...accounts.map(a => a.name), ...uniqueOptions(rows, 'account')].filter(Boolean))], [accounts, rows]);

  async function loadDashboards() {
    const list = await api.get('/dashboards');
    setDashboards(list);
    setActive(current => current && list.find(d => d.id === current.id) ? current : list[0] || null);
    return list;
  }
  async function load() {
    const list = dashboards.length ? dashboards : await loadDashboards();
    const dash = active || list[0];
    const [tx, ws, cats, gs, accs] = await Promise.all([
      api.get('/transactions'),
      dash ? api.get(`/widgets?dashboard_id=${dash.id}`) : [],
      api.get('/categories'),
      api.get('/groups'),
      api.get('/accounts')
    ]);
    setRows(tx);
    setWidgets(ws);
    setCategories(cats);
    setGroups(gs);
    setAccounts(accs);
  }
  useEffect(() => { loadDashboards(); }, []);
  useEffect(() => { if (active) load(); }, [active?.id]);

  async function createDashboard() {
    const title = prompt('Nome da dashboard?') || 'Nova dashboard';
    const d = await api.post('/dashboards', { title });
    await loadDashboards();
    setActive(d);
  }
  async function renameDashboard() {
    const title = prompt('Novo nome?', active?.title);
    if (!title) return;
    const d = await api.put(`/dashboards/${active.id}`, { title });
    setActive(d);
    loadDashboards();
  }
  async function deleteDashboard() {
    if (!active || dashboards.length <= 1) return;
    if (!confirm('Excluir dashboard?')) return;
    await api.del(`/dashboards/${active.id}`);
    await loadDashboards();
  }
  async function saveWidget(w) {
    const payload = { ...w, dashboard_id: active?.id };
    if (editing) await api.put(`/widgets/${editing.id}`, payload);
    else await api.post('/widgets', payload);
    setEditing(null);
    setShowWidget(false);
    load();
  }
  async function deleteWidget(id) {
    if (!confirm('Excluir widget?')) return;
    await api.del(`/widgets/${id}`);
    load();
  }

  return (
    <div>
      <div className="page-header">
        <div><h1>{active?.title || 'Dashboards'}</h1><div className="subtitle">Resumo financeiro e gráficos personalizáveis</div></div>
        <div className="row-flex">
          <button className="btn" onClick={() => setFiltersOpen(true)}>Filtros</button>
          {!readOnly && <button className="btn" onClick={renameDashboard} disabled={!active}>Renomear</button>}
          {!readOnly && <button className="btn danger" onClick={deleteDashboard} disabled={dashboards.length <= 1}>Excluir página</button>}
          {!readOnly && <button className="btn accent" onClick={() => setShowWidget(true)}>+ Widget</button>}
        </div>
      </div>
      <div className="dash-pages">
        {dashboards.map(d => <button key={d.id} className={`range-pill ${active?.id === d.id ? 'active' : ''}`} onClick={() => setActive(d)}>{d.title}</button>)}
        {!readOnly && <button className="range-pill add" onClick={createDashboard}>NewTab</button>}
      </div>
      {widgets.length === 0 ? (
        <div className="glass empty-state"><h3>Nenhum widget ainda</h3>{!readOnly && <button className="btn accent mt-2" onClick={() => setShowWidget(true)}>+ Criar widget</button>}</div>
      ) : (
        <div className="widgets-grid">
          {widgets.map(w => <WidgetCard key={w.id} widget={w} rows={filteredRows} categories={categories} onEdit={readOnly ? null : () => setEditing(w)} onDelete={readOnly ? null : () => deleteWidget(w.id)} />)}
        </div>
      )}
      {filtersOpen && <FilterDrawer filters={filters} setFilters={setFilters} categories={categories} groups={groups} accounts={accountNames} onClose={() => setFiltersOpen(false)} />}
      {(showWidget || editing) && <WidgetEditor initial={editing} onClose={() => { setEditing(null); setShowWidget(false); }} onSave={saveWidget} />}
    </div>
  );
}

function WidgetCard({ widget, rows, categories, onEdit, onDelete }) {
  const [hidden, setHidden] = useState(false);
  const data = useMemo(() => grouped(rows, widget.metric, widget.group_by), [rows, widget]);
  const value = aggregate(rows, widget.metric);
  const moneyMetric = widget.metric !== 'count';
  const formatted = moneyMetric ? money(value) : value.toLocaleString('pt-BR');
  return (
    <div className={`widget size-${widget.size || 4}`}>
      <div className="widget-head">
        <div className="widget-title">{widget.title}</div>
        <div className="widget-actions">
          <button title={hidden ? 'Mostrar dados' : 'Ocultar dados'} onClick={() => setHidden(v => !v)}>{hidden ? '◌' : '●'}</button>
          {onEdit && <button onClick={onEdit}>✎</button>}
          {onDelete && <button onClick={onDelete}>×</button>}
        </div>
      </div>
      {widget.chart_type === 'kpi' ? (
        <div className="widget-kpi"><div className="value">{hidden ? '••••' : formatted}</div><div className="label">{METRICS.find(m => m.v === widget.metric)?.label}</div></div>
      ) : (
        <div className="widget-chart"><Chart type={widget.chart_type} data={data} color={widget.color} hidden={hidden} categories={categories} groupBy={widget.group_by} /></div>
      )}
    </div>
  );
}

function Chart({ type, data, color, hidden, categories, groupBy }) {
  const tooltip = hidden ? null : <Tooltip formatter={v => money(v)} contentStyle={{ background: '#141415', border: '1px solid rgba(255,255,255,.1)', borderRadius: 10 }} />;
  if (type === 'pie') return (
    <ResponsiveContainer width="100%" height={220}><PieChart><Pie data={data} dataKey="value" nameKey="name" outerRadius={80} label={hidden ? false : { fill: '#acadb1' }}>{data.map((d, i) => <Cell key={i} fill={groupBy === 'category' ? colorForCategory(categories, d.name, COLORS[i % COLORS.length]) : COLORS[i % COLORS.length]} />)}</Pie>{tooltip}</PieChart></ResponsiveContainer>
  );
  if (type === 'line') return (
    <ResponsiveContainer width="100%" height={220}><LineChart data={data}><CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.06)" /><XAxis dataKey="name" stroke="#acadb1" /><YAxis tick={hidden ? false : { fill: '#acadb1' }} stroke="#acadb1" />{tooltip}<Line dataKey="value" stroke={color} strokeWidth={2} /></LineChart></ResponsiveContainer>
  );
  if (type === 'area') return (
    <ResponsiveContainer width="100%" height={220}><AreaChart data={data}><CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.06)" /><XAxis dataKey="name" stroke="#acadb1" /><YAxis tick={hidden ? false : { fill: '#acadb1' }} stroke="#acadb1" />{tooltip}<Area dataKey="value" stroke={color} fill={color} fillOpacity={0.25} /></AreaChart></ResponsiveContainer>
  );
  return (
    <ResponsiveContainer width="100%" height={220}><BarChart data={data}><CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.06)" /><XAxis dataKey="name" stroke="#acadb1" /><YAxis tick={hidden ? false : { fill: '#acadb1' }} stroke="#acadb1" />{tooltip}<Bar dataKey="value" fill={color} radius={[6, 6, 0, 0]} minPointSize={1}>{data.map((d, i) => <Cell key={i} fill={groupBy === 'category' ? colorForCategory(categories, d.name, color) : color} />)}</Bar></BarChart></ResponsiveContainer>
  );
}

function WidgetEditor({ initial, onClose, onSave }) {
  const [w, setW] = useState(initial || { title: 'Novo widget', chart_type: 'kpi', metric: 'expense', group_by: 'category', color: '#6d71f0', size: 4 });
  const set = (k, v) => setW({ ...w, [k]: v });
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <div className="modal-header"><h2>{initial ? 'Editar widget' : 'Novo widget'}</h2><button className="modal-close" onClick={onClose}>×</button></div>
        <div className="field"><label className="label">Título</label><input className="input" value={w.title} onChange={e => set('title', e.target.value)} /></div>
        <div className="grid-2">
          <div className="field"><label className="label">Métrica</label><select className="select" value={w.metric} onChange={e => set('metric', e.target.value)}>{METRICS.map(m => <option key={m.v} value={m.v}>{m.label}</option>)}</select></div>
          <div className="field"><label className="label">Agrupar por</label><select className="select" value={w.group_by} onChange={e => set('group_by', e.target.value)}>{GROUPS.map(g => <option key={g.v} value={g.v}>{g.label}</option>)}</select></div>
          <div className="field"><label className="label">Tipo</label><select className="select" value={w.chart_type} onChange={e => set('chart_type', e.target.value)}>{TYPES.map(t => <option key={t.v} value={t.v}>{t.label}</option>)}</select></div>
          <div className="field"><label className="label">Tamanho</label><select className="select" value={w.size} onChange={e => set('size', Number(e.target.value))}>{SIZES.map(s => <option key={s} value={s}>{s}</option>)}</select></div>
        </div>
        <div className="field"><label className="label">Cor</label><div className="color-picker">{COLORS.map(c => <button type="button" key={c} className={`color-dot ${w.color === c ? 'selected' : ''}`} style={{ background: c }} onClick={() => set('color', c)} />)}</div></div>
        <div className="modal-actions"><button className="btn ghost" onClick={onClose}>Cancelar</button><button className="btn accent" onClick={() => onSave(w)}>Salvar</button></div>
      </div>
    </div>
  );
}

function Transactions() {
  const [rows, setRows] = useState([]);
  const [groups, setGroups] = useState([]);
  const [categories, setCategories] = useState([]);
  const [accounts, setAccounts] = useState([]);
  const [groupFilter, setGroupFilter] = useState('');
  const [form, setForm] = useState({ date: today(), type: 'expense', description: '', category: '', group: '', account: '', amount: '' });
  const [editing, setEditing] = useState(null);
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [filters, setFilters] = useState({ month: '', week: '', from: '', to: '', type: '', category: '', group: '', account: '' });
  async function load() {
    const [tx, gs, cats, accs] = await Promise.all([api.get('/transactions'), api.get('/groups'), api.get('/categories'), api.get('/accounts')]);
    setRows(tx);
    setGroups(gs);
    setCategories(cats);
    setAccounts(accs);
  }
  useEffect(() => { load(); }, []);
  const categoryOptions = useMemo(() => mergeNamed(categories, uniqueOptions(rows, 'category')), [categories, rows]);
  const accountOptions = useMemo(() => mergeNamed(accounts, uniqueOptions(rows, 'account')), [accounts, rows]);
  const visibleRows = useMemo(() => filterRows(groupFilter ? rows.filter(r => r.group === groupFilter) : rows, filters), [rows, groupFilter, filters]);
  async function save(e) {
    e.preventDefault();
    const payload = { ...form, amount: Number(form.amount) || 0 };
    if (editing) await api.put('/transactions/' + editing.id, payload);
    else await api.post('/transactions', payload);
    setEditing(null);
    setForm({ date: today(), type: 'expense', description: '', category: '', group: '', account: '', amount: '' });
    load();
  }
  function editRow(r) {
    setEditing(r);
    setForm({ date: r.date, type: r.type, description: r.description, category: r.category, group: r.group || '', account: r.account, amount: r.amount });
  }
  async function delRow(r) {
    if (!confirm('Excluir transação?')) return;
    await api.del('/transactions/' + r.id);
    load();
  }
  async function exportXLS() {
    const ws = XLSX.utils.json_to_sheet(visibleRows.map(r => ({ Data: r.date, Tipo: r.type, Descrição: r.description, Categoria: r.category, Grupo: r.group, 'Cartão/Conta': r.account, Valor: r.amount })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Transações');
    XLSX.writeFile(wb, 'finans_transacoes.xlsx');
  }
  return (
    <div>
      <div className="page-header"><div><h1>Transações</h1><div className="subtitle">Entradas, saídas, cartões, contas e categorias</div></div><div className="row-flex"><button className="btn" onClick={() => setFiltersOpen(true)}>Filtros</button><button className="btn" onClick={exportXLS}>↓ Excel</button></div></div>
      <section className="glass mb-2">
        <div className="label mb-2">Filtro rápido</div>
        <div className="row-flex mb-2">
          <button className={`range-pill ${!groupFilter ? 'active' : ''}`} onClick={() => setGroupFilter('')}>Todos</button>
          {groups.map(g => <button key={g.id} className={`range-pill ${groupFilter === g.name ? 'active' : ''}`} onClick={() => setGroupFilter(g.name)}><span className="dot" style={{ background: g.color }} />{g.name}</button>)}
        </div>
      </section>
      <form className="glass form-grid" onSubmit={save}>
        <input className="input" type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
        <select className="select" value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}><option value="expense">Saída</option><option value="income">Entrada</option></select>
        <input className="input" placeholder="Descrição" value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} />
        <select className="select" value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}><option value="">Categoria</option>{categoryOptions.map(c => <option key={c.name} value={c.name}>{c.name}</option>)}</select>
        <select className="select" value={form.group} onChange={e => setForm({ ...form, group: e.target.value })}><option value="">Grupo</option>{groups.map(g => <option key={g.id} value={g.name}>{g.name}</option>)}</select>
        <select className="select" value={form.account} onChange={e => setForm({ ...form, account: e.target.value })}><option value="">Conta/Cartão</option>{accountOptions.map(a => <option key={a.name} value={a.name}>{a.name}</option>)}</select>
        <input className="input" placeholder="Valor" type="number" step="0.01" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} />
        <button className="btn accent">{editing ? 'Atualizar' : 'Salvar'}</button>
        {editing && <button type="button" className="btn ghost" onClick={() => { setEditing(null); setForm({ date: today(), type: 'expense', description: '', category: '', group: '', account: '', amount: '' }); }}>Cancelar</button>}
      </form>
      <section className="glass table-panel">
        <table><thead><tr><th>Data</th><th>Tipo</th><th>Descrição</th><th>Categoria</th><th>Grupo</th><th>Cartão/Conta</th><th>Valor</th><th></th></tr></thead>
        <tbody>{visibleRows.map(r => <tr key={r.id}><td>{brDate(r.date)}</td><td>{r.type === 'income' ? 'Entrada' : 'Saída'}</td><td>{r.description}</td><td>{r.category}</td><td>{r.group}</td><td>{r.account}</td><td>{money(r.amount)}</td><td><div className="table-actions"><button className="btn sm" onClick={() => editRow(r)}>Editar</button><button className="btn sm danger" onClick={() => delRow(r)}>×</button></div></td></tr>)}</tbody></table>
      </section>
      {filtersOpen && <FilterDrawer filters={filters} setFilters={setFilters} categories={categoryOptions} groups={groups} accounts={accountOptions.map(a => a.name)} onClose={() => setFiltersOpen(false)} />}
    </div>
  );
}

function Debts() {
  const [rows, setRows] = useState([]);
  const [groups, setGroups] = useState([]);
  const [categories, setCategories] = useState([]);
  const [accounts, setAccounts] = useState([]);
  const [form, setForm] = useState({ date: today(), due_date: '', creditor: '', description: '', category: '', group: '', account: '', amount: '' });
  const [editing, setEditing] = useState(null);
  const [paying, setPaying] = useState(null);
  const [payForm, setPayForm] = useState({ date: today(), account: '', amount: '' });
  async function load() {
    const [ds, gs, cats, accs] = await Promise.all([api.get('/debts'), api.get('/groups'), api.get('/categories'), api.get('/accounts')]);
    setRows(ds);
    setGroups(gs);
    setCategories(cats);
    setAccounts(accs);
  }
  useEffect(() => { load(); }, []);
  async function save(e) {
    e.preventDefault();
    const payload = { ...form, amount: Number(form.amount) || 0 };
    if (editing) await api.put('/debts/' + editing.id, payload);
    else await api.post('/debts', payload);
    setEditing(null);
    setForm({ date: today(), due_date: '', creditor: '', description: '', category: '', group: '', account: '', amount: '' });
    load();
  }
  function edit(row) {
    setEditing(row);
    setForm({ date: row.date, due_date: row.due_date || '', creditor: row.creditor || '', description: row.description || '', category: row.category || '', group: row.group || '', account: row.account || '', amount: row.amount });
  }
  async function del(row) {
    if (!confirm('Excluir dívida?')) return;
    await api.del('/debts/' + row.id);
    load();
  }
  function openPay(row) {
    setPaying(row);
    setPayForm({ date: today(), account: row.account || '', amount: row.amount });
  }
  async function pay(e) {
    e.preventDefault();
    await api.post('/debts/' + paying.id + '/pay', { ...payForm, amount: Number(payForm.amount) || 0 });
    setPaying(null);
    load();
  }
  const open = rows.filter(r => !r.paid);
  const paid = rows.filter(r => r.paid);
  return (
    <div>
      <div className="page-header"><div><h1>Dívidas</h1><div className="subtitle">Valores pendentes e pagamentos vinculados</div></div></div>
      <form className="glass form-grid debt-form" onSubmit={save}>
        <input className="input" type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
        <input className="input" type="date" value={form.due_date} onChange={e => setForm({ ...form, due_date: e.target.value })} />
        <input className="input" placeholder="Credor" value={form.creditor} onChange={e => setForm({ ...form, creditor: e.target.value })} />
        <input className="input" placeholder="Descrição" value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} />
        <select className="select" value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}><option value="">Tag</option>{categories.map(c => <option key={c.id} value={c.name}>{c.name}</option>)}</select>
        <select className="select" value={form.group} onChange={e => setForm({ ...form, group: e.target.value })}><option value="">Grupo</option>{groups.map(g => <option key={g.id} value={g.name}>{g.name}</option>)}</select>
        <select className="select" value={form.account} onChange={e => setForm({ ...form, account: e.target.value })}><option value="">Conta prevista</option>{accounts.map(a => <option key={a.id} value={a.name}>{a.name}</option>)}</select>
        <input className="input" placeholder="Valor" type="number" step="0.01" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} />
        <button className="btn accent">{editing ? 'Atualizar' : 'Salvar dívida'}</button>
        {editing && <button type="button" className="btn ghost" onClick={() => { setEditing(null); setForm({ date: today(), due_date: '', creditor: '', description: '', category: '', group: '', account: '', amount: '' }); }}>Cancelar</button>}
      </form>
      <DebtTable title="Em aberto" rows={open} onEdit={edit} onDelete={del} onPay={openPay} />
      <DebtTable title="Pagas" rows={paid} onEdit={edit} onDelete={del} />
      {paying && (
        <div className="modal-backdrop" onClick={() => setPaying(null)}>
          <form className="modal" onSubmit={pay} onClick={e => e.stopPropagation()}>
            <div className="modal-header"><h2>Pagar dívida</h2><button type="button" className="modal-close" onClick={() => setPaying(null)}>×</button></div>
            <div className="field"><label className="label">Data pagamento</label><input className="input" type="date" value={payForm.date} onChange={e => setPayForm({ ...payForm, date: e.target.value })} /></div>
            <div className="field"><label className="label">Conta/cartão</label><select className="select" value={payForm.account} onChange={e => setPayForm({ ...payForm, account: e.target.value })}><option value="">Conta</option>{accounts.map(a => <option key={a.id} value={a.name}>{a.name}</option>)}</select></div>
            <div className="field"><label className="label">Valor</label><input className="input" type="number" step="0.01" value={payForm.amount} onChange={e => setPayForm({ ...payForm, amount: e.target.value })} /></div>
            <div className="modal-actions"><button className="btn ghost" type="button" onClick={() => setPaying(null)}>Cancelar</button><button className="btn accent">Confirmar</button></div>
          </form>
        </div>
      )}
    </div>
  );
}

function DebtTable({ title, rows, onEdit, onDelete, onPay }) {
  return (
    <section className="glass table-panel mb-2">
      <div className="table-title">{title}</div>
      <table><thead><tr><th>Data</th><th>Vencimento</th><th>Credor</th><th>Descrição</th><th>Grupo</th><th>Conta</th><th>Valor</th><th></th></tr></thead>
      <tbody>{rows.map(r => <tr key={r.id}><td>{brDate(r.date)}</td><td>{brDate(r.due_date)}</td><td>{r.creditor}</td><td>{r.description}</td><td>{r.group}</td><td>{r.account}</td><td>{money(r.amount)}</td><td><div className="table-actions">{onPay && <button className="btn sm accent" onClick={() => onPay(r)}>Pagar</button>}<button className="btn sm" onClick={() => onEdit(r)}>Editar</button><button className="btn sm danger" onClick={() => onDelete(r)}>×</button></div></td></tr>)}</tbody></table>
    </section>
  );
}

function FixedDebts() {
  const [rows, setRows] = useState([]);
  const [groups, setGroups] = useState([]);
  const [categories, setCategories] = useState([]);
  const [accounts, setAccounts] = useState([]);
  const [form, setForm] = useState({ start_date: today(), months: 12, creditor: '', description: '', category: '', group: '', account: '', amount: '' });
  const [editing, setEditing] = useState(null);
  async function load() {
    const [fs, gs, cats, accs] = await Promise.all([api.get('/fixed-debts'), api.get('/groups'), api.get('/categories'), api.get('/accounts')]);
    setRows(fs);
    setGroups(gs);
    setCategories(cats);
    setAccounts(accs);
  }
  useEffect(() => { load(); }, []);
  async function save(e) {
    e.preventDefault();
    const payload = { ...form, amount: Number(form.amount) || 0, months: Math.max(1, Number(form.months) || 1) };
    if (editing) await api.put('/fixed-debts/' + editing.id, payload);
    else await api.post('/fixed-debts', payload);
    setEditing(null);
    setForm({ start_date: today(), months: 12, creditor: '', description: '', category: '', group: '', account: '', amount: '' });
    load();
  }
  function edit(row) {
    setEditing(row);
    setForm({ start_date: row.start_date || today(), months: row.months || 1, creditor: row.creditor || '', description: row.description || '', category: row.category || '', group: row.group || '', account: row.account || '', amount: row.amount });
  }
  async function del(row) {
    if (!confirm('Excluir dívida fixa e parcelas em aberto?')) return;
    await api.del('/fixed-debts/' + row.id);
    load();
  }
  return (
    <div>
      <div className="page-header"><div><h1>Dívidas fixas</h1><div className="subtitle">Parcelas mensais geradas em dívidas</div></div></div>
      <form className="glass form-grid fixed-debt-form" onSubmit={save}>
        <input className="input" type="date" value={form.start_date} onChange={e => setForm({ ...form, start_date: e.target.value })} />
        <input className="input" placeholder="Meses" type="number" min="1" max="240" value={form.months} onChange={e => setForm({ ...form, months: e.target.value })} />
        <input className="input" placeholder="Credor" value={form.creditor} onChange={e => setForm({ ...form, creditor: e.target.value })} />
        <input className="input" placeholder="Descrição" value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} />
        <select className="select" value={form.category} onChange={e => setForm({ ...form, category: e.target.value })}><option value="">Tag</option>{categories.map(c => <option key={c.id} value={c.name}>{c.name}</option>)}</select>
        <select className="select" value={form.group} onChange={e => setForm({ ...form, group: e.target.value })}><option value="">Grupo</option>{groups.map(g => <option key={g.id} value={g.name}>{g.name}</option>)}</select>
        <select className="select" value={form.account} onChange={e => setForm({ ...form, account: e.target.value })}><option value="">Conta/cartão</option>{accounts.map(a => <option key={a.id} value={a.name}>{a.name}</option>)}</select>
        <input className="input" placeholder="Valor mensal" type="number" step="0.01" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} />
        <button className="btn accent">{editing ? 'Atualizar' : 'Gerar parcelas'}</button>
        {editing && <button type="button" className="btn ghost" onClick={() => { setEditing(null); setForm({ start_date: today(), months: 12, creditor: '', description: '', category: '', group: '', account: '', amount: '' }); }}>Cancelar</button>}
      </form>
      <section className="glass table-panel">
        <table><thead><tr><th>Início</th><th>Meses</th><th>Credor</th><th>Descrição</th><th>Grupo</th><th>Conta/cartão</th><th>Valor/mês</th><th>Total</th><th></th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td>{brDate(r.start_date)}</td><td>{r.months}</td><td>{r.creditor}</td><td>{r.description}</td><td>{r.group}</td><td>{r.account}</td><td>{money(r.amount)}</td><td>{money(Number(r.amount) * Number(r.months || 0))}</td><td><div className="table-actions"><button className="btn sm" onClick={() => edit(r)}>Editar</button><button className="btn sm danger" onClick={() => del(r)}>×</button></div></td></tr>)}</tbody></table>
      </section>
    </div>
  );
}

function Market() {
  const [rows, setRows] = useState([]);
  const [payments, setPayments] = useState([]);
  const [groups, setGroups] = useState([]);
  const [accounts, setAccounts] = useState([]);
  const [form, setForm] = useState({ date: today(), store: '', total: '', url: '', items: [], payments: [] });
  const [editing, setEditing] = useState(null);
  const [details, setDetails] = useState(null);
  const [scanOpen, setScanOpen] = useState(false);
  const [error, setError] = useState('');
  async function load() {
    const [rs, ps, gs, accs] = await Promise.all([api.get('/receipts'), api.get('/payment-methods'), api.get('/groups'), api.get('/accounts')]);
    setRows(rs);
    setPayments(ps);
    setGroups(gs);
    setAccounts(accs);
  }
  useEffect(() => { load(); }, []);
  async function parseUrl(url) {
    setError('');
    try {
      const data = await api.post('/receipts/parse', { url });
      setForm({ date: data.date || today(), store: data.store || '', total: data.total || '', url, items: data.items || [], payments: data.payments || [] });
      setScanOpen(false);
    } catch (err) {
      setError(err.message);
      setForm(f => ({ ...f, url }));
    }
  }
  async function save(e) {
    e.preventDefault();
    const payload = { ...form, total: Number(form.total) || 0 };
    if (editing) await api.put('/receipts/' + editing.id, payload);
    else await api.post('/receipts', payload);
    setEditing(null);
    setForm({ date: today(), store: '', total: '', url: '', items: [], payments: [] });
    load();
  }
  function edit(row) {
    setEditing(row);
    setForm({ date: row.date, store: row.store, total: row.total, subtotal: row.subtotal || row.total, discount: row.discount || 0, url: row.url || '', items: row.items || [], payments: row.payments || [] });
  }
  async function del(row) {
    if (!confirm('Excluir compra?')) return;
    await api.del('/receipts/' + row.id);
    load();
  }
  function setPayment(idx, patch) {
    setForm(f => ({ ...f, payments: f.payments.map((p, i) => i === idx ? { ...p, ...patch } : p) }));
  }
  function addPayment() {
    setForm(f => ({ ...f, payments: [...(f.payments || []), { method: payments[0]?.name || '', amount: '', account: '', group: '' }] }));
  }
  function removePayment(idx) {
    setForm(f => ({ ...f, payments: f.payments.filter((_, i) => i !== idx) }));
  }
  return (
    <div>
      <div className="page-header"><div><h1>Mercado</h1><div className="subtitle">Compras por NFC-e, mercado e itens</div></div><button className="btn accent" onClick={() => setScanOpen(true)}>Ler QR</button></div>
      {error && <div className="error-msg mb-2">{error}</div>}
      <form className="glass form-grid market-form" onSubmit={save}>
        <input className="input" type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
        <input className="input" placeholder="Mercado" value={form.store} onChange={e => setForm({ ...form, store: e.target.value })} />
        <input className="input" placeholder="Valor total" type="number" step="0.01" value={form.total} onChange={e => setForm({ ...form, total: e.target.value })} />
        <input className="input" placeholder="URL NFC-e" value={form.url} onChange={e => setForm({ ...form, url: e.target.value })} />
        <button type="button" className="btn" onClick={() => parseUrl(form.url)}>Buscar nota</button>
        <button className="btn accent">{editing ? 'Atualizar' : 'Salvar compra'}</button>
        {editing && <button type="button" className="btn ghost" onClick={() => { setEditing(null); setForm({ date: today(), store: '', total: '', url: '', items: [], payments: [] }); }}>Cancelar</button>}
      </form>
      <section className="glass mb-2">
        <div className="settings-head mb-2"><div className="label">Pagamentos</div><button className="btn sm" onClick={addPayment}>+ Forma</button></div>
        <div className="payment-list">
          {(form.payments || []).map((p, idx) => (
            <div className="payment-row" key={idx}>
              <select className="select" value={p.method} onChange={e => setPayment(idx, { method: e.target.value })}>{mergeNamed(payments, [p.method]).map(x => <option key={x.name} value={x.name}>{x.name}</option>)}</select>
              <select className="select" value={p.account || ''} onChange={e => setPayment(idx, { account: e.target.value })}><option value="">Conta/cartão</option>{mergeNamed(accounts, p.account ? [p.account] : []).map(x => <option key={x.name} value={x.name}>{x.name}</option>)}</select>
              <select className="select" value={p.group || ''} onChange={e => setPayment(idx, { group: e.target.value })}><option value="">Grupo</option>{mergeNamed(groups, p.group ? [p.group] : []).map(x => <option key={x.name} value={x.name}>{x.name}</option>)}</select>
              <input className="input" type="number" step="0.01" value={p.amount} onChange={e => setPayment(idx, { amount: e.target.value })} />
              <button className="btn sm danger" onClick={() => removePayment(idx)}>×</button>
            </div>
          ))}
        </div>
      </section>
      <section className="glass table-panel">
        <table><thead><tr><th>Data</th><th>Mercado</th><th>Valor</th><th>Itens</th><th></th></tr></thead>
        <tbody>{rows.map(r => <tr key={r.id}><td>{brDate(r.date)}</td><td>{r.store}</td><td>{money(r.total)}</td><td>{r.items?.length || 0}</td><td><div className="table-actions"><button className="btn sm" onClick={() => setDetails(r)}>Detalhes</button><button className="btn sm" onClick={() => edit(r)}>Editar</button><button className="btn sm danger" onClick={() => del(r)}>×</button></div></td></tr>)}</tbody></table>
      </section>
      {scanOpen && <QrScanner onClose={() => setScanOpen(false)} onResult={parseUrl} />}
      {details && <ReceiptDetails receipt={details} onClose={() => setDetails(null)} />}
    </div>
  );
}

function QrScanner({ onClose, onResult }) {
  const videoRef = React.useRef(null);
  const [msg, setMsg] = useState('Aponte para QR Code');
  useEffect(() => {
    let stream;
    let timer;
    async function start() {
      try {
        if (!('BarcodeDetector' in window)) {
          setMsg('Navegador sem leitor nativo. Cole URL NFC-e.');
          return;
        }
        stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
        videoRef.current.srcObject = stream;
        await videoRef.current.play();
        const detector = new BarcodeDetector({ formats: ['qr_code'] });
        timer = setInterval(async () => {
          const codes = await detector.detect(videoRef.current).catch(() => []);
          if (codes[0]?.rawValue) {
            clearInterval(timer);
            stream.getTracks().forEach(t => t.stop());
            onResult(codes[0].rawValue);
          }
        }, 500);
      } catch (err) {
        setMsg(err.message);
      }
    }
    start();
    return () => {
      if (timer) clearInterval(timer);
      if (stream) stream.getTracks().forEach(t => t.stop());
    };
  }, []);
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()}>
        <div className="modal-header"><h2>Ler QR NFC-e</h2><button className="modal-close" onClick={onClose}>×</button></div>
        <video ref={videoRef} className="qr-video" muted playsInline />
        <div className="subtitle mt-2">{msg}</div>
      </div>
    </div>
  );
}

function ReceiptDetails({ receipt, onClose }) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal wide" onClick={e => e.stopPropagation()}>
        <div className="modal-header"><h2>{receipt.store}</h2><button className="modal-close" onClick={onClose}>×</button></div>
        <div className="quick-grid details-summary">
          <div className="glass-sm"><label>Data</label><strong>{brDate(receipt.date)}</strong></div>
          <div className="glass-sm"><label>Pago</label><strong>{money(receipt.total)}</strong></div>
          <div className="glass-sm"><label>Subtotal</label><strong>{money(receipt.subtotal || receipt.total)}</strong></div>
          <div className="glass-sm"><label>Desconto</label><strong>{money(receipt.discount || 0)}</strong></div>
        </div>
        {(receipt.payments || []).length > 0 && <div className="category-list mb-2">{receipt.payments.map((p, idx) => <div className="config-row" key={idx}><span className="badge">{p.method}{p.account ? ` · ${p.account}` : ''}{p.group ? ` · ${p.group}` : ''}</span><strong>{money(p.amount)}</strong></div>)}</div>}
        <div className="table-panel">
          <table><thead><tr><th>Item</th><th>Qtd</th><th>Unit.</th><th>Total</th></tr></thead>
          <tbody>{(receipt.items || []).map((i, idx) => <tr key={idx}><td>{i.name}</td><td>{i.qty || '-'}</td><td>{i.unit ? money(i.unit) : '-'}</td><td>{money(i.total)}</td></tr>)}</tbody></table>
        </div>
      </div>
    </div>
  );
}

function Settings() {
  return (
    <div>
      <div className="page-header"><div><h1>Configurações</h1><div className="subtitle">Cadastros usados nos filtros e transações</div></div></div>
      <AppearanceManager />
      <div className="settings-grid">
        <ConfigManager title="Tags" endpoint="/categories" emptyName="Nova tag" />
        <ConfigManager title="Grupos" endpoint="/groups" emptyName="Novo grupo" />
        <ConfigManager title="Contas e cartões" endpoint="/accounts" emptyName="Nova conta" hasType />
        <ConfigManager title="Formas de pagamento" endpoint="/payment-methods" emptyName="Nova forma" />
      </div>
    </div>
  );
}

function AppearanceManager() {
  useEffect(() => applySavedTheme(), []);
  async function upload(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const text = await file.text();
    const colors = [...new Set(text.match(/#[0-9a-fA-F]{6}/g) || [])];
    if (!colors.length) return alert('Arquivo sem cores hex.');
    const theme = {
      '--accent': colors[0],
      '--accent-strong': colors[1] || colors[0],
      '--accent-text': colors[2] || colors[1] || colors[0],
      '--bg-glow-primary': hexToRgba(colors[0], 0.14)
    };
    localStorage.setItem('finans_theme', JSON.stringify(theme));
    applyTheme(theme);
  }
  function reset() {
    localStorage.removeItem('finans_theme');
    location.reload();
  }
  return (
    <section className="glass mb-2">
      <div className="settings-head"><div><h3>Visual</h3><div className="subtitle">Importar padrão por .md</div></div><div className="row-flex"><label className="btn"><input type="file" accept=".md,text/markdown" hidden onChange={upload} />Trocar visual</label><button className="btn danger" onClick={reset}>Reset</button></div></div>
    </section>
  );
}

function ConfigManager({ title, endpoint, emptyName, hasType = false }) {
  const [rows, setRows] = useState([]);
  const [form, setForm] = useState({ name: '', color: COLORS[0], type: 'account' });
  const [editing, setEditing] = useState(null);
  async function load() { setRows(await api.get(endpoint)); }
  useEffect(() => { load(); }, []);
  async function save(e) {
    e.preventDefault();
    const payload = { ...form, name: form.name || emptyName };
    if (editing) await api.put(endpoint + '/' + editing.id, payload);
    else await api.post(endpoint, payload);
    setEditing(null);
    setForm({ name: '', color: COLORS[0], type: 'account' });
    load();
  }
  function edit(row) {
    setEditing(row);
    setForm({ name: row.name, color: row.color || COLORS[0], type: row.type || 'account' });
  }
  async function del(row) {
    if (!confirm('Excluir item?')) return;
    await api.del(endpoint + '/' + row.id);
    load();
  }
  return (
    <section className="glass config-card">
      <h3>{title}</h3>
      <form className="config-form" onSubmit={save}>
        <input className="input" placeholder="Nome" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
        {hasType && <select className="select" value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}><option value="account">Conta</option><option value="card">Cartão</option></select>}
        <input className="color-input" type="color" value={form.color} onChange={e => setForm({ ...form, color: e.target.value })} />
        <button className="btn accent">{editing ? 'Atualizar' : 'Adicionar'}</button>
        {editing && <button type="button" className="btn ghost" onClick={() => { setEditing(null); setForm({ name: '', color: COLORS[0], type: 'account' }); }}>Cancelar</button>}
      </form>
      <div className="category-list">
        {rows.map(row => (
          <div className="config-row" key={row.id}>
            <span className="badge"><span className="dot" style={{ background: row.color }} />{row.name}{row.type ? ` · ${row.type === 'card' ? 'Cartão' : 'Conta'}` : ''}</span>
            <div className="table-actions"><button className="btn sm" onClick={() => edit(row)}>Editar</button><button className="btn sm danger" onClick={() => del(row)}>×</button></div>
          </div>
        ))}
      </div>
    </section>
  );
}

function Users() {
  const [rows, setRows] = useState([]);
  const [form, setForm] = useState({ name: '', code: '', password: '', role: 'user' });
  const [editing, setEditing] = useState(null);
  async function load() { setRows(await api.get('/users')); }
  useEffect(() => { load(); }, []);
  async function save(e) {
    e.preventDefault();
    if (editing) await api.put('/users/' + editing.id, form);
    else await api.post('/users', form);
    setEditing(null);
    setForm({ name: '', code: '', password: '', role: 'user' });
    load();
  }
  function editUser(u) {
    setEditing(u);
    setForm({ name: u.name, code: u.code, password: '', role: u.role });
  }
  async function deleteUser(u) {
    if (!confirm('Excluir usuário?')) return;
    await api.del('/users/' + u.id);
    load();
  }
  return (
    <div>
      <div className="page-header"><div><h1>Usuários</h1><div className="subtitle">Acesso administrativo e visualização de dashboard</div></div></div>
      <form className="glass form-grid" onSubmit={save}>
        <input className="input" placeholder="Nome" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
        <input className="input" placeholder="Código" value={form.code} onChange={e => setForm({ ...form, code: e.target.value })} />
        <input className="input" placeholder="Senha" value={form.password} onChange={e => setForm({ ...form, password: e.target.value })} />
        <select className="select" value={form.role} onChange={e => setForm({ ...form, role: e.target.value })}><option value="user">Somente dashboard</option><option value="admin">Admin</option></select>
        <button className="btn accent">{editing ? 'Atualizar' : 'Criar usuário'}</button>
        {editing && <button type="button" className="btn ghost" onClick={() => { setEditing(null); setForm({ name: '', code: '', password: '', role: 'user' }); }}>Cancelar</button>}
      </form>
      <section className="glass table-panel">
        <table><thead><tr><th>Nome</th><th>Código</th><th>Perfil</th><th></th></tr></thead><tbody>{rows.map(u => <tr key={u.id}><td>{u.name}</td><td>{u.code}</td><td>{u.role === 'user' ? 'Somente dashboard' : u.role}</td><td><div className="table-actions"><button className="btn sm" onClick={() => editUser(u)}>Editar</button><button className="btn sm danger" onClick={() => deleteUser(u)}>×</button></div></td></tr>)}</tbody></table>
      </section>
    </div>
  );
}

function uniqueOptions(rows, key) {
  return [...new Set(rows.map(r => r[key]).filter(Boolean))];
}

function mergeNamed(rows, names) {
  const map = new Map();
  rows.forEach(r => map.set(r.name, r));
  names.forEach(name => { if (!map.has(name)) map.set(name, { name }); });
  return [...map.values()];
}

function filterRows(rows, filters) {
  return rows.filter(r => {
    if (filters.month && monthKey(r.date) !== filters.month) return false;
    if (filters.week && weekOfMonth(r.date) !== Number(filters.week)) return false;
    if (filters.from && r.date < filters.from) return false;
    if (filters.to && r.date > filters.to) return false;
    if (filters.type && r.type !== filters.type) return false;
    if (filters.category && r.category !== filters.category) return false;
    if (filters.group && r.group !== filters.group && !paymentSlice(r, filters).matched) return false;
    if (filters.account && r.account !== filters.account && !paymentSlice(r, filters).matched) return false;
    return true;
  }).map(r => {
    const slice = paymentSlice(r, filters);
    return slice.matched ? { ...r, amount: slice.amount, display_amount: slice.amount, original_amount: r.amount } : r;
  });
}

function weekOfMonth(date) {
  const day = Number(String(date || '').slice(8, 10));
  if (!day) return 0;
  return Math.min(5, Math.ceil(day / 7));
}

function paymentSlice(row, filters) {
  const payments = row.payments || [];
  if (!payments.length || (!filters.group && !filters.account)) return { matched: false, amount: 0 };
  const selected = payments.filter(p => (!filters.group || p.group === filters.group) && (!filters.account || p.account === filters.account));
  const amount = selected.reduce((sum, p) => sum + Number(p.amount || 0), 0);
  return { matched: selected.length > 0, amount };
}

function applyTheme(theme) {
  Object.entries(theme).forEach(([key, value]) => document.documentElement.style.setProperty(key, value));
}

function applySavedTheme() {
  const raw = localStorage.getItem('finans_theme');
  if (raw) applyTheme(JSON.parse(raw));
}

function hexToRgba(hex, alpha) {
  const n = parseInt(hex.slice(1), 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
}

function FilterDrawer({ filters, setFilters, categories, groups, accounts, onClose }) {
  const set = (k, v) => setFilters({ ...filters, [k]: v });
  const clear = () => setFilters({ month: '', week: '', from: '', to: '', type: '', category: '', group: '', account: '' });
  const weekOptions = filters.month ? monthWeeks(filters.month) : [1, 2, 3, 4, 5];
  return (
    <div className="drawer-backdrop" onClick={onClose}>
      <aside className="filter-drawer" onClick={e => e.stopPropagation()}>
        <div className="drawer-head"><h2>Filtros</h2><button className="modal-close" onClick={onClose}>×</button></div>
        <div className="drawer-section">
          <div className="grid-2">
            <div className="field"><label className="label">Mês</label><input className="input" type="month" value={filters.month || ''} onChange={e => setFilters({ ...filters, month: e.target.value, week: '' })} /></div>
            <div className="field"><label className="label">Semana do mês</label><select className="select" value={filters.week || ''} onChange={e => set('week', e.target.value)}><option value="">Todas</option>{weekOptions.map(w => <option key={w} value={w}>Semana {w}</option>)}</select></div>
          </div>
          <div className="grid-2">
            <div className="field"><label className="label">De</label><input className="input" type="date" value={filters.from} onChange={e => set('from', e.target.value)} /></div>
            <div className="field"><label className="label">Até</label><input className="input" type="date" value={filters.to} onChange={e => set('to', e.target.value)} /></div>
          </div>
          <div className="field"><label className="label">Tipo</label><select className="select" value={filters.type} onChange={e => set('type', e.target.value)}><option value="">Todos</option><option value="expense">Saídas</option><option value="income">Entradas</option></select></div>
          <div className="field"><label className="label">Categoria</label><select className="select" value={filters.category} onChange={e => set('category', e.target.value)}><option value="">Todas</option>{categories.map(c => <option key={c.name} value={c.name}>{c.name}</option>)}</select></div>
          <div className="field"><label className="label">Grupo</label><select className="select" value={filters.group} onChange={e => set('group', e.target.value)}><option value="">Todos</option>{groups.map(g => <option key={g.id} value={g.name}>{g.name}</option>)}</select></div>
          <div className="field"><label className="label">Cartão/Conta</label><select className="select" value={filters.account} onChange={e => set('account', e.target.value)}><option value="">Todos</option>{accounts.map(a => <option key={a} value={a}>{a}</option>)}</select></div>
        </div>
        <div className="modal-actions"><button className="btn ghost" onClick={clear}>Limpar</button><button className="btn accent" onClick={onClose}>Aplicar</button></div>
      </aside>
    </div>
  );
}

function monthWeeks(month) {
  const [year, m] = String(month).split('-').map(Number);
  const days = new Date(year, m, 0).getDate();
  return Array.from({ length: Math.ceil(days / 7) }, (_, i) => i + 1);
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
