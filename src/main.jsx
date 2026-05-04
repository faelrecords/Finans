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
            ['categorias', 'Categorias'],
            ['usuarios', 'Usuários']
          ]).map(([key, label]) => <button key={key} className={`nav-link ${tab === key ? 'active' : ''}`} onClick={() => setTab(key)}>{label}</button>)}
        </div>
        <div className="nav-user"><span>{user.name}</span><button onClick={logout}>Sair</button></div>
      </nav>
      <main className="container">
        {tab === 'dashboard' && <Dashboard readOnly={readOnly} />}
        {!readOnly && tab === 'transacoes' && <Transactions />}
        {!readOnly && tab === 'categorias' && <Categories />}
        {!readOnly && tab === 'usuarios' && <Users />}
      </main>
    </>
  );
}

function Dashboard({ readOnly = false }) {
  const [rows, setRows] = useState([]);
  const [categories, setCategories] = useState([]);
  const [groups, setGroups] = useState([]);
  const [dashboards, setDashboards] = useState([]);
  const [active, setActive] = useState(null);
  const [widgets, setWidgets] = useState([]);
  const [editing, setEditing] = useState(null);
  const [showWidget, setShowWidget] = useState(false);
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [filters, setFilters] = useState({ from: '', to: '', type: '', category: '', group: '', account: '' });
  const month = today().slice(0, 7);
  const filteredRows = useMemo(() => filterRows(rows, filters), [rows, filters]);
  const monthRows = filteredRows.filter(r => monthKey(r.date) === month);
  const accounts = useMemo(() => uniqueOptions(rows, 'account'), [rows]);

  async function loadDashboards() {
    const list = await api.get('/dashboards');
    setDashboards(list);
    setActive(current => current && list.find(d => d.id === current.id) ? current : list[0] || null);
    return list;
  }
  async function load() {
    const list = dashboards.length ? dashboards : await loadDashboards();
    const dash = active || list[0];
    const [tx, ws, cats, gs] = await Promise.all([
      api.get('/transactions'),
      dash ? api.get(`/widgets?dashboard_id=${dash.id}`) : [],
      api.get('/categories'),
      api.get('/groups')
    ]);
    setRows(tx);
    setWidgets(ws);
    setCategories(cats);
    setGroups(gs);
  }
  useEffect(() => { loadDashboards(); }, []);
  useEffect(() => { if (active) load(); }, [active?.id]);

  const income = aggregate(monthRows, 'income');
  const expense = aggregate(monthRows, 'expense');
  const balance = aggregate(filteredRows, 'balance');

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
      <section className="quick-grid">
        <div className="glass-sm"><label>Saldo</label><strong>{money(balance)}</strong></div>
        <div className="glass-sm"><label>Entradas mês</label><strong>{money(income)}</strong></div>
        <div className="glass-sm"><label>Saídas mês</label><strong>{money(expense)}</strong></div>
        <div className="glass-sm"><label>Resultado mês</label><strong>{money(income - expense)}</strong></div>
      </section>
      {widgets.length === 0 ? (
        <div className="glass empty-state"><h3>Nenhum widget ainda</h3>{!readOnly && <button className="btn accent mt-2" onClick={() => setShowWidget(true)}>+ Criar widget</button>}</div>
      ) : (
        <div className="widgets-grid">
          {widgets.map(w => <WidgetCard key={w.id} widget={w} rows={filteredRows} categories={categories} onEdit={readOnly ? null : () => setEditing(w)} onDelete={readOnly ? null : () => deleteWidget(w.id)} />)}
        </div>
      )}
      {filtersOpen && <FilterDrawer filters={filters} setFilters={setFilters} categories={categories} groups={groups} accounts={accounts} onClose={() => setFiltersOpen(false)} />}
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
  const [groupFilter, setGroupFilter] = useState('');
  const [groupForm, setGroupForm] = useState({ name: '', color: COLORS[0] });
  const [form, setForm] = useState({ date: today(), type: 'expense', description: '', category: '', group: '', account: '', amount: '' });
  const [editing, setEditing] = useState(null);
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [filters, setFilters] = useState({ from: '', to: '', type: '', category: '', group: '', account: '' });
  async function load() {
    const [tx, gs] = await Promise.all([api.get('/transactions'), api.get('/groups')]);
    setRows(tx);
    setGroups(gs);
  }
  useEffect(() => { load(); }, []);
  const categories = useMemo(() => uniqueOptions(rows, 'category').map(name => ({ name })), [rows]);
  const accounts = useMemo(() => uniqueOptions(rows, 'account'), [rows]);
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
  async function addGroup(e) {
    e.preventDefault();
    await api.post('/groups', groupForm);
    setGroupForm({ name: '', color: COLORS[0] });
    load();
  }
  async function updateGroup(g, color) {
    await api.put('/groups/' + g.id, { ...g, color });
    load();
  }
  async function deleteGroup(g) {
    if (!confirm('Excluir grupo?')) return;
    await api.del('/groups/' + g.id);
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
        <div className="label mb-2">Grupos</div>
        <div className="row-flex mb-2">
          <button className={`range-pill ${!groupFilter ? 'active' : ''}`} onClick={() => setGroupFilter('')}>Todos</button>
          {groups.map(g => <button key={g.id} className={`range-pill ${groupFilter === g.name ? 'active' : ''}`} onClick={() => setGroupFilter(g.name)}><span className="dot" style={{ background: g.color }} />{g.name}</button>)}
        </div>
        <form className="row-flex" onSubmit={addGroup}>
          <input className="input compact" placeholder="Novo grupo" value={groupForm.name} onChange={e => setGroupForm({ ...groupForm, name: e.target.value })} />
          <input className="color-input" type="color" value={groupForm.color} onChange={e => setGroupForm({ ...groupForm, color: e.target.value })} />
          <button className="btn accent">+ Grupo</button>
        </form>
        <div className="group-list">
          {groups.map(g => <div className="group-row" key={g.id}><span className="badge"><span className="dot" style={{ background: g.color }} />{g.name}</span><input className="color-input" type="color" value={g.color} onChange={e => updateGroup(g, e.target.value)} /><button className="btn sm danger" onClick={() => deleteGroup(g)}>×</button></div>)}
        </div>
      </section>
      <form className="glass form-grid" onSubmit={save}>
        <input className="input" type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
        <select className="select" value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}><option value="expense">Saída</option><option value="income">Entrada</option></select>
        <input className="input" placeholder="Descrição" value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} />
        <input className="input" placeholder="Categoria" value={form.category} onChange={e => setForm({ ...form, category: e.target.value })} />
        <select className="select" value={form.group} onChange={e => setForm({ ...form, group: e.target.value })}><option value="">Grupo</option>{groups.map(g => <option key={g.id} value={g.name}>{g.name}</option>)}</select>
        <input className="input" placeholder="Cartão/Conta" value={form.account} onChange={e => setForm({ ...form, account: e.target.value })} />
        <input className="input" placeholder="Valor" type="number" step="0.01" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} />
        <button className="btn accent">{editing ? 'Atualizar' : 'Salvar'}</button>
        {editing && <button type="button" className="btn ghost" onClick={() => { setEditing(null); setForm({ date: today(), type: 'expense', description: '', category: '', group: '', account: '', amount: '' }); }}>Cancelar</button>}
      </form>
      <section className="glass table-panel">
        <table><thead><tr><th>Data</th><th>Tipo</th><th>Descrição</th><th>Categoria</th><th>Grupo</th><th>Cartão/Conta</th><th>Valor</th><th></th></tr></thead>
        <tbody>{visibleRows.map(r => <tr key={r.id}><td>{brDate(r.date)}</td><td>{r.type === 'income' ? 'Entrada' : 'Saída'}</td><td>{r.description}</td><td>{r.category}</td><td>{r.group}</td><td>{r.account}</td><td>{money(r.amount)}</td><td><div className="table-actions"><button className="btn sm" onClick={() => editRow(r)}>Editar</button><button className="btn sm danger" onClick={() => delRow(r)}>×</button></div></td></tr>)}</tbody></table>
      </section>
      {filtersOpen && <FilterDrawer filters={filters} setFilters={setFilters} categories={categories} groups={groups} accounts={accounts} onClose={() => setFiltersOpen(false)} />}
    </div>
  );
}

function Categories() {
  const [rows, setRows] = useState([]);
  const [form, setForm] = useState({ name: '', color: COLORS[0] });
  async function load() { setRows(await api.get('/categories')); }
  useEffect(() => { load(); }, []);
  async function add(e) {
    e.preventDefault();
    await api.post('/categories', form);
    setForm({ name: '', color: COLORS[0] });
    load();
  }
  async function setColor(c, color) {
    await api.put('/categories/' + c.id, { ...c, color });
    load();
  }
  async function del(c) {
    if (!confirm('Excluir categoria?')) return;
    await api.del('/categories/' + c.id);
    load();
  }
  return (
    <div><div className="page-header"><div><h1>Categorias</h1><div className="subtitle">Classificação de entradas e saídas</div></div></div>
      <section className="glass">
        <form className="row-flex mb-2" onSubmit={add}>
          <input className="input" placeholder="Categoria" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })} />
          <div className="color-picker">{COLORS.map(c => <button type="button" key={c} className={`color-dot ${form.color === c ? 'selected' : ''}`} style={{ background: c }} onClick={() => setForm({ ...form, color: c })} />)}</div>
          <button className="btn accent">Adicionar</button>
        </form>
        <div className="category-list">
          {rows.map(c => (
            <div className="category-row" key={c.id}>
              <span className="badge"><span className="dot" style={{ background: c.color }} />{c.name}</span>
              <div className="color-picker small">{COLORS.map(color => <button type="button" key={color} className={`color-dot ${c.color === color ? 'selected' : ''}`} style={{ background: color }} onClick={() => setColor(c, color)} />)}</div>
              <button className="btn sm danger" onClick={() => del(c)}>×</button>
            </div>
          ))}
        </div>
      </section>
    </div>
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

function filterRows(rows, filters) {
  return rows.filter(r => {
    if (filters.from && r.date < filters.from) return false;
    if (filters.to && r.date > filters.to) return false;
    if (filters.type && r.type !== filters.type) return false;
    if (filters.category && r.category !== filters.category) return false;
    if (filters.group && r.group !== filters.group) return false;
    if (filters.account && r.account !== filters.account) return false;
    return true;
  });
}

function FilterDrawer({ filters, setFilters, categories, groups, accounts, onClose }) {
  const set = (k, v) => setFilters({ ...filters, [k]: v });
  const clear = () => setFilters({ from: '', to: '', type: '', category: '', group: '', account: '' });
  return (
    <div className="drawer-backdrop" onClick={onClose}>
      <aside className="filter-drawer" onClick={e => e.stopPropagation()}>
        <div className="drawer-head"><h2>Filtros</h2><button className="modal-close" onClick={onClose}>×</button></div>
        <div className="drawer-section">
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

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
