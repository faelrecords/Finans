import React, { useEffect, useMemo, useState } from 'react';
import ReactDOM from 'react-dom/client';
import { Bar, BarChart, CartesianGrid, ResponsiveContainer, Tooltip, XAxis, YAxis } from 'recharts';
import * as XLSX from 'xlsx';
import { api } from './api.js';
import './styles.css';

const today = () => new Date().toISOString().slice(0, 10);
const money = v => 'R$ ' + Number(v || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const monthKey = d => String(d || '').slice(0, 7);

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
    <main className="login">
      <form className="panel login-card" onSubmit={submit}>
        <div className="brand">Finans</div>
        <input placeholder="Código" value={form.code} onChange={e => setForm({ ...form, code: e.target.value })} />
        <input placeholder="Senha" type="password" value={form.password} onChange={e => setForm({ ...form, password: e.target.value })} />
        {error && <div className="error">{error}</div>}
        <button>Entrar</button>
      </form>
    </main>
  );
}

function App() {
  const [user, setUser] = useState(api.user);
  const [tab, setTab] = useState('dashboard');
  const [rows, setRows] = useState([]);
  const [cats, setCats] = useState([]);
  const [form, setForm] = useState({ date: today(), type: 'expense', description: '', category: '', account: '', amount: '' });

  async function load() {
    const [tx, c] = await Promise.all([api.get('/transactions'), api.get('/categories')]);
    setRows(tx);
    setCats(c);
  }
  useEffect(() => { if (user) load(); }, [user]);
  if (!user) return <Login onLogin={setUser} />;

  const month = today().slice(0, 7);
  const current = rows.filter(r => monthKey(r.date) === month);
  const income = current.filter(r => r.type === 'income').reduce((s, r) => s + Number(r.amount), 0);
  const expense = current.filter(r => r.type === 'expense').reduce((s, r) => s + Number(r.amount), 0);
  const balance = rows.reduce((s, r) => s + (r.type === 'income' ? Number(r.amount) : -Number(r.amount)), 0);
  const byCat = Object.values(current.filter(r => r.type === 'expense').reduce((acc, r) => {
    acc[r.category] ||= { name: r.category || 'Sem categoria', value: 0 };
    acc[r.category].value += Number(r.amount);
    return acc;
  }, {}));

  async function save(e) {
    e.preventDefault();
    await api.post('/transactions', { ...form, amount: Number(form.amount) || 0 });
    setForm({ date: today(), type: 'expense', description: '', category: '', account: '', amount: '' });
    load();
  }
  async function exportXLS() {
    const ws = XLSX.utils.json_to_sheet(rows.map(r => ({ Data: r.date, Tipo: r.type, Descrição: r.description, Categoria: r.category, Conta: r.account, Valor: r.amount })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Transações');
    XLSX.writeFile(wb, 'finans_transacoes.xlsx');
  }
  async function logout() {
    api.setAuth(null, null);
    setUser(null);
  }

  return (
    <>
      <nav>
        <div className="brand small">Finans</div>
        {['dashboard', 'transacoes', 'categorias'].map(t => <button className={tab === t ? 'active' : ''} onClick={() => setTab(t)} key={t}>{t}</button>)}
        <span className="spacer" />
        <span>{user.name}</span>
        <button onClick={logout}>Sair</button>
      </nav>
      <main className="wrap">
        {tab === 'dashboard' && (
          <>
            <section className="cards">
              <div className="panel"><label>Saldo</label><strong>{money(balance)}</strong></div>
              <div className="panel"><label>Receitas mês</label><strong>{money(income)}</strong></div>
              <div className="panel"><label>Saídas mês</label><strong>{money(expense)}</strong></div>
              <div className="panel"><label>Resultado mês</label><strong>{money(income - expense)}</strong></div>
            </section>
            <section className="panel chart">
              <h2>Gastos por categoria</h2>
              <ResponsiveContainer width="100%" height={280}>
                <BarChart data={byCat}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.08)" />
                  <XAxis dataKey="name" stroke="#aaa" />
                  <YAxis stroke="#aaa" />
                  <Tooltip />
                  <Bar dataKey="value" fill="#65d6ad" radius={[6, 6, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </section>
          </>
        )}
        {tab === 'transacoes' && (
          <>
            <form className="panel form-grid" onSubmit={save}>
              <input type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })} />
              <select value={form.type} onChange={e => setForm({ ...form, type: e.target.value })}><option value="expense">Saída</option><option value="income">Entrada</option></select>
              <input placeholder="Descrição" value={form.description} onChange={e => setForm({ ...form, description: e.target.value })} />
              <input placeholder="Categoria" value={form.category} onChange={e => setForm({ ...form, category: e.target.value })} />
              <input placeholder="Conta" value={form.account} onChange={e => setForm({ ...form, account: e.target.value })} />
              <input placeholder="Valor" type="number" step="0.01" value={form.amount} onChange={e => setForm({ ...form, amount: e.target.value })} />
              <button>Salvar</button>
              <button type="button" onClick={exportXLS}>Excel</button>
            </form>
            <section className="panel table-panel">
              <table>
                <thead><tr><th>Data</th><th>Tipo</th><th>Descrição</th><th>Categoria</th><th>Conta</th><th>Valor</th><th></th></tr></thead>
                <tbody>{rows.map(r => <tr key={r.id}><td>{r.date}</td><td>{r.type === 'income' ? 'Entrada' : 'Saída'}</td><td>{r.description}</td><td>{r.category}</td><td>{r.account}</td><td>{money(r.amount)}</td><td><button onClick={async () => { await api.del('/transactions/' + r.id); load(); }}>×</button></td></tr>)}</tbody>
              </table>
            </section>
          </>
        )}
        {tab === 'categorias' && <Categories rows={cats} reload={load} />}
      </main>
    </>
  );
}

function Categories({ rows, reload }) {
  const [name, setName] = useState('');
  async function add(e) {
    e.preventDefault();
    await api.post('/categories', { name });
    setName('');
    reload();
  }
  return <section className="panel"><form className="row" onSubmit={add}><input placeholder="Categoria" value={name} onChange={e => setName(e.target.value)} /><button>Adicionar</button></form>{rows.map(c => <span className="pill" key={c.id}>{c.name}</span>)}</section>;
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);

