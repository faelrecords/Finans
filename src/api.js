const API_URL = import.meta.env.VITE_API_URL || 'https://wblsxsidbjnrvgywjvim.supabase.co/functions/v1/api';

export const api = {
  token: localStorage.getItem('finans_token'),
  user: JSON.parse(localStorage.getItem('finans_user') || 'null'),
  setAuth(token, user) {
    this.token = token;
    this.user = user;
    if (token) localStorage.setItem('finans_token', token);
    else localStorage.removeItem('finans_token');
    if (user) localStorage.setItem('finans_user', JSON.stringify(user));
    else localStorage.removeItem('finans_user');
  },
  async req(path, opts = {}) {
    const res = await fetch(API_URL + path, {
      ...opts,
      headers: {
        'Content-Type': 'application/json',
        ...(this.token ? { Authorization: `Bearer ${this.token}` } : {}),
        ...(opts.headers || {})
      }
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Erro');
    return data;
  },
  get(path) { return this.req(path); },
  post(path, body) { return this.req(path, { method: 'POST', body: JSON.stringify(body) }); },
  put(path, body) { return this.req(path, { method: 'PUT', body: JSON.stringify(body) }); },
  del(path) { return this.req(path, { method: 'DELETE' }); }
};


