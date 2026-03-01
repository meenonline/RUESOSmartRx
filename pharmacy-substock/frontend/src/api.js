import axios from 'axios';

const BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001';

const api = axios.create({ baseURL: BASE_URL });

export async function apiGet(path) {
  const res = await api.get(path);
  return res.data;
}

export async function apiPost(path, body) {
  const res = await api.post(path, body);
  return res.data;
}

export async function apiPatch(path, body) {
  const res = await api.patch(path, body);
  return res.data;
}

export async function apiDelete(path, body) {
  const res = await api.delete(path, { data: body });
  return res.data;
}
