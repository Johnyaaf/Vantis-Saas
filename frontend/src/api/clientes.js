import { apiRequest } from "./client";

export function listClientes(token, buscar = "") {
  const query = buscar.trim() ? `?buscar=${encodeURIComponent(buscar.trim())}` : "";
  return apiRequest(`/api/clientes${query}`, { token });
}

export function createCliente(token, payload) {
  return apiRequest("/api/clientes", {
    method: "POST",
    token,
    body: JSON.stringify(payload),
  });
}

export function updateCliente(token, id, payload) {
  return apiRequest(`/api/clientes/${id}`, {
    method: "PUT",
    token,
    body: JSON.stringify(payload),
  });
}

export function deleteCliente(token, id) {
  return apiRequest(`/api/clientes/${id}`, {
    method: "DELETE",
    token,
  });
}
