/**
 * Small client-side API helper for the Vercel/Prisma backend.
 * The existing HRM UI is intentionally not switched over automatically;
 * use these helpers when migrating each module from browser storage to DB.
 */
async function request(path, options = {}) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || `API request failed (${res.status})`);
  return data;
}

export const hrmApi = {
  health: () => request("/api/health"),
  getState: () => request("/api/state"),
  saveState: (state) => request("/api/state", { method: "POST", body: JSON.stringify({ state }) }),
  listEmployees: () => request("/api/employees"),
  createEmployee: (employee) => request("/api/employees", { method: "POST", body: JSON.stringify(employee) }),
  updateEmployee: (employeeId, employee) => request(`/api/employees?employeeId=${encodeURIComponent(employeeId)}`, { method: "PUT", body: JSON.stringify(employee) }),
  deleteEmployee: (employeeId) => request(`/api/employees?employeeId=${encodeURIComponent(employeeId)}`, { method: "DELETE", body: JSON.stringify({ employeeId }) }),
  listAttendance: (month) => request(`/api/attendance${month ? `?month=${encodeURIComponent(month)}` : ""}`),
  saveAttendance: (record) => request("/api/attendance", { method: "POST", body: JSON.stringify(record) }),
  listLeave: (employeeId) => request(`/api/leave${employeeId ? `?employeeId=${encodeURIComponent(employeeId)}` : ""}`),
  createLeave: (leave) => request("/api/leave", { method: "POST", body: JSON.stringify(leave) }),
  listPayroll: (month) => request(`/api/payroll${month ? `?month=${encodeURIComponent(month)}` : ""}`),
  savePayroll: (payroll) => request("/api/payroll", { method: "POST", body: JSON.stringify(payroll) }),
};
