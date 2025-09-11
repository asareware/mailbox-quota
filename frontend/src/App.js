import React, { useState } from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "./authConfig";
import { fetchMailFolders } from "./apiClient";
import { bytesToGB } from "./utils";
import { PieChart, Pie, Tooltip, ResponsiveContainer, Cell } from "recharts";

function App() {
  const { instance } = useMsal();
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);

  async function signInAndFetch() {
    setLoading(true);
    try {
      const loginResp = await instance.loginPopup(loginRequest);
      const account = loginResp.account;
      const tokenResp = await instance.acquireTokenSilent({ scopes: loginRequest.scopes, account })
        .catch(() => instance.acquireTokenPopup({ scopes: loginRequest.scopes }));
      const accessToken = tokenResp.accessToken;
      const payload = await fetchMailFolders(accessToken);
      setData(payload);
    } catch (err) {
      console.error(err);
      alert('Error: ' + (err.message || err));
    } finally {
      setLoading(false);
    }
  }

  if (!data) {
    return (
      <div style={{ padding: 20 }}>
        <h1>Mailbox Space Dashboard</h1>
        <p>Sign in to see your mailbox folder sizes (visible folders only).</p>
        <button onClick={signInAndFetch} disabled={loading}>
          {loading ? "Loading..." : "Sign in & Load folders"}
        </button>
      </div>
    );
  }

  const top = data.folders.slice(0, 8).map(f => ({ name: f.folderPath, value: f.cumulativeBytes }));
  const colors = ['#3366CC','#DC3912','#FF9900','#109618','#990099','#3B3EAC','#0099C6','#DD4477'];

  return (
    <div style={{ padding: 20 }}>
      <h1>Mailbox Space Dashboard</h1>
      <h3>Total: {bytesToGB(data.totalMailboxBytes, 2)}</h3>
      <div style={{ display: 'flex', gap: 20 }}>
        <div style={{ width: 350, height: 300 }}>
          <ResponsiveContainer>
            <PieChart>
              <Pie data={top} dataKey="value" nameKey="name" outerRadius={100} label>
                {top.map((_, i) => <Cell key={i} fill={colors[i % colors.length]} />)}
              </Pie>
              <Tooltip formatter={(val) => bytesToGB(val)} />
            </PieChart>
          </ResponsiveContainer>
        </div>
        <div style={{ flex: 1 }}>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr style={{ textAlign:'left' }}>
                <th>Folder</th><th style={{textAlign:'right'}}>Size (GB)</th><th style={{textAlign:'right'}}>Items</th>
              </tr>
            </thead>
            <tbody>
              {data.folders.sort((a,b) => b.cumulativeBytes - a.cumulativeBytes).map(f =>
                <tr key={f.id} style={{ borderBottom:'1px solid #eee' }}>
                  <td>{f.folderPath}</td>
                  <td style={{ textAlign:'right' }}>{bytesToGB(f.cumulativeBytes)}</td>
                  <td style={{ textAlign:'right' }}>{f.totalItemCount}</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

export default App;