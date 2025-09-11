export async function fetchMailFolders(accessToken) {
  const res = await fetch('https://<FUNCTION_APP_NAME>.azurewebsites.net/api/MailFoldersFunction', {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Backend error ${res.status}: ${txt}`);
  }
  return res.json();
}