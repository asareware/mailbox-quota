export async function fetchMailFolders(accessToken) {
  const res = await fetch('https://fa-mailbox-quota-gzecfehbb7emggf0.centralus-01.azurewebsites.net/api/MailFoldersFunction', {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Backend error ${res.status}: ${txt}`);
  }
  return res.json();
}
