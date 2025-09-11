export function bytesToGB(bytes, decimals = 2) {
  if (!bytes) return (0).toFixed(decimals) + ' GB';
  return (bytes / (1024 ** 3)).toFixed(decimals) + ' GB';
}