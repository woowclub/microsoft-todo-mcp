/** Bound decompressed bytes, including responses without a Content-Length. */
export async function readBoundedText(response: Response, maxBytes: number, invalid: Error): Promise<string> {
  const reader = response.body?.getReader();
  if (!reader) throw invalid;
  const chunks: Uint8Array[] = [];
  let bytes = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      bytes += value.byteLength;
      if (bytes > maxBytes) {
        await reader.cancel();
        throw invalid;
      }
      chunks.push(value);
    }
    return Buffer.concat(chunks).toString("utf8");
  } finally {
    reader.releaseLock();
  }
}
