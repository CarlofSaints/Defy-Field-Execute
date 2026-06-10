// Resolve the Vercel Blob read/write token.
//
// This project has been connected to several Blob stores over time, and each
// connection injected a differently-prefixed token (BLOB_READ_WRITE_TOKEN,
// NEW_READ_WRITE_TOKEN, BLOB_STORE_ID_DEFY_READ_WRITE_TOKEN, …). The @vercel/blob
// SDK and our data modules only look for BLOB_READ_WRITE_TOKEN, so if that exact
// name isn't set we alias the first token we can find into it at startup. Every
// module that touches Blob imports { blobEnabled } from here, which guarantees
// this aliasing runs before any put/list/del call.
if (!process.env.BLOB_READ_WRITE_TOKEN) {
  const fallback =
    process.env.NEW_READ_WRITE_TOKEN ||
    process.env.BLOB_STORE_ID_DEFY_READ_WRITE_TOKEN;
  if (fallback) process.env.BLOB_READ_WRITE_TOKEN = fallback;
}

export const blobEnabled = !!process.env.BLOB_READ_WRITE_TOKEN;
