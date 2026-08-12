import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { fileURLToPath } from "node:url";

const DEFAULT_HOST = "0.0.0.0";
const DEFAULT_PORT = 3210;
const MAX_BODY_BYTES = 64 * 1024;

type ClippyRequest = {
  text: string;
};

function sendJson(
  response: ServerResponse,
  statusCode: number,
  body: Record<string, unknown>,
): void {
  const encoded = JSON.stringify(body);
  response.writeHead(statusCode, {
    "content-type": "application/json; charset=utf-8",
    "content-length": Buffer.byteLength(encoded),
  });
  response.end(encoded);
}

async function readBody(request: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = [];
  let byteCount = 0;

  for await (const chunk of request) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    byteCount += buffer.length;
    if (byteCount > MAX_BODY_BYTES) {
      throw new Error("Request body is too large");
    }
    chunks.push(buffer);
  }

  return Buffer.concat(chunks).toString("utf8");
}

function parseClippyRequest(body: string): ClippyRequest {
  const value: unknown = JSON.parse(body);
  if (
    typeof value !== "object" ||
    value === null ||
    !("text" in value) ||
    typeof value.text !== "string" ||
    value.text.trim().length === 0
  ) {
    throw new Error("Expected a non-empty string property named text");
  }
  return { text: value.text };
}

export function createClippyServer() {
  return createServer(async (request, response) => {
    if (request.method !== "POST" || request.url !== "/message") {
      sendJson(response, 404, { error: "Not found" });
      return;
    }

    try {
      const message = parseClippyRequest(await readBody(request));
      console.log(`Clippy: ${message.text}`);
      sendJson(response, 200, { text: message.text });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Invalid request";
      sendJson(response, 400, { error: message });
    }
  });
}

function parsePort(value: string | undefined): number {
  if (value === undefined) {
    return DEFAULT_PORT;
  }
  const port = Number(value);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error(`Invalid CLIPPY_SERVER_PORT: ${value}`);
  }
  return port;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const host = process.env.CLIPPY_SERVER_BIND ?? DEFAULT_HOST;
  const port = parsePort(process.env.CLIPPY_SERVER_PORT);
  createClippyServer().listen(port, host, () => {
    console.log(`Clippy host server listening on http://${host}:${port}/message`);
  });
}
