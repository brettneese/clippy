import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { fileURLToPath } from "node:url";
import OpenAI from "openai";

const DEFAULT_HOST = "0.0.0.0";
const DEFAULT_PORT = 3210;
const DEFAULT_MODEL = "gpt-5.6-luna";
const MAX_BODY_BYTES = 64 * 1024;
const MAX_OUTPUT_TOKENS = 512;
const OPENAI_TIMEOUT_MS = 45_000;
const CLIPPY_INSTRUCTIONS = [
  "You are Clippy, the helpful Microsoft Office Assistant.",
  "Answer the user's question directly in lightweight Markdown suitable for a small Office Assistant balloon.",
  "Use no more than 150 words and at most three short paragraphs.",
  "Be accurate, concise, and lightly playful, but never let the character voice get in the way of the answer.",
  "When it improves readability, use **strong emphasis**, *emphasis*, `inline code`, links, a level 1-3 heading, or one final bullet or numbered list with at most five items.",
  "Do not use Markdown tables, HTML, images, nested lists, generic introductions, or unnecessary sign-offs.",
  "Never emit Office Assistant brace directives such as {ul}, {cf}, {bmp}, or {wmf}.",
  "If you are uncertain, say so plainly.",
].join(" ");

type ClippyRequest = {
  text: string;
};

type Logger = Pick<Console, "log" | "error">;

export type ReplyGenerator = (text: string) => Promise<string>;

export type ClippyServerOptions = {
  generateReply: ReplyGenerator;
  logger?: Logger;
};

type ReplyErrorKind = "timeout" | "upstream";

export class ReplyGenerationError extends Error {
  constructor(
    readonly kind: ReplyErrorKind,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "ReplyGenerationError";
  }
}

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

function describeOpenAIError(error: unknown): string {
  if (error instanceof OpenAI.APIError) {
    const requestId = error.requestID === undefined || error.requestID === null
      ? "unavailable"
      : error.requestID;
    return `${error.name}: ${error.message} (request id: ${requestId})`;
  }
  return error instanceof Error ? `${error.name}: ${error.message}` : String(error);
}

export function createOpenAIReplyGenerator(
  apiKey: string,
  model = DEFAULT_MODEL,
): ReplyGenerator {
  const client = new OpenAI({
    apiKey,
    timeout: OPENAI_TIMEOUT_MS,
    maxRetries: 0,
  });

  return async (text: string): Promise<string> => {
    try {
      const response = await client.responses.create({
        model,
        instructions: CLIPPY_INSTRUCTIONS,
        input: text,
        reasoning: { effort: "low" },
        text: { verbosity: "low" },
        max_output_tokens: MAX_OUTPUT_TOKENS,
        store: false,
      });
      const reply = response.output_text.trim();
      if (reply.length === 0) {
        throw new ReplyGenerationError(
          "upstream",
          `OpenAI response ${response.id} did not contain output text`,
        );
      }
      return reply;
    } catch (error) {
      if (error instanceof ReplyGenerationError) {
        throw error;
      }
      const kind: ReplyErrorKind = error instanceof OpenAI.APIConnectionTimeoutError
        ? "timeout"
        : "upstream";
      throw new ReplyGenerationError(kind, describeOpenAIError(error), {
        cause: error,
      });
    }
  };
}

export function createClippyServer(options: ClippyServerOptions) {
  const logger = options.logger ?? console;

  return createServer(async (request, response) => {
    if (request.method !== "POST" || request.url !== "/message") {
      sendJson(response, 404, { error: "Not found" });
      return;
    }

    let message: ClippyRequest;
    try {
      message = parseClippyRequest(await readBody(request));
    } catch (error) {
      const detail = error instanceof Error ? error.message : "Invalid request";
      sendJson(response, 400, { error: detail });
      return;
    }

    logger.log(`Clippy: ${message.text}`);
    try {
      const reply = (await options.generateReply(message.text)).trim();
      if (reply.length === 0) {
        throw new ReplyGenerationError(
          "upstream",
          "Reply generator did not return any text",
        );
      }
      sendJson(response, 200, { text: reply });
    } catch (error) {
      const replyError = error instanceof ReplyGenerationError
        ? error
        : new ReplyGenerationError("upstream", describeOpenAIError(error), {
            cause: error,
          });
      logger.error(`OpenAI request failed: ${replyError.message}`);
      if (replyError.kind === "timeout") {
        sendJson(response, 504, { error: "OpenAI request timed out" });
      } else {
        sendJson(response, 502, { error: "OpenAI request failed" });
      }
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

function requireApiKey(value: string | undefined): string {
  if (value === undefined || value.trim().length === 0) {
    throw new Error("OPENAI_API_KEY must be set before starting the Clippy host server");
  }
  return value;
}

function parseModel(value: string | undefined): string {
  if (value === undefined) {
    return DEFAULT_MODEL;
  }
  const model = value.trim();
  if (model.length === 0) {
    throw new Error("CLIPPY_OPENAI_MODEL cannot be empty");
  }
  return model;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  try {
    const host = process.env.CLIPPY_SERVER_BIND ?? DEFAULT_HOST;
    const port = parsePort(process.env.CLIPPY_SERVER_PORT);
    const model = parseModel(process.env.CLIPPY_OPENAI_MODEL);
    const apiKey = requireApiKey(process.env.OPENAI_API_KEY);
    const server = createClippyServer({
      generateReply: createOpenAIReplyGenerator(apiKey, model),
    });
    server.listen(port, host, () => {
      console.log(
        `Clippy host server listening on http://${host}:${port}/message using ${model}`,
      );
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`Clippy host server failed to start: ${message}`);
    process.exitCode = 1;
  }
}
