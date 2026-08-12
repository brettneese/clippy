import assert from "node:assert/strict";
import type { AddressInfo } from "node:net";
import { after, before, describe, it } from "node:test";
import {
  createClippyServer,
  ReplyGenerationError,
  type ReplyGenerator,
} from "./server.js";

const quietLogger = {
  log(): void {},
  error(): void {},
};

function testServer(generateReply: ReplyGenerator) {
  const server = createClippyServer({ generateReply, logger: quietLogger });
  let baseUrl = "";

  before(async () => {
    await new Promise<void>((resolve, reject) => {
      server.once("error", reject);
      server.listen(0, "127.0.0.1", resolve);
    });
    const address = server.address() as AddressInfo;
    baseUrl = `http://127.0.0.1:${address.port}`;
  });

  after(async () => {
    await new Promise<void>((resolve, reject) => {
      server.close((error) => error === undefined ? resolve() : reject(error));
    });
  });

  return {
    url(path = "/message"): string {
      return `${baseUrl}${path}`;
    },
  };
}

async function withTestServer(
  generateReply: ReplyGenerator,
  run: (url: string) => Promise<void>,
): Promise<void> {
  const server = createClippyServer({ generateReply, logger: quietLogger });
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address() as AddressInfo;
  try {
    await run(`http://127.0.0.1:${address.port}/message`);
  } finally {
    await new Promise<void>((resolve, reject) => {
      server.close((error) => error === undefined ? resolve() : reject(error));
    });
  }
}

describe("Clippy host server routing and validation", () => {
  let calls = 0;
  const host = testServer(async () => {
    calls += 1;
    return "unused";
  });

  it("rejects unknown routes and methods", async () => {
    const wrongRoute = await fetch(host.url("/other"), { method: "POST" });
    assert.equal(wrongRoute.status, 404);
    assert.deepEqual(await wrongRoute.json(), { error: "Not found" });

    const wrongMethod = await fetch(host.url(), { method: "GET" });
    assert.equal(wrongMethod.status, 404);
    assert.deepEqual(await wrongMethod.json(), { error: "Not found" });
    assert.equal(calls, 0);
  });

  it("rejects malformed JSON", async () => {
    const response = await fetch(host.url(), {
      method: "POST",
      body: "{not json",
    });
    assert.equal(response.status, 400);
    assert.match((await response.json() as { error: string }).error, /JSON/);
    assert.equal(calls, 0);
  });

  it("rejects empty text", async () => {
    const response = await fetch(host.url(), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ text: "   " }),
    });
    assert.equal(response.status, 400);
    assert.deepEqual(await response.json(), {
      error: "Expected a non-empty string property named text",
    });
    assert.equal(calls, 0);
  });
});

describe("Clippy host server replies", () => {
  const input = "Snowman ☃ says \"hello\"\non the next line";
  const output = "📎 Unicode, quotes, and newlines made it through.";
  let received = "";
  const host = testServer(async (text) => {
    received = text;
    return output;
  });

  it("returns the generated reply with UTF-8 and JSON escaping intact", async () => {
    const response = await fetch(host.url(), {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ text: input }),
    });
    assert.equal(response.status, 200);
    assert.match(response.headers.get("content-type") ?? "", /^application\/json/);
    assert.deepEqual(await response.json(), { text: output });
    assert.equal(received, input);
  });
});

describe("Clippy host server upstream failures", () => {
  it("maps an empty generated reply to HTTP 502", async () => {
    await withTestServer(async () => "   ", async (url) => {
      const response = await fetch(url, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ text: "Hello" }),
      });
      assert.equal(response.status, 502);
      assert.deepEqual(await response.json(), { error: "OpenAI request failed" });
    });
  });

  it("maps a model timeout to HTTP 504", async () => {
    await withTestServer(async () => {
      throw new ReplyGenerationError("timeout", "simulated timeout");
    }, async (url) => {
      const response = await fetch(url, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ text: "Hello" }),
      });
      assert.equal(response.status, 504);
      assert.deepEqual(await response.json(), { error: "OpenAI request timed out" });
    });
  });

  it("maps other model failures to HTTP 502", async () => {
    await withTestServer(async () => {
      throw new Error("simulated upstream failure");
    }, async (url) => {
      const response = await fetch(url, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ text: "Hello" }),
      });
      assert.equal(response.status, 502);
      assert.deepEqual(await response.json(), { error: "OpenAI request failed" });
    });
  });
});
