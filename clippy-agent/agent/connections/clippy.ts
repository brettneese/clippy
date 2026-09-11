import { defineMcpClientConnection } from "eve/connections";

const url =
  process.env.CLIPPY_MCP_URL ??
  "https://pleasing-unicorn-legally.ngrok-free.app/mcp";
const apiKey = process.env.CLIPPY_MCP_API_KEY;

export default defineMcpClientConnection({
  url,
  description:
    "Clippy's primary embodiment: the authentic Microsoft Office character running on Windows XP. Use this connection by default for Clippy's visible presence, movement, installed animations, speech, and thought-balloon responses.",
  tools: {
    allow: [
      "clippy.animations",
      "clippy.queue_status",
      "clippy.release",
      "clippy.show",
      "clippy.hide",
      "clippy.move",
      "clippy.speak",
      "clippy.think",
      "clippy.play",
    ],
  },
  ...(apiKey ? { headers: { "X-API-Key": apiKey } } : {}),
});
