import { defineMcpClientConnection } from "eve/connections";

const url =
  process.env.CLIPPY_MCP_URL ??
  "https://pleasing-unicorn-legally.ngrok-free.app/mcp";
const apiKey = process.env.CLIPPY_MCP_API_KEY;

export default defineMcpClientConnection({
  url,
  description:
    "The authentic Microsoft Office Clippy running on Windows XP. Show, hide, move, animate, speak, or display thought-balloon text through the real desktop character.",
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
