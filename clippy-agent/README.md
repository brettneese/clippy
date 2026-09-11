# clippy-agent

This is the [eve](https://eve.dev) agent behind the project's authentic Clippy.
Its standing context gives the agent Clippy's eager, playful Office Assistant
personality and makes the Windows XP Clippy MCP connection its default
user-facing interface.

## Getting started

First, run the development server:

```bash
eve dev
```

The development TUI opens an interactive session where you can send messages to your agent.

Start by editing `agent/instructions.md` to define the agent's identity, purpose, tone, and response guidelines. Configure its model and runtime behavior in `agent/agent.ts`.

Add capabilities under `agent/`, including tools, connections, channels, skills, subagents, and schedules. eve reloads your changes as you work.

## Connect to the authentic Clippy MCP server

The Clippy connection targets the public Streamable HTTP endpoint at
`https://pleasing-unicorn-legally.ngrok-free.app/mcp`. Start the transport
proxy on the Mac host in a separate terminal before running the agent:

```bash
npm run clippy:mcp
```

Keep an ngrok tunnel forwarding that loopback listener in another terminal:

```bash
ngrok http 3212
```

The configured ngrok domain must forward to `http://127.0.0.1:3212`. The proxy
accepts modern MCP requests from eve and translates them to the legacy MCP
version negotiated by the repository's SSH stdio bridge. The bridge connects
to the persistent Microsoft Agent controller on the visible Windows XP
desktop. The XP service must already be listening on `127.0.0.1:3211`; see
[`../xp/README.md`](../xp/README.md) for its setup and lifecycle.

Set `CLIPPY_MCP_URL=http://127.0.0.1:3212/mcp` to bypass ngrok for local-only
development, or use another Streamable HTTP endpoint. If the proxy is
protected with `MCP_PROXY_API_KEY` (or `mcp-proxy --apiKey`), set the same value
as `CLIPPY_MCP_API_KEY`; the connection sends it in the `X-API-Key` header.
The checked-in public URL is transport-secured but is not authenticated unless
that API key is configured. Do not leave this desktop-control endpoint public
without authentication.

## Learn more

To learn more about eve, explore these resources:

- [eve documentation](https://eve.dev/docs) — learn about eve's features and authoring APIs.
- [Build an Agent tutorial](https://eve.dev/docs/tutorial/first-agent) — build and deploy an agent step by step.
- [eve on GitHub](https://github.com/vercel/eve) — view the source and contribute.

## Deploy on Vercel

Deploy your agent to [Vercel](https://vercel.com) from the project root:

```bash
eve deploy
```

`eve deploy` links a Vercel project if needed and deploys the agent to production. See the [eve deployment documentation](https://eve.dev/docs/guides/deployment/vercel) for authentication, environment variables, and deployment options.
