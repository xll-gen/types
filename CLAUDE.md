# Claude Code Instructions

**This file is intentionally minimal.** All durable agent guidance — the schema-source-of-truth rule, co-change clusters, conversion logic, and improvement backlog — lives in [`AGENTS.md`](./AGENTS.md).

Before doing anything in this repository:

1. Read **[AGENTS.md](./AGENTS.md)** in full. It is the single source of truth.
2. Treat `go/protocol/protocol.fbs` as the schema source of truth: any change there requires regenerating both the C++ header (`include/types/protocol_generated.h`) and the Go files in `go/protocol/` in the same commit.
3. If your change crosses the `xll-gen` or `shm` repo boundary, read those repos' `AGENTS.md` too.
4. Do **not** add project-specific guidance to this file. Add it to `AGENTS.md` so every agent tool (Claude Code, Codex, Cursor, Aider, etc.) sees it.

Updating `CLAUDE.md` to anything other than this redirect is a policy violation; update `AGENTS.md` instead.
