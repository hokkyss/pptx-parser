## 📋 Pull Request Description

### Summary of Changes
<!-- Provide a clear and concise summary of what this PR introduces or fixes -->

### Target Package(s)
- [ ] `@hokkyss/pptx` (Fluent SDK)
- [ ] `@hokkyss/pptx-reader` (Parser)
- [ ] `@hokkyss/pptx-writer` (Serializer)
- [ ] `@hokkyss/pptx-core` (AST & Units)
- [ ] Documentation / Tooling

---

## 🧪 Verification & Quality Checklist

- [ ] `pnpm lint` passes with 0 errors across all workspace packages
- [ ] `pnpm build` compiles cleanly and generates type definitions (`.d.ts`)
- [ ] `pnpm test` runs and all unit & integration test suites pass
- [ ] Type-safe branded units (`inches()`, `points()`, `emu()`, `emuDegree()`) are used exclusively for physical coordinates and dimensions
- [ ] Code is 100% isomorphic (compatible with Node.js, Web Browsers, Cloudflare Workers, Deno, and Bun)
- [ ] OpenXML XML elements follow ECMA-376 schema sequence and standard naming
