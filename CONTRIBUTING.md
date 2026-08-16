# Contributing to @hokkyss/pptx

Thank you for your interest in contributing to the `@hokkyss/pptx` ecosystem! This monorepo powers a next-generation, 100% isomorphic TypeScript suite for parsing, manipulating, and serializing OpenXML PowerPoint (`.pptx`) presentations.

This guide outlines our development workflow, core architectural guardrails, and pull request guidelines.

---

## 🛠️ Development Setup

### Prerequisites
- **Node.js**: $\ge 18$ (Node.js 22 LTS recommended)
- **pnpm**: $\ge 9$ (pnpm 11+ recommended)

### Clone & Install

```bash
# Clone the repository
git clone https://github.com/hokkyss/pptx-parser.git
cd pptx-parser

# Install all workspace dependencies
pnpm install

# Build all packages & generate TypeScript declarations
pnpm build

# Run all unit and integration test suites
pnpm test
```

---

## 🏗️ Monorepo Architecture

Our packages are strictly decoupled with clear boundaries of responsibility:

| Package | Directory | Role |
| :--- | :--- | :--- |
| **`@hokkyss/pptx`** | `packages/pptx` | High-level fluent SDK (`Presentation`, `Slide`, `SlideMaster`, `TableBuilder`, charts). |
| **`@hokkyss/pptx-reader`** | `packages/pptx-reader` | Isomorphic OpenXML parser converting binary archives to structured AST with 3-tier slide layers. |
| **`@hokkyss/pptx-writer`** | `packages/pptx-writer` | Pure archive XML serializer converting AST into standard `.pptx` archives. |
| **`@hokkyss/pptx-core`** | `packages/pptx-core` | Universal AST schemas, branded units (`Inches`, `Points`, `Emu`), and color/theme types. |

---

## 🛡️ Core Architectural Guardrails

When contributing code, please ensure your changes adhere to our core platform principles:

### 1. 100% Isomorphic & Zero Native Binaries
- All core packages (`pptx`, `pptx-reader`, `pptx-writer`, `pptx-core`) must execute identically in **Node.js, Web Browsers, Cloudflare Workers, Deno, and Bun**.
- **No Node.js-only native dependencies** (e.g. native C++ addons, `node:fs` synchronous calls in runtime code, or desktop-only binaries).
- Use `fflate` for pure JavaScript/WebAssembly compression and decompression.

### 2. Type-Safe Branded Units
- Never pass raw, unbranded numbers for physical coordinates, font sizes, or rotations.
- Always use the branded constructor factories from `@hokkyss/pptx-core` (or re-exported from `@hokkyss/pptx`):
  ```typescript
  import { inches, points, emu, emuDegree } from '@hokkyss/pptx';

  const width = inches(10);        // Branded Inches
  const fontSize = points(24);     // Branded Points
  const rotation = emuDegree(45);  // Branded EmuDegree
  ```

### 3. OpenXML Schema Fidelity (ECMA-376)
- DrawingML and PresentationML XML serializers must strictly follow the element sequence required by the [ECMA-376 OpenXML standard](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) (e.g. `<p:spPr>` $\rightarrow$ `<p:style>` $\rightarrow$ `<p:txBody>`).

---

## 🌿 Branching & Pull Request Workflow

1. **Fork & Branch**: Create a descriptive feature or fix branch from `main`:
   ```bash
   git checkout -b custom-chart-colors
   ```
2. **Make Changes**: Implement your changes along with corresponding unit or integration tests in `tests/`.
3. **Verify Locally**: Before opening a PR, run the local verification suite:
   ```bash
   # 1. Run workspace linter
   pnpm lint

   # 2. Build declaration files and bundles
   pnpm build

   # 3. Run test suites
   pnpm test
   ```
4. **Open a Pull Request**: Submit your pull request against `main`.
   - Ensure the PR title clearly describes the change.
   - Our automated CI will run `Build, Test & Benchmark` and post automated bundle size and runtime performance telemetry on your PR.

---

## 📝 Commit Message Conventions

We follow the [Conventional Commits](https://www.conventionalcommits.org/) specification:

```
<type>(<optional scope>): <description>
```

### Common Types
- `feat`: A new feature or API capability
- `fix`: A bug fix or OpenXML schema correction
- `perf`: Performance optimization or bundle size reduction
- `docs`: Documentation updates or README additions
- `refactor`: Code changes that neither fix a bug nor add a feature
- `test`: Adding or updating test suites
- `chore`: Tooling, workflow, or dependency maintenance

### Examples
- `feat(charts): add dataPointColors slice palette support for pie and doughnut charts`
- `fix(theme): serialize extraClrSchemeLst for Slide Master custom color palette listing`
- `docs: add OpenXML specifications and schema standards references`

---

## 💡 Submitting Large Features

For substantial architectural changes, new chart topologies, or breaking API proposals, we strongly encourage opening a **GitHub Discussion** or **Issue** first. This ensures alignment on design and API ergonomics before writing substantial code.

---

## 📜 License

By contributing to `@hokkyss/pptx`, you agree that your contributions will be licensed under the project's [MIT License](./LICENSE).
