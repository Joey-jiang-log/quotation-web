# Prototype Instructions

Run the local server yourself and open the preview in the browser available to this environment. Do not give the user server-start instructions when you can run it.

Before making substantial visual changes, use the Product Design plugin's `get-context` skill when the visual source is unclear or no longer matches the current goal. When the user gives durable prototype-specific design feedback, preferences, or decisions, record them in `AGENTS.md`.

When implementing from a selected generated mock, treat that image as the source of truth for layout, component anatomy, density, spacing, color, typography, visible content, and hierarchy.

Build app UI in `src/`. Keep `.openai/hosting.json`, `worker/index.js`, `scripts/prepare-sites-build.mjs`, and `tests/sites-worker.test.mjs` intact so the same local prototype can be handed to Sites. Before a Sites handoff, run `npm run build` and `npm run test:sites`; the build must leave `dist/client/index.html`, `dist/server/index.js`, and `dist/.openai/hosting.json`.

## Durable product decisions

- 客户询价上传仅支持 `.xlsx`、`.xls`、`.pdf`、`.png`、`.jpg`。
- 报价上传框必须同时支持点击选择和真实拖拽，并显示拖入、解析中、成功、失败状态。
- PDF 和图片识别后必须复用现有型号与价格映射，不直接信任 OCR 结果生成价格。
- 价格摘要中的“已定位价格列”状态必须始终留在卡片内部，窄宽度下改为独立一行。
