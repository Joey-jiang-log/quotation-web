# Design QA

- source visual truth path 1: `/var/folders/t6/08z8pf4s6svfv04ctb726ghw0000gn/T/codex-clipboard-6df89caf-02f4-41de-a5f5-91f609f13a26.png`
- source visual truth path 2: `/var/folders/t6/08z8pf4s6svfv04ctb726ghw0000gn/T/codex-clipboard-6a1b00d7-d614-4841-85c9-d1ea486058b4.png`
- implementation URL: `https://quotation-rdjvz76d0-jincanjiang729-9470s-projects.vercel.app`
- implementation screenshot path: unavailable
- source pixel dimensions: 1200 × 790 and 1144 × 897
- implementation pixel dimensions: unavailable
- CSS viewport and density: intended desktop viewport 1200 × 790 at 1x; capture unavailable
- state: 批量报价空状态与价格类型已选择状态

## Full-view comparison evidence

两张源截图均可读。实现构建和预览部署成功，但 Codex 内置浏览器与 Chrome 均因管理员安全策略无法访问本地或 Vercel 预览地址，因此没有浏览器渲染截图，不能完成有效并排比较。

## Focused region comparison evidence

未执行。缺少实现截图，无法直接核对价格摘要右侧状态、拖拽悬停态及文件选中态。

## Required fidelity surfaces

- Fonts and typography: blocked，缺少浏览器渲染证据。
- Spacing and layout rhythm: 代码已将价格摘要改为受限三列布局，窄屏改为状态独立一行，但视觉证据缺失。
- Colors and visual tokens: 沿用既有蓝白令牌，新增拖拽蓝色高亮与上传进度条；未完成视觉采样。
- Image quality and asset fidelity: 界面图标继续使用 Phosphor 图标库，没有新增占位图或自制图形。
- Copy and content: 上传说明已限定为 `.xlsx、.xls、.pdf、.png、.jpg`，代码层核对完成。

## Functional evidence

- 生产构建成功；Sites 包装测试 4/4 通过。
- Excel 沿用原工作簿解析流程。
- PDF 使用文字提取，扫描件回退 OCR；PNG/JPG 使用 OCR。
- OCR 预处理会二值化、移除表格横竖线并放大页面。
- 使用用户提供的询价图片离线识别出 10 条 OE/数量行，包括 `90919-02248 × 150`、`27301-2B010 × 150`、`27301-03010 × 30`。
- 浏览器拖拽事件、文件选择器、控制台错误与响应式渲染：blocked，原因同上。

## Findings

- [P1] 浏览器渲染与真实拖拽操作未取得可见证据。
  - Location: 批量报价上传区和价格摘要。
  - Evidence: 两种浏览器均被管理员安全策略阻断。
  - Impact: 无法在当前环境确认最终像素级裁切与拖拽悬停态。
  - Fix: 浏览器策略恢复后，在 1200px 和 620px 两个宽度重新截图并执行文件拖入。

## Comparison history

- Iteration 1: 发现原始 OCR 无法稳定读取表格线内文字。
- Fix: 增加二值化、横竖线移除、高清放大和单区块识别。
- Post-fix evidence: 同一询价图片从 0 条提升到 10 条型号/数量行。
- Iteration 2: 本地和公开预览的浏览器捕获均被安全策略阻断，视觉循环无法继续。

## Follow-up polish

- 浏览器策略恢复后补做 1200px 与 620px 的视觉截图，不影响当前构建与服务端报价能力。

final result: blocked
