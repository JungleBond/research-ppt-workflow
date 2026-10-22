# Research PPT Workflow

一套固定蓝白模板的中文科研 PPT 工作流。先写背景、完整讲稿和逐页逻辑，再制作可编辑幻灯片，最后审查证据、图文、排版和讲者备注。

## 下载与使用

1. 点击 **Code → Download ZIP** 并解压，或克隆本仓库。
2. 将 `artifact-template-research-ppt-blue` 整个目录复制到 Codex 的个人 skills 目录，通常为用户主目录下的 `.codex/skills/`；设置过 `CODEX_HOME` 时，使用其下的 `skills/`。
3. 新建 Codex 会话，附上论文并使用下方提示词。浏览案例可直接打开 `打开工作流.html`。

> 使用 $artifact-template-research-ppt-blue，依据附件论文制作约 15 分钟的中文科研汇报 PPT。听众不一定熟悉课题，开头讲清应用背景、已有方法与研究问题，中间逐步深入。先写完整讲稿和逐页逻辑，再按固定蓝白模板排版；每页有相关图表与必要说明，字号不小于 18 磅，重点局部标红，不加多余总结。将可能的追问和回答写入备注。不要视频。逐页检查并直接修复，交付 PPT、预览及对应讲稿。

## 文件入口

| 目录 | 内容 |
|---|---|
| [工作流与模板](artifact-template-research-ppt-blue/SKILL.md) | 蓝白模板、制作脚本、叙事与验收规则 |
| [单篇论文案例](examples/segment-anything/) | Segment Anything，13 页 |
| [长论文精读案例](examples/sam2-paper/) | SAM 2，21 页 |
| [日常科研汇报案例](examples/sam2-progress/) | SAM 2 复现准备，12 页，备注含追问准备 |

每个案例提供 PPT、PDF、图文讲稿及来源说明。

## 环境与更新

制作依赖 Codex 提供的 artifact-tool、Node 和 Python；模板使用微软雅黑，当前案例在 Windows PowerPoint 中验证。其他系统需确认字体和渲染器。仅阅读成品 PPT/PDF 不需要制作依赖。

Git 用户可通过 `git pull --ff-only` 更新；ZIP 用户重新下载。随后将更新后的 `artifact-template-research-ppt-blue` 文件夹复制到个人 skills 目录，并打开新的 Codex 会话使用。
