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
| [工作流与模板](artifact-template-research-ppt-blue/SKILL.md) | 去身份信息的模板、制作脚本、叙事与验收规则 |
| [单篇论文案例](examples/segment-anything/) | Segment Anything，13 页 |
| [长论文精读案例](examples/sam2-paper/) | SAM 2，21 页 |
| [日常科研汇报案例](examples/sam2-progress/) | SAM 2 复现准备，12 页，备注含追问准备 |

每个案例只保留 PPT、PDF、图文讲稿及来源说明。论文全文、审查日志、设备路径和历史聊天记录不在公开版中。

## 旧讲座不影响使用

本仓库不包含旧人工智能讲座 PPT。制作脚本使用独立的精简母版与目录模板，文字、色彩、字号、版式和逻辑要求已经写入 skill；预览改为科研案例。因此不需要下载旧讲座，也不会继承其中的内容。

## 环境与更新

制作依赖 Codex 提供的 artifact-tool、Node 和 Python；模板使用微软雅黑，当前案例在 Windows PowerPoint 中验证。其他系统需确认字体和渲染器。仅阅读成品 PPT/PDF 不需要制作依赖。

Git 用户可通过 `git pull --ff-only` 更新；ZIP 用户重新下载，再更新个人 skills 目录。仓库内容为手动同步，不会自动上传本地文件。后续只同步审查后的 skill 与 examples，不直接镜像整个本地制作目录。

从首次上传版本更新时，请重新下载或重新克隆本仓库：本次为移除身份信息和旧资料，重新整理了公开分支历史。已有个人修改请先另行保存。

## 分享范围

模板已清除校名、校徽和个人作者属性。案例保留公开论文署名、来源与必要实验条件，不包含私人实验室仪器、资产或运行环境信息。来源图片和论文内容的权利归原权利人。公开前需按 skill 的 [公开版本检查](artifact-template-research-ppt-blue/references/public-sharing.md) 再审查新增内容。
