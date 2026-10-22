# 制作说明与数据格式

## 文件与流程

准备 `evidence.json`、完整讲稿、逐页逻辑，随后生成 `deck.json`。示例项目的 `lecture_plan.json` 是排版前内容，`deck.json` 是加入版式后的内容。两者分开便于先审逻辑。

1. 用 workspace dependency loader 获得 Node、Python、node_modules；不要假定其他设备安装路径相同。
2. 按当前 Presentations skill 要求启动制作，使用其当前版本的本地生成和 finalization 流程。
3. 设置 `RUNTIME_NODE_MODULES` 为返回的 node_modules。
4. 执行：`node scripts/build_deck.mjs deck.json build目录`。
5. Windows PowerPoint：`powershell -File scripts/render_powerpoint.ps1 -BuildDirectory build目录`。它读取 `base.pptx`，规范字体和边距，导出 `candidate.pptx`、PDF、逐页 PNG 与实际文字边界。
6. 执行 `python scripts/audit_and_guide.py build目录`。通过后逐页看 PNG 并检查讲稿映射，有问题回改 `deck.json` 再制作。
7. 执行 `node scripts/finalize.mjs build目录 输出.pptx 当前Presentations技能目录 Python路径`。使用单独的新最终文件，不覆盖源模板。
8. 执行 `python scripts/audit_and_guide.py build目录 输出.pptx`，确认最终文件与备注；复制 PDF、HTML 讲稿到交付目录。
9. 执行 `powershell -File scripts/verify_final_powerpoint.ps1 -Presentation 输出.pptx -ReportDirectory 私有验证目录`。在副本中重新打开最终文件、渲染所有页、测试原生图表的数据修改和缩放，再恢复；原文件哈希必须不变。

这里的 shell 命令是结构示意，实际路径作为独立参数传入并正确转义。PowerPoint 脚本只关闭自己打开的文档，不退出整个应用。若没有 PowerPoint，选择可用渲染器并如实记录未完成的软件验证。

## JSON 结构

顶层：`title`、`shortTitle`、`chapters`、`noVideo`、`sources`、`pages`。

来源：`id`、`title`、`url`、可选 `locator`。

每页：

- `title / chapter / minutes`：标题、从 0 开始的章节索引、预计时间。
- `question / previous / next`：本页问题、上一页关系、下一页理由。
- `speech / cue`：完整口播与指图操作。
- 关键页的预期追问、简短回答、证据定位和回答边界，可追加到 `cue` 中的 `【问答准备】` 段；现有引擎会写入备注与讲稿，不添加到投影正文。较长准备可另存 `答辩准备.md`，其中页码须随最终页序更新。不要把这些准备话术当作口播全文或自动生成一张问答页。
- `sources / sourceLocator / claimIds`：来源 ID、具体定位和证据台账 ID。
- `shortSource`：屏幕来源短句；可留空，但正文研究页原则上有出处。
- `fullBleed`：封面、目录、感谢页设 true；其余自动使用模板导航、标题与页脚。
- `objects`：按叠放顺序列内容对象。坐标使用 1280×720 CSS px。
- `tocEntries`：目录页的 3—5 个章名。设 `fullBleed:true`；渲染时复制 `assets/toc_reference.pptx` 原生形状，替换文字，保留本页讲稿备注。无需手工绘制仿制目录。
- `emphasis`：本页需要标红加粗的完整词组或数值数组，通常 1—3 项。渲染器处理原生文字和表格，记录实际命中位置；不将一位数字匹配到另一个数值内部。检查实际命中，避免不匹配或重复强调过多。

对象公共字段：`type / x / y / w / h`。类型：

- `text`：`text,size,color,bold,align,valign`。size ≥24 px。文字尽量按语意换行；不依靠手工空格对齐。
- `rect`：`fill,line`。仅用于匹配模板中的区块背景，不用来拼装插画。
- `image`：`file,role,source,sourceRegion,fit,crop`。文件路径相对 deck.json，默认 contain。可选 crop 为 left/top/right/bottom 四个 0—1 数值，表示从原图各边移去的比例，两侧之和须小于 1；使用 PowerPoint 原生裁切，原始图片仍保留在文件中。裁切坐标与原图大小进入 sourceRegion，裁切窗口按保留区域的宽高比放置并实际渲染核对。不要裁掉坐标轴、完整连接线或改变结论的画面部分。
- `table`：`values,widths,size`；首行蓝底、交替浅色行、微软雅黑、可编辑单元格。
- 表格可设 `alignments` 按列指定 left/center/right。未指定时，纯数值列及其表头居中，文字列左对齐。状态列可显式居中；不把长解释列强行居中。
- `chart`：`chartType,config,source,role`；config 使用 artifact-tool 当前图表配置，包含 categories、series、轴和数据标签，所有图表字体 ≥24 px。PowerPoint 目标可使用原生图表中的字面数据；这不等于含有 Excel 工作簿。实际编辑能力由副本中的修改和缩放测试验证；若明确要求通过“编辑数据”打开工作表，需另外生成并验证内嵌工作簿，不能只凭结构检查声称支持。
- `thanks`：保留模板白底、蓝色横带与阴影，只放“谢谢大家”。

此引擎不强制把每页排成一样。对象位置由具体内容决定；可复用布局起点参考 design-and-logic.md。若任务需要新的对象，按当前 artifact-tool API 扩展，而不是把数据图降级为截图。

## 工具边界

自动检查涵盖原生字体、文字边界、表格边界、备注一致、原生图表／表格存在和禁用媒体。它不证明科研结论正确、照片主体裁切合理或叙事顺畅，必须由逐页审阅补充。

工具当前处理静态 PPT 与转场。视频任务需要按当前 Presentations 能力另行嵌入与实际播放检查，不能在本引擎里放一个外链当作内嵌完成。默认无视频因此优先保证这一条路径可靠。


## 公开交付

对外发布前按 public-sharing.md 审查内容，再用 sanitize_office_metadata.py 输出独立的分享副本。此辅助脚本使用 Python 的 lxml 与 pypdf。清理后复查结构和实际渲染；后续再经 Office 保存可能重新写入登录者信息。
