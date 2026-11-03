# Fall Academic PPT

面向导师交流的中文学术汇报 Codex skill：从研究材料确定主线与逐页脚本，沿用模板制作或修改可编辑 PPT，配套讲稿并核对公式、来源与版面。

[English](#english) · [技能入口](SKILL.md) · [MIT 许可证](LICENSE)

## 能做什么

- 先核对材料、最新用户版本与证据，再组织“问题—方法—验证”的研究主线
- 按阶段工作：只要脚本时不提前制作 PPT；采纳后直接制作；局部修改保留手动排版和页序
- 讲清模型分工、输入输出、公式符号及计算用途；区分既有方法、拟扩展和实际实验结果
- 沿用模板、原生对象与公式；正文默认在空间允许时使用 1.5 倍行距，底部总结框加粗、单倍行距、水平和垂直居中
- 按需协调图标、正文编号与 APA 书目，核对正文、隐藏页、讲稿和备注

这些是 Fall 的个人默认偏好；当前用户要求和指定模板优先，适合在此基础上修改成自己的工作流。

## 安装

在 Codex 中提出以下请求，让内置 skill installer 将仓库根目录安装为 `fall-academic-ppt`：

```text
请从 https://github.com/Nubfall/fall-academic-ppt 安装 skill，技能位于仓库根目录，安装名称为 fall-academic-ppt
```

也可以将仓库克隆到本机技能目录下的 `fall-academic-ppt` 文件夹：

```powershell
git clone https://github.com/Nubfall/fall-academic-ppt.git "$HOME/.codex/skills/fall-academic-ppt"
```

如果已有同名技能，先备份并比较，不要覆盖个人修改。若设置了 `CODEX_HOME`，使用该目录下的 `skills`。安装后在新对话或重启 Codex 后使用。

## 使用示例

```text
使用 $fall-academic-ppt，先读取我的研究方案和模板，整理适合导师交流的逐页脚本，先讨论脚本，确认后再做 PPT
```

```text
使用 $fall-academic-ppt，以我最新保存的 PPT 为基础，只调整方法页的公式解释和留白，保留页序、手动行距和隐藏页
```

## 运行条件与检查器

技能说明本身不依赖 Python。实际制作需要环境中可用的 PowerPoint、演示文稿工具或其他已验证的制作方式；本仓库不附模板、字体、图标、付费软件或其他 skill。

技能会按任务需要查找 `presentations`、`ppt-master`、`nature-polishing`、`nature-writing`、`reader-first-writing`、`anti-defensive-writing` 等能力。缺失时应说明并使用可用替代方式，不能声称执行了不存在的工具。

随附检查器仅依赖 Python 3.9+ 标准库，在仓库目录运行：

```powershell
python -X utf8 scripts/inspect_pptx.py "presentation.pptx" --output "inspection.json"
```

它按演示顺序清点正文／隐藏页、文本、备注、原生公式，并提示底部框显式格式等线索，不修改输入 PPT。警告不等于失败；主题继承、分组坐标和兼容公式仍需实际渲染复核。结构读取成功不代表事实、版面或公式正确。


## 许可证与引用

本仓库原创说明与检查脚本按 MIT 许可证开放。外部工具、素材、字体与参考资料遵守各自许可证，不包含在本仓库授权中。

引用格式：Nubfall. (2026). *Fall Academic PPT* (Version 1.2.0, public edition) [Codex skill]. https://github.com/Nubfall/fall-academic-ppt

## English

Fall Academic PPT is a Codex skill for Chinese academic presentations and speaker scripts intended for advisor discussions. It helps organize source materials into a coherent research argument, reuse an existing template, explain models and equations, and check consistency across slides, notes, references and scripts.

### Workflow and defaults

Start with current evidence and the latest user-saved deck. Prepare a script first when requested; create slides after adoption; preserve manual edits during local revisions. Explain the research question, inherited methods, proposed changes and validation separately. Match the supplied template, retain editable objects and equations, and use citations where they directly support a model, formula or source. Body text defaults to 1.5 line spacing when space permits; bottom summaries use bold text, single spacing and horizontal/vertical centering. These are personal defaults, overridden by the current user's instructions and template.

### Installation and use

Ask Codex to install `https://github.com/Nubfall/fall-academic-ppt`, using repository path `.` and installation name `fall-academic-ppt`. Alternatively clone this repository into `~/.codex/skills/fall-academic-ppt` (or `$CODEX_HOME/skills/fall-academic-ppt`). Back up and compare an existing installation before replacing it. Start a new conversation or restart Codex after installation.

Example: “Use $fall-academic-ppt to read my research plan and template, prepare a page-by-page Chinese speaker script, and wait until I adopt the script before creating slides.”

### Requirements and limits

The instructions require no Python runtime. Slide production requires presentation tools available in the host environment. Related skills are discovered as needed and are not bundled; unavailable tools must be disclosed and replaced with verified alternatives where possible. No templates, fonts, icons or commercial software are included.

The read-only inspector uses the Python 3.9+ standard library:

```text
python -X utf8 scripts/inspect_pptx.py presentation.pptx --output inspection.json
```

It reports slide order, hidden slides, text, notes, native equations and explicit formatting clues. It does not modify the input. Warnings need review and do not imply failure. Inherited styles, grouped coordinates and compatibility branches require visual inspection; structural success does not establish scientific or layout correctness.

### Public edition, attribution and license

Original repository instructions and scripts are MIT-licensed. External tools, assets, fonts and references retain their own licenses. To cite this skill, use the citation above; version 1.2.0 refers to the public edition.
