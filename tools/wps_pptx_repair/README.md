# WPS 演示文稿跨系统文字错位检测与修复工具

> 专治：Windows 上做的 PPTX，拷到**统信 UOS / 麒麟（Kylin）** 上用 **WPS 演示**打开后，
> 文字跑位、溢出色块、行距变大、断行奇怪、表格被撑开、符号变字母……

```
tools/wps_pptx_repair/
├── wps_pptx_repair.sh     # 【主程序】自包含 shell 脚本，拷这一个文件即可用
├── wps_pptx_repair.py     # 与 .sh 完全同逻辑的 Python 源码版（便于阅读/二次修改）
├── README.md              # 本文档
└── examples/
    └── make_sample.py     # 生成一份“典型错位”样例，用于体验/演示
```

> `.sh` 是**单文件自包含**的：完整实现内嵌在脚本里，运行时释放到临时文件交给
> 系统自带的 `python3` 执行。统信 UOS / 麒麟默认预装 python3，且本工具只用
> 标准库、不需要 pip 安装任何第三方包，所以把 `wps_pptx_repair.sh` 一个文件拷
> 到目标机器就能直接跑（离线可用）。

---

## 一、为什么会出现“错位”

演示文稿本身没有“绝对坐标的文字”，文字位置是**字体度量 + 文本框内边距 + 自动调整策略**共同算出来的。
Windows 与 Linux/WPS 只要有一项不同，累积起来就是肉眼可见的错位。按影响从大到小：

| # | 原因 | 说明 |
|---|---|---|
| 1 | **字体缺失/替换** | 宋体、微软雅黑、黑体、楷体、仿宋、Calibri、Cambria 等在 Linux 上通常没有。WPS 交给 fontconfig 回退到文泉驿/思源/Noto/DejaVu，字宽、字重、字面全变 → 换行位置变、行数变、文字溢出。 |
| 2 | **行高（字体度量）差异** | “单倍行距”按字体自身 ascent/descent 计算。宋体约 1.0~1.2 em，而思源黑体/思源宋体约 **1.44 em**。同一段文字换字体后被“撑高” 20%~40%，多行文本整体下移。 |
| 3 | **中英文混排** | 中文 Windows 字体自带西文字形；Linux 回退后西文字宽不同，中英混排处断行位置改变。 |
| 4 | **自动缩放残留** | PowerPoint 保存时会把自己的排版结果写成 `normAutofit fontScale` / `lnSpcReduction`。WPS 按另一套字体度量重算或直接套用比例，与 Windows 对不上。 |
| 5 | **内边距 / 单元格边距默认值不同** | 文本框 `lIns/rIns/tIns/bIns`、表格单元格 `marL/marR/marT/marB` 缺省时各实现默认值可能不同，文字整体偏几个点，可用宽度变化又会改变换行。 |
| 6 | **不换行文本框 / 随文字自动调整高度** | `wrap="none"` 的文本框宽度完全依赖字体宽度；`spAutoFit` 让形状高度由文字决定。字体一变，宽度/高度就变，直接遮挡相邻元素。 |
| 7 | **符号字体缺失** | Wingdings / Webdings / Symbol / Marlett 用“字母码位映射图形”显示 ✓ → ★。Linux 没有这些字体时，同一个码位显示成字母——看起来就是乱码/错位。 |
| 8 | **语言标记与断行规则** | run 的 `lang="en-US"` 会让 WPS 按西文规则处理中文断行、行首禁则、中英文间距。 |
| 9 | **其他** | 组合形状非等比缩放、竖排（`vert="eaVert"`）、文字变形（WordArt）、非标准页面比例被 WPS “适应页面”缩放、旧版 `.ppt` 二进制格式。 |

本工具对上述 9 类原因逐项体检（规则 R001~R020），能自动修的自动修，不能自动修的明确告诉你为什么、该怎么办。

---

## 二、快速开始（交互式）

在 UOS / 麒麟上，把整个 `wps_pptx_repair` 目录拷过去，双击终端执行：

```bash
cd wps_pptx_repair
chmod +x wps_pptx_repair.sh
./wps_pptx_repair.sh
```

（也可以用 `sh wps_pptx_repair.sh`；若系统 python 不在 PATH，可指定
`WPS_REPAIR_PYTHON=/usr/bin/python3 ./wps_pptx_repair.sh`。）

看到主菜单：

```
[1] 选择演示文稿（支持拖入路径 / 目录 / 通配符）
[2] 环境与字体体检
[3] 扫描诊断（只读，不改任何文件）
[4] 修复向导（扫描 → 选择修复项 → 备份并修复）
[5] 渲染验证（LibreOffice 对比修复前后）
[6] 字体兼容方案（导入 Windows 字体 / 生成或撤销字体别名）
[7] 帮助：原因说明、修复原理与推荐流程
[0] 退出
```

**推荐流程**：

1. `[2]` 看环境：确认系统、WPS、中文字体是否正常；
2. `[1]` 把课件拖进终端（或输入目录），`[3]` 扫描，先看问题分布；
3. 缺字体很多时，去 `[6]` **优先导入 Windows 字体**（最保真）；没有 Windows 字体就跳过，后面的修复向导会改写字体名；
4. `[4]` 修复向导，选「推荐修复」，字体策略选「改写文件内字体名」；
5. 打开 `xxx.fixed.pptx` 在 WPS 里逐页检查；
6. 个别页面仍溢出，再跑一次 `[4]`，勾选 **R009（溢出缩字）** 或 **R014（固定行距）**。

> 工具**默认不覆盖原文件**：输出 `原名.fixed.pptx`，原文件另存为 `原名.orig-时间戳.pptx`，
> 并生成 `原名.report.md` / `原名.report.json` 报告。确认没问题后再手动替换。

### 想先体验？

```bash
python3 examples/make_sample.py /tmp/demo        # 生成一份“典型错位”样例
./wps_pptx_repair.sh /tmp/demo/sample_broken.pptx
```

---

## 三、三种字体策略，怎么选

| 策略 | 做法 | 优点 | 缺点 / 适用 |
|---|---|---|---|
| **改写文件内字体名**（默认） | 把 PPTX 里的 `宋体` 改成 `Noto Serif CJK SC`、`微软雅黑`→`Noto Sans CJK SC`、`Calibri`→`Carlito`、`Arial`→`Liberation Sans`……同时把中文 run 的 `<a:ea>` / 西文 `<a:latin>` 拆开写清楚 | 生成的文件到任何机器上都不缺字；可独立分发；不需要改系统 | 字体风格会变（宋体→思源宋体），需要接受观感变化；再拿回 Windows 打开也是新字体 |
| **fontconfig 别名**（不改文件） | 在 `~/.config/fontconfig/conf.d/99-wps-windows-fonts.conf` 写别名，把缺失字体统一映射到本机字体 | 文件保持原样；一次配置，所有文档受益；适合批量/长期 | 只对**本机**生效，换机器又错位；个别 WPS 版本自带字体替换表可能覆盖别名；**WPS 需重启** |
| **导入 Windows 字体**（最保真） | 从 `C:\Windows\Fonts`（双系统挂载、U 盘、共享目录）复制字体到 `~/.local/share/fonts/windows-fonts/`，刷新 fontconfig | 字体度量与 Windows **完全一致**，观感不变，绝大多数错位直接消失 | 需要能拿到 Windows 字体；字体版权需自行确认；只对本机生效 |

交互菜单 `[6]` 里可随时导入字体、生成/撤销别名；`[4]` 修复向导会询问字体策略。

---

## 四、命令行用法（可脚本化 / 批量）

```bash
# 只体检（不改文件），有问题时退出码 2
./wps_pptx_repair.sh 课件.pptx --scan
./wps_pptx_repair.sh ./slides --scan --json > scan.json

# 直接修复（推荐项）
./wps_pptx_repair.sh 课件.pptx --fix -y
./wps_pptx_repair.sh ./slides --fix -y --out-dir ./fixed

# 修复并渲染验证
./wps_pptx_repair.sh 课件.pptx --fix -y --render-check

# 全部规则（含 R009 溢出缩字 / R014 固定行距等改变视觉的项）
./wps_pptx_repair.sh 课件.pptx --fix -y --rules all

# 只修字体，或只修版式
./wps_pptx_repair.sh 课件.pptx --fix -y --rules font
./wps_pptx_repair.sh 课件.pptx --fix -y --rules layout

# 字体：先导入 Windows 字体，再修（最保真）
./wps_pptx_repair.sh --import-fonts /mnt/c/Windows/Fonts
./wps_pptx_repair.sh 课件.pptx --fix -y --font-mode install --fonts-dir /mnt/c/Windows/Fonts

# 字体：不改文件，只生成系统别名 / 撤销别名
./wps_pptx_repair.sh 课件.pptx --install-aliases
./wps_pptx_repair.sh --remove-aliases

# 规则说明
./wps_pptx_repair.sh --list-rules
./wps_pptx_repair.sh --explain R014
```

主要参数：

| 参数 | 说明 |
|---|---|
| `paths...` | PPTX 文件 / 目录 / 通配符，可多个 |
| `--scan` / `--fix` | 只体检 / 执行修复（非交互） |
| `--rules default\|all\|font\|layout\|R001,R009` | 修复规则集 |
| `--font-mode rewrite\|alias\|install\|none` | 字体策略，默认 `rewrite` |
| `--fonts-dir DIR` | `--font-mode install` 时使用的 Windows 字体目录 |
| `--out-dir DIR` / `--output FILE` / `--suffix .fixed` | 输出位置 |
| `--overwrite` | 覆盖原文件（仍会先备份） |
| `--no-backup` | 不备份（不建议） |
| `--render-check` | 用 LibreOffice 渲染对比修复前后，确认没丢内容 |
| `--line-ratio 1.2` | R014 固定行距的倍率 |
| `--shrink-min 0.25` | R009 缩字下限 |
| `--json` | 标准输出追加 JSON 报告 |
| `--import-fonts DIR` / `--install-aliases` / `--remove-aliases` | 字体方案 |

**退出码**：`0` 成功；`1` 有文件失败；`2` 体检发现高危问题。

---

## 五、检测/修复规则一览

| 规则 | 级别 | 默认 | 问题 | 自动修复做了什么 |
|---|---|---|---|---|
| R001 | 高 | ✅ | 引用了本机不存在的字体 | 改写为度量/风格最接近的已安装字体；清掉失配的 panose |
| R002 | 高 | ✅ | 中文 run 没有显式 `<a:ea>`；中文字体兼任西文 | 补 `<a:ea>`；把 `<a:latin>` 拆成度量兼容的西文字体 |
| R003 | 中 | ✅ | 主题 `majorFont/minorFont` 缺中文（Hans/Hant）映射或 `ea` 为空 | 补写映射到本机中文字体 |
| R004 | 低 | ✅ | 文件内嵌字体（WPS 通常忽略 `fntdata`） | 仅提示：建议装原字体或用 R001 替换 |
| R005 | 高 | ✅ | 符号字体缺失（Wingdings/Webdings/Symbol/Marlett） | **不自动替换**（会变成字母）；提示安装原字体或改符号 |
| R006 | 中 | ✅ | 文本框内边距未显式设置 | 写入 PowerPoint 默认值（左右 0.1in、上下 0.05in） |
| R007 | 高 | ✅ | PowerPoint 自动缩放残留（`fontScale`/`lnSpcReduction`） | 用实际字体度量重算：放得下就清残留，仍溢出就写实测比例 |
| R008 | 中 | ❌ | 形状随文字自动调整高度（`spAutoFit`） | 改为 `noAutofit` 固定尺寸（会改变版式行为，需确认） |
| R009 | 高 | ❌ | 按 Linux 实际字体度量估算会溢出 | 写入实测需要的 `normAutofit` 缩放比例（≥25%） |
| R010 | 中 | ✅ | `wrap="none"` 文本框宽度不足 | 以中心为基准扩宽到实测文本宽度（不超出页面）；放不下就改为允许换行 |
| R011 | 中 | ✅ | 表格单元格边距未显式设置 | 写入 PowerPoint 默认单元格边距 |
| R012 | 中 | ❌ | 表格行高不足以容纳文本 | 按实测内容增高行高 |
| R013 | 低 | ✅ | 中文文本语言标记不是中文 | 按内容设置 `lang="zh-CN"`（日/韩保持各自语言） |
| R014 | 中 | ❌ | 行距依赖字体度量（单倍/百分比行距） | 把行距固定为磅值（默认 1.2×字号，可用 `--line-ratio` 调） |
| R015 | 低 | ✅ | 存在手动换行 `a:br` | 仅提示（硬回车无法自动重排） |
| R016 | 低 | ✅ | 组合形状缩放比例不一致 | 仅提示（建议在 WPS 里取消组合后重新组合） |
| R017 | 低 | ✅ | 竖排文字 / 文字变形（WordArt） | 仅提示（如需完全一致请改横排普通文本） |
| R018 | 高 | ✅ | 旧版 `.ppt/.dps/.wps` 二进制格式 | 提示先另存为 `.pptx`（`soffice --headless --convert-to pptx`） |
| R019 | 中 | ❌ | 替换字体缺部分字形（生僻字/繁体/符号） | 改写为字形覆盖更全的中文字体（如思源黑体） |
| R020 | 提示 | ✅ | 幻灯片尺寸非标准比例 | 仅提示（检查 WPS 页面设置/缩放） |

“默认”列 = 交互式「推荐修复」与命令行 `--rules default` 包含的规则。
标 ❌ 的规则会改变视觉观感，需要你明确选择。

---

## 六、输出、报告与回滚

| 产物 | 说明 |
|---|---|
| `原名.fixed.pptx` | 修复后的文件（默认同目录，可用 `--out-dir` 改） |
| `原名.orig-YYYYmmdd-HHMMSS.pptx` | 原文件备份（首次修复前自动生成） |
| `原名.report.md` | 人类可读报告：环境、问题明细、修复记录、复检结果、人工检查清单 |
| `原名.report.json` | 机器可读报告：findings / changes / metrics，便于批量统计 |

**回滚**：

* 文件：用备份文件覆盖修复版即可；
* 字体别名：`./wps_pptx_repair.sh --remove-aliases`；
* 导入的 Windows 字体：删除 `~/.local/share/fonts/windows-fonts/` 后 `fc-cache -f`。

---

## 七、常见问题（FAQ）

**Q1：扫描报「字体缺失」，但 WPS 里中文显示正常？**
显示正常 ≠ 度量一致。回退字体能显示中文，但字宽和行高已经变了，足以让多行文本重新换行。以 R001 报告为准。

**Q2：生成别名后 WPS 还是原样？**
① WPS 必须**完全退出后重启**；② 确认文件写在 `~/.config/fontconfig/conf.d/`；③ 运行
`fc-match 宋体` 看是否已指向新字体；④ 少数 WPS 版本自带字体替换表优先级更高，此时改用「改写文件字体名」。

**Q3：修复后有的文字变小了？**
说明勾选了 R009（溢出缩字）。它按实测溢出量缩字，属于“保版式、牺牲字号”的取舍；也可以改用 R014 固定行距 + 不缩字。

**Q4：为什么工具不直接把宋体换成“看起来一样”的字体？**
Linux 上没有与宋体/微软雅黑**度量完全一致**的自由字体；思源系列是最接近且字形覆盖最全的选择。真正“完全一致”只有导入原版权重字体（`[6]` 导入）。

**Q5：表格、图表、SmartArt 里的文字也会查吗？**
表格会做单元格边距/行高/溢出检查与修复；图表、SmartArt 的字体引用会被 R001/R002 处理（其文字排版由 WPS 内部计算，工具不做几何修复）。

**Q6：会不会改坏文件？**
工具只做三类操作：① 改字体名/补 `ea`、`lang` 等属性；② 补内边距、单元格边距；③ 在明确启用时调整 autofit / 行高 / wrap/宽度。**不重排版式、不改文字内容、不改颜色图片**。所有写操作前自动备份，并输出 `.report.json` 记录每一处修改。

**Q7：为什么用纯标准库，不用 python-pptx？**
UOS/麒麟 上 pip 环境常受限；直接操作 OOXML 更可控，也能保证老版本 Python（3.7+）可用。代价是需要自己处理字体度量——这部分工具已内置一个极简 TTF/OTF/TTC 度量读取器（cmap/hmtx/hhea）。

**Q8：扫描很慢 / 溢出估算准吗？**
每个文本框会按实际替代字体做换行模拟，页面多时秒级到十几秒。估算与 WPS 真实渲染可能有 1~2% 误差，所以 R009 默认关闭，且报告里给的是“风险”而非“一定溢出”。最终请以 WPS 打开效果为准。

**Q9：能处理 `.ppt`（97-2003）吗？**
不能直接改二进制格式。先另存为 `.pptx`：`soffice --headless --convert-to pptx 旧文件.ppt`，再运行本工具。

**Q10：渲染验证通过就代表 WPS 里一定没问题吗？**
不是。渲染验证用 LibreOffice，只用于确认“页数没变、文字没丢”，WPS 的实际排版仍需人工过一遍。

---

## 八、工作原理（简版）

1. **解包**：按 OPC 规范读取 zip 内全部 part，只重新序列化被修改的 XML；未修改的 part 按原字节复制。
2. **建模**：解析 `presentation.xml` 页序与页面尺寸，沿 slide→layout→master→theme 关系链解析主题字体、占位符几何与文本样式继承（`lstStyle`/`pPr/defRPr`/`rPr`）。
3. **字体决策**：用 `fc-list` 判断字体是否安装，用 `fc-match` 预测 WPS 实际会用的回退字体；缺失字体按内置映射表 + 字形/风格分类选择替代字体。
4. **度量估算**：内置 TrueType/OpenType（含 TTC）解析器读取 `cmap` 字形映射、`hmtx` 字宽、`hhea` 行高；按“中文用 `<a:ea>`、西文用 `<a:latin>`”逐字计算宽度，做贪心换行（中文字字可断、西文按词断），叠加行距/段距/内边距/autofit，得到所需高度与行宽，与文本框可用区域比较。
5. **修复**：对勾选的规则改写对应属性（必要时按 schema 顺序插入子元素），然后保存、复检、出报告。

---

## 九、`.sh` 与 `.py` 的关系

* `wps_pptx_repair.sh`：交付主文件。自身是 shell 脚本，内嵌完整 Python 实现；
  运行时在临时目录释放并执行，退出后自动清理，不污染系统。
* `wps_pptx_repair.py`：同一份逻辑的源码版，方便阅读、审计、二次开发；
  两者行为一致（`.sh` 由 `.py` 生成）。
* 若你的环境**完全没有 Python**，可告知，我可以再提供一版纯 shell/perl 的
  精简实现（功能覆盖字体内边距/自动缩放等主要项，溢出估算精度会下降）。

## 十、目录与环境要求

* **Python 3.7+**：`.sh` 运行时需要系统自带的 python3（UOS/麒麟/Ubuntu 默认预装；
  也可用 `WPS_REPAIR_PYTHON=/path/to/python3` 指定），**不需要 pip 装任何包**；
* **fontconfig**（`fc-list` / `fc-match`，绝大多数桌面发行版自带）——没有也能跑，只是无法判断“字体是否已安装”，会退化为按内置映射替换；
* 可选：**LibreOffice** + `pdftotext`（渲染验证）；
* 可选：**WPS Office for Linux**（工具只做检测，不依赖 WPS 进程）。

隐私说明：工具完全离线运行，不上传任何文件；只在本地读写你的 PPTX、用户字体目录与 fontconfig 配置。

---

## 十一、已知限制

* 图表/SmartArt/艺术字内部的文字只做字体替换，不做几何重排；
* 溢出估算是“按字体度量的静态计算”，不等于 WPS 的最终渲染；
* 组合形状内的非等比缩放只提示、不自动烘焙坐标；
* 中文 Windows 字体与自由字体不可能字字等宽，极端紧凑的版式仍需在 WPS 里微调；
* fontconfig 别名不改变文件，因此把文件发给别人仍可能错位——要“文件自带一致性”请用改写模式。
