# 无 BP 公司简报工作流（公开信息尽调建档）

> WorkBuddy / Claude 风格 Agent 的技能包：针对**没有商业计划书、没有官网、没有产品资料**的
> 极早期/低信息透明度公司，纯用公开渠道信息搭出一份可向上汇报、可指导尽调访谈的公司简报。

## 这个技能解决什么

拿到一个公司名（可能只有品牌名/简称/英文名），没有任何材料时的预筛建档：

**六阶段工作流**：
```
输入（公司名/线索）
  → P0 实体锚定与消歧（锁定统一社会信用代码，排查同名主体——最大翻车点是查错公司）
  → P1 七类信息采集（工商治理/融资动态/团队线索/技术知识产权/行业市场/司法合规/区域政策）
  → P2 交叉验证与证据分级（五级证据等级，双源交叉）
  → P3 报告撰写（五章骨架 + 开篇声明 + 小结）
  → P4 交付自检（10 项清单，任一不过即返工）
  → P5 归档与记忆
```

**核心纪律**：
- **实体锚定优先**：全文以统一社会信用代码为唯一锚点，同名/近似名主体必设专节辨析
- **事实与推测分离**：每条信息标证据等级（`verified`/`corroborated`/`self_claimed`/`unverified`/`gap`），
  推测显式标注，查不到写"未公开披露"，禁止编造
- **双源交叉**：工商关键项企查查 × 天眼查双源核对
- **不给投资建议**：投资判断由用户定夺
- 产出公司简报格式 Word（Letter / 仿宋 10.5pt / 深蓝表头 #1A1A2E / PAGE 域页脚）

## 安装

### WorkBuddy 用户（推荐）

```bash
# 用户级（所有项目可用）
git clone https://github.com/ZhouxSong/yuexiang-wubpjianbao.git \
    ~/.workbuddy/skills/yuexiang-wubpjianbao1.0

# 或项目级
git clone https://github.com/ZhouxSong/yuexiang-wubpjianbao.git \
    <你的工作区>/.workbuddy/skills/yuexiang-wubpjianbao1.0
```

安装后重启/刷新会话，给出一个公司名称说"查一下这家公司出个简报"即可触发。

### 其他兼容 Agent（Claude Code 等）

遵循通用 `SKILL.md` 规范，放入对应 Agent 的 skills 目录即可。

## 依赖

**核心脚本依赖**（生成 Word 简报）：

| 依赖 | 用途 | 使用它的脚本 |
|---|---|---|
| `python-docx` | 公司简报格式排版（标题块/深蓝表头/页眉页脚 PAGE 域） | `scripts/brief_docx_template.py` |

```bash
pip install -r requirements.txt
```

**外部数据服务**（管线推荐接入，缺失时自动降级为公开页检索并显式标注"数据源受限"）：
企查查 / 天眼查（工商、股权、司法、专利双源交叉）、券商研报知识库、金融数据查询。
技能本体**不需要任何 API Key**；未接入外部连接器时仍可运行（用通用搜索检索
爱企查/天眼查公开页替代，不得编造数据）。

## 快速上手

| 阶段 | 动作 | 规范文件 |
|---|---|---|
| P0 | 锚定信用代码 + 同名辨析 | `SKILL.md` P0 节 |
| P1 | 七类检索（含关键字模板与双源规则） | `sources-matrix.md` |
| P2 | 证据分级 | `SKILL.md` P2 节 |
| P3 | 五章骨架撰写（开篇声明/表格 caption 强制） | `report-structure.md` |
| P4 | 10 项交付自检 | `delivery-checklist.md` |
| P5 | 归档命名与记忆 | `SKILL.md` P5 节 |

## 目录结构

```
SKILL.md                              技能主文件（六阶段工作流 + 铁律 + Word 规范 + 常见坑）
requirements.txt                      Python 依赖（python-docx）
scripts/
  brief_docx_template.py              公司简报 Word 参数化模板（复制后填数据区即可）
references/
  report-structure.md                 五章结构规范（章节/表格清单/caption 规范，强制）
  sources-matrix.md                   七类数据源检索矩阵（关键字模板 + 双源规则 + 降级铁律）
  delivery-checklist.md               10 项交付自检清单 + 去AI味写作规则
```

## 能力边界

✅ 无 BP 公司公开信息建档、预筛简报、同名主体辨析、证据分级、尽调核实清单
❌ 有 BP 的项目判断/深度尽调（配合同作者 `yuexiang-xiangmupanduan` 技能）、
行业赛道级深度研究报告（配合同作者 `yuexiang-hangyeyanjiu` 技能）

## 隐私说明

发布前已完成敏感信息扫描与清洗（自动扫描 + 人工全文复核）：
无本机路径、用户名、密钥、私人邮箱；真实尽调案例的公司名/人名/境外主体名已泛化。
若发现遗漏，请提 Issue 或直接提交 PR。

## License

[MIT](LICENSE)
