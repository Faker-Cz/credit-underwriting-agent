# Credit Underwriting Agent

面向对公客户的授信调查报告 Agent + Skill。项目围绕真实材料取数、财务分析、流动资金需求测算、担保与融资分析、Word 文档填充和交付质量检查构建。

首次使用、交接新项目或需要了解完整操作顺序时，请先阅读[授信报告智能体使用说明与操作指引](授信报告智能体使用说明与操作指引.md)。该文件覆盖材料访谈、上市／非上市分流、项目提示词、模板和版本管理、财务及保证人分析、流贷测算、Word 格式、质量门禁和交付清理。

## 项目结构

```text
skill/corporate-credit-investigation-cn/
├── SKILL.md                  # 主流程、证据边界和交付规则
├── agents/openai.yaml        # Codex Agent 展示及默认调用配置
├── references/               # 财务、授信方案、担保、写作与质检规则
├── scripts/                  # 材料、Word、Excel 审计脚本
└── assets/                   # 已确认的标准空白Word模板及无客户数据的流贷测算模板
scripts/
└── generate_credit_report_template.py
```

## 核心能力

- 盘点审计报告、三张财务报表、科目明细、征信、工商、上下游和担保材料；
- 建立来源台账并校验报表内、跨表和跨年勾稽；
- 通过财务分析专门代理协议完成偿债、营运、盈利、成长和现金流分析；
- 编制授信方案、非财务、融资及或有负债、担保和风险管控分析；
- 区分低风险、一般授信、高风险／复杂项目及上市公司等报告模式，按审批需要调整分析深度；
- 以客户经理口吻形成“事实与数据—业务原因—偿债或授信意义—风险边界”的完整逻辑链；
- 在用户授权的本地 Word 模板中定点补充报告内容；
- 只补充或替换既有填写位置，不改动全文排版、格式、字体、字号、表格结构、页眉页脚或分页设置；
- 对最终 DOCX 和流动资金测算工作簿执行结构化审计。

## 安装 Skill

将 Skill 目录复制到个人 Codex Skills 目录：

```bash
cp -R skill/corporate-credit-investigation-cn ~/.codex/skills/
```

重新启动或刷新 Codex 后，可以通过 `$corporate-credit-investigation-cn` 显式调用。

## 模板来源

正式报告只使用 Skill 内两份经确认的标准 Word 模板：

```text
skill/corporate-credit-investigation-cn/assets/单一客户授信调查报告模版.docx
skill/corporate-credit-investigation-cn/assets/低风险报告模版.docx
```

一般单一客户报告使用第一份，低风险／无风险报告使用第二份。两份模板只读复制后另存项目文件，不得混用、覆盖或用生成脚本替代。

仓库根目录同时保留两份同名模板，便于人工查看和替换；提交前应核对根目录文件与`skill/corporate-credit-investigation-cn/assets/`运行资产的哈希一致，防止出现两套版本。

## 模板生成工具（非正式模板来源）

安装依赖：

```bash
python -m pip install -r requirements.txt
```

下列脚本仅用于技术测试或生成通用空白文档，不作为正式授信报告模板来源：

```bash
python scripts/generate_credit_report_template.py --output <local-docx-path>
```

复选框应在 Microsoft Word 桌面版中点击；PDF 和部分第三方预览工具只显示当前状态。

## 私有材料与流贷测算表

本仓库包含上述两份空白 Word 标准模板和一份无客户数据的流动资金测算清洁模板，不包含客户审计报告、征信、身份证、工商账户信息、已填财务底稿或其他业务数据。

空白流动资金测算模板放在：

```text
skill/corporate-credit-investigation-cn/assets/流动资金类贷款授信额度测算表.xlsx
```

该模板为单工作表清洁版，包含调节（保险）系数、特定需求1/2及最终新增额度公式；空白输入时不显示除零错误。项目已有用户手工更新版时，项目文件优先，不用此通用模板覆盖。
仓库根目录同时保留同名工作簿便于人工查看；提交前须与`assets/`运行资产核对哈希一致。

不得向公开仓库提交任何已填客户数据、客户材料或访问凭据。

## 财务分析实现说明

当前项目未引入第三方 GitHub 财务分析模型。专业分析依赖：

- 审计报告及附注、财务报表和科目明细；
- 明确的报表勾稽与财务指标公式；
- 重要科目变化门槛；
- 财务分析专门代理任务书；
- 主 Agent 的来源抽查和独立第二遍复核；
- 面向机构授信场景的融资、担保和流动资金测算规则。

财务结论不能替代正式审批、审计意见或持牌专业人员的最终判断。
