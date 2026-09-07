# Credit Underwriting Agent

面向对公客户的授信调查报告 Agent + Skill。项目围绕真实材料取数、财务分析、流动资金需求测算、担保与融资分析、Word 文档填充和交付质量检查构建。

## 项目结构

```text
skill/corporate-credit-investigation-cn/
├── SKILL.md                  # 主流程、证据边界和交付规则
├── agents/openai.yaml        # Codex Agent 展示及默认调用配置
├── references/               # 财务、授信方案、担保、写作与质检规则
├── scripts/                  # 材料、Word、Excel 审计脚本
└── assets/                   # 使用者本地放置的私有模板和测算表（不入库）
scripts/
└── generate_credit_report_template.py
```

## 核心能力

- 盘点审计报告、三张财务报表、科目明细、征信、工商、上下游和担保材料；
- 建立来源台账并校验报表内、跨表和跨年勾稽；
- 通过财务分析专门代理协议完成偿债、营运、盈利、成长和现金流分析；
- 编制授信方案、非财务、融资及或有负债、担保和风险管控分析；
- 在用户授权的本地 Word 模板中定点补充报告内容；
- 只补充或替换既有填写位置，不改动全文排版、格式、字体、字号、表格结构、页眉页脚或分页设置；
- 对最终 DOCX 和流动资金测算工作簿执行结构化审计。

## 安装 Skill

将 Skill 目录复制到个人 Codex Skills 目录：

```bash
cp -R skill/corporate-credit-investigation-cn ~/.codex/skills/
```

重新启动或刷新 Codex 后，可以通过 `$corporate-credit-investigation-cn` 显式调用。

## 生成标准 Word 模板

安装依赖：

```bash
python -m pip install -r requirements.txt
```

生成模板时由使用者自行指定本地输出文件：

```bash
python scripts/generate_credit_report_template.py --output <local-docx-path>
```

复选框应在 Microsoft Word 桌面版中点击；PDF 和部分第三方预览工具只显示当前状态。

## 私有材料与流贷测算表

本仓库不包含机构内部 Word 模板、客户审计报告、征信、身份证、工商账户信息、已填财务底稿或其他业务数据。

如使用机构内部正式 Word 模板，应在获得授权后仅放入本地：

```text
skill/corporate-credit-investigation-cn/assets/授信报告模版-skill.docx
```

该文件作为本地默认授信报告模板使用，不再沿用旧2018版模板，并继续受`.gitignore`保护，不提交到公开仓库。

流动资金测算工作簿的样例文件可能含实际数值，因此未提交。使用者应将内部批准使用的工作簿放在：

```text
skill/corporate-credit-investigation-cn/assets/working-capital-calculation-template.xlsx
```

不得向公开仓库提交任何客户数据或访问凭据。

## 财务分析实现说明

当前项目未引入第三方 GitHub 财务分析模型。专业分析依赖：

- 审计报告及附注、财务报表和科目明细；
- 明确的报表勾稽与财务指标公式；
- 重要科目变化门槛；
- 财务分析专门代理任务书；
- 主 Agent 的来源抽查和独立第二遍复核；
- 面向机构授信场景的融资、担保和流动资金测算规则。

财务结论不能替代正式审批、审计意见或持牌专业人员的最终判断。
