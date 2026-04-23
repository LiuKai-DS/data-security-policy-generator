# data-security-policy-generator

> 📋 基于《以人为本的数据安全运营方法》的 WorkBuddy Skill

三阶段数据安全策略设计工具，帮助 IT 安全人员快速生成可落地的 DLP/UEBA 策略配置方案。

---

## 🔑 核心能力

- **第一阶段**：行为管控策略 — 331 条通用行为策略（R/L/A/C/X/S 六类）
- **第二阶段**：业务数据策略 — 13 个部门的业务关键字提取（制药/半导体/食品/能源/金融/地产/现代服务等 BU 类型）
- **第三阶段**：融合策略 — 行为 × 内容双维度融合策略，含 200 条融合规则
- **评分体系**：支持激进/平衡/保守三种风险偏好，自动计算策略分值与执行动作

---

## 🚀 快速开始

### 安装

1. 下载 `data-security-policy-generator-v2.0-release.zip`
2. 解压到 WorkBuddy Skills 目录：

```bash
unzip data-security-policy-generator-v2.0-release.zip \
  -d ~/.workbuddy/skills/data-security-policy-generator/
```

### 使用

在 WorkBuddy 中说出以下触发词即可：

```
"帮我设计研发部的数据安全防护策略"
"提取财务部门的敏感数据关键字"
"按保守风险偏好设计审批策略"
"设计非工作时间+USB+核心商密的融合策略"
```

---

## 📁 文件结构

```
data-security-policy-generator/
├── SKILL.md                              # 技能主文件（AI 读取的指令）
├── references/
│   ├── 01_strategy_catalog.md           # 策略目录示例（预览）
│   ├── 02_department_mapping.md          # 部门×策略适用性矩阵
│   ├── 03_legal_framework.md             # 数据安全法规手册
│   ├── 04_bu_keywords.md                # BU关键字字典示例（预览）
│   ├── 05_keyword_taxonomy.md           # 扩充字典示例（预览）
│   └── api_reference.md                  # DLP/UEBA API 对接参考
├── SOP/
│   ├── 01_mvp_strategy.md               # MVP 策略设计流程
│   ├── 02_business_data_extract.md      # 业务数据关键字提取
│   ├── 03_fusion_strategy.md            # 融合策略设计
│   └── 04_scoring_system.md             # 策略评分体系
└── scripts/
    └── __pycache__/
        └── strategy_lib.cpython-313.pyc  # 🔒 加密策略库（431条策略 + 完整关键字）
```

---

## ⚙️ 工作流程

```
用户需求 → 风险偏好确认 → 三阶段策略设计 → 评分计算 → 输出策略表
```

1. **确认风险偏好**：激进 / 平衡 / 保守
2. **第一阶段**：行为驱动策略（异常行为触发）
3. **第二阶段**：业务数据策略（敏感内容检测）
4. **第三阶段**：行为 × 内容融合（双维度精准拦截）

---

## 📜 法规依据

策略说明引用以下法规条款：

- 《网络安全法》
- 《数据安全法》
- 《个人信息保护法》
- 《网络数据安全管理条例》
- 《关键信息基础设施安全保护条例》

---

## ⚠️ 使用须知

- 本技能仅作为**辅助工具**使用，帮助 IT 人员理解和设计策略框架
- 如需精准化、定制化的数据安全管控方案，建议咨询专业安全团队
- 用户上传的数据和信息仅用于本次分析，不在本地存储

---

## 📄 License

MIT License — 可自由使用、修改和分发
