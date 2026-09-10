# AGENTS Skills 仓库

> 所有技能的**真相源**。Claude、Kimi 和其他兼容 AGENTS 协议的工具通过此目录消费技能。

## 目录关系

```
D:\data\
├── .agents\skills\        ← 本仓库，所有技能的真实源代码
│                             所有修改必须在此提交
│
└── .claude\skills\        ← Claude 运行时加载目录
    ├── symlink × N       → 指向 ../.agents/skills/<name>
    └── 本地独有目录        → 待迁移到此处
```

## 维护规则

1. **改技能内容** → 在 `.agents/skills/<name>/` 里改，然后 `git commit`
2. **新增技能** → 放在 `.agents/skills/`，在 `.claude/skills/` 里创建 symlink
3. **不要直接编辑 `.claude/skills/` 下的 symlink 目标**（编辑等同于改源码，但容易忘记 commit）
4. **敏感数据**（API Key、密码）放在 `.env`，已在 `.gitignore` 中排除

## 当前技能（24+）

详见各子目录的 `SKILL.md`。

## 创建 symlink（Windows / Git Bash）

```bash
ln -s /d/data/.agents/skills/<name> /d/data/.claude/skills/<name>
```

需要开启 Windows 开发者模式或以管理员身份运行。

> **许可说明**：本仓库原创/整合内容采用 [MIT](LICENSE)；聚合的第三方技能沿用各自许可，详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
