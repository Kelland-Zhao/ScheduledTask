---
description: 查看变更、暂存、按版本号规范提交并推送到远程仓库
---

# git-push 工作流

按以下步骤将本地变更提交并推送到远程仓库。版本号命名规范：`VYYYYMMDD.数字_改动描述`（如 `V20260616.01_去除小组Kaizen筛选视图`）。

1. 查看当前变更状态，确认要提交的文件。
// turbo
```powershell
git status
```

2. 查看具体改动内容（可选，用于确认本次提交范围）。
// turbo
```powershell
git diff --stat
```

3. 暂存所有变更。
// turbo
```powershell
git add -A
```

4. 提交变更。版本号格式为 `VYYYYMMDD.数字_改动描述`：
   - `YYYYMMDD` 为当天日期
   - `数字` 在当天前一个版本基础上递增（每天从 01 开始）
   - `改动描述` 用简短中文概括本次修改

```powershell
git commit -m "V20260616.01_改动描述"
```

5. 推送到远程仓库。
```powershell
git push
```

## 注意事项

- 提交前确认 `Google API/`、`.clasp.json`、`appsscript.json` 等敏感/本地配置文件未被纳入提交（已在 `.gitignore` 中排除）。
- 同一天多次提交时，版本号数字依次递增（`.01` → `.02` → `.03`）。
- 若需同步部署到 GAS，请在推送前手动执行 `clasp push`。
