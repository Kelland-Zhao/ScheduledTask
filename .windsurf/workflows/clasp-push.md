---
description: 查看本地与 GAS 的差异并将代码推送到 Google Apps Script
---

# clasp-push 工作流

按以下步骤将本地代码推送到 Google Apps Script (GAS)。

1. 查看本地与 GAS 远端的文件差异，确认要推送的内容。
// turbo
```powershell
clasp status
```

2. 推送代码到 GAS。
```powershell
clasp push
```

## 注意事项

- `.clasp.json` 和 `appsscript.json` 为本地配置，已在 `.gitignore` 中排除，不影响推送。
- 推送会以本地文件覆盖 GAS 端对应文件，请确认本地为最新版本。
- 如遇文件冲突或需强制覆盖，可使用 `clasp push --force`。
- 推送完成后如需更新 Web App 部署版本，请另行执行 `clasp deploy`。
