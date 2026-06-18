---
description: 将代码推送到 GAS 并更新指定部署，发布 Web App 新版本
---

# deploy 工作流

按以下步骤将本地代码部署为 GAS Web App 新版本。通过更新**现有部署 ID** 来保持 Web App URL 不变。

1. 查看现有部署列表，确认要更新的部署 ID（`AKfycb...` 开头）。
// turbo
```powershell
clasp deployments
```

2. 更新指定部署，发布新版本。将 `<DEPLOYMENT_ID>` 替换为步骤 1 中要更新的部署 ID，`<描述>` 用简短中文概括本次部署内容。
```powershell
clasp deploy --deploymentId <DEPLOYMENT_ID> --description "<描述>"
```

## 注意事项

- **更新现有部署**（指定 `--deploymentId`）会保持 Web App URL 不变，适合线上发布。
- 若不指定 `--deploymentId`，`clasp deploy` 会创建全新部署并生成新 URL，一般不要这样做。
- 常见部署参考（执行 `clasp deployments` 获取最新列表）：
  - `@HEAD`：开发/测试部署，始终跟随最新代码，无需 deploy。
  - `default meta-version`：默认主部署，通常为线上正式 URL。
- 部署前确保已执行 `clasp push`，否则发布的是旧代码。
