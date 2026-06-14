# 可执行测试代码

为保持现有工具链的标准发现机制，测试代码保留在业务模块旁：

| 路径 | 内容 |
|---|---|
| `../../backend/tests/` | 后端单元、API、集成、安全、WebSocket 和并发风险测试 |
| `../../frontend/tests/contracts.test.mjs` | 前端 API、路由守卫和认证契约测试 |
| `../../backend/pytest.ini` | pytest 与覆盖率配置 |
| `../../backend/requirements-test.txt` | 后端测试依赖 |
| `../../run_tests.ps1` | 全套测试、构建、日志和覆盖率入口 |

在仓库根目录执行：

```powershell
.\run_tests.ps1
```
