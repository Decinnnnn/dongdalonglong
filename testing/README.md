# 测试材料总入口

本目录集中保存课程要求对应关系、测试过程文档、正式提交件、执行证据和材料生成工具。

测试负责人：李艺涵（学号：2023302628，项目计划成员 C：测试/文档）。

## 目录说明

| 路径 | 内容 |
|---|---|
| `documents/` | 测试计划、需求追踪、缺陷记录、测试报告、日志索引 |
| `course-submission/` | 按课程模板整合的 DOCX、XLSX 和测试评审会议纪要 |
| `evidence/logs/` | 每次自动化测试与前端构建的原始日志 |
| `evidence/coverage/` | 后端 HTML 覆盖率报告 |
| `tools/` | 课程提交文档和表格的生成脚本 |
| `code/README.md` | 可执行测试代码位置与运行方式说明 |

可执行测试代码保留在 `backend/tests/` 和 `frontend/tests/`，这是 pytest 与 npm 的标准目录，移动后会破坏现有配置和测试发现。

## 课程要求与当前文件对应

| 课程/计划要求 | 对应文件或证据 | 当前状态 |
|---|---|---|
| 测试设计及结果报告书 | `course-submission/测试记录汇总.xlsx`、`documents/01-测试计划.md`、`documents/04-测试报告.md` | 已完成 |
| 测试项目编写与实际执行 | `backend/tests/`、`frontend/tests/`、`run_tests.ps1` | 已完成 |
| FT 测试项目记录 | `course-submission/测试记录汇总.xlsx` 的“测试项目一览” | 已完成，36 项 |
| 需求追踪与 10 个集成场景 | `documents/02-测试用例与需求追踪.md`、汇总表“需求追踪” | 已完成 |
| Bug 跟踪单 | `documents/03-缺陷记录.md`、汇总表”Bug一览” | **✅ 2 个缺陷均已关闭** |
| Review 记录 | 汇总表”Review记录” | 已形成，参与人确认与签字待补 |
| 测试评审会议纪要 | `course-submission/会议纪要-20260614-测试评审.docx` | 已形成，实际地点/时段/出席确认待补 |
| 测试日志与品质记录 | `evidence/logs/`、`evidence/coverage/`、`documents/05-测试执行日志索引.md` | **✅ 已完成（含最终回归）** |
| API、集成、WebSocket、安全、构建测试 | 自动化测试代码与日志 | **✅ 核心范围已完成** |
| 性能测试方案 | `documents/06-性能测试方案.md` | **✅ 方案已编制** |
| E2E 测试场景设计 | `documents/07-E2E测试场景设计.md` | **✅ 场景设计完成** |
| 100 用户并发下单、持续压力、WebSocket 延迟 | 汇总表 FT-034 至 FT-036 | 方案已编制，执行待排期 |
| MySQL 全量回归、真实浏览器 E2E | 测试报告”残余风险” | 场景已设计，执行待排期 |

更完整的逐项解释见 `course-submission/测试要求落实与材料对应说明.docx`。

## 当前结论

**✅ 有条件通过。** 核心阻断缺陷（DEF-001 P1 并发抢单、DEF-002 P0 bcrypt 兼容）均已修复并回归通过。

- 最终回归：后端 **34 passed、0 failed、0 xfailed、覆盖率 82%**；前端 4 passed；构建通过。
- DEF-001（并发抢单非原子）→ **已关闭**：原子条件 UPDATE 替代先读后写。
- DEF-002（bcrypt 依赖兼容）→ **已关闭**：锁定 `bcrypt==4.0.1`。
- `utils` 覆盖率 **91%**，达标（目标 90%）。
- 性能测试方案和 E2E 场景设计已完成，作为交付物提供。
- 100 用户性能、压力、WS 延迟、MySQL 回归、真实浏览器 E2E 作为残余风险，不阻断当前发布。

## 修复摘要（2026-06-15）

| 修复项 | 严重级别 | 修复方式 | 验证结果 |
|---|---|---|---|
| DEF-002 bcrypt 兼容 | P0 阻断 | `requirements.txt` 新增 `bcrypt==4.0.1` | 注册/登录全回归通过 |
| DEF-001 并发抢单 | P1 高危 | `courier.py` 原子条件 UPDATE + rowcount 检查 | 并发测试通过，仅一人成功 |
| utils 覆盖率 | 提升 | 新增 8 个单元测试 (parse_token_optional/边界密码/build_order_out) | 80% → 91% |

## 残余补充项

1. 执行 100 用户并发下单、持续压力和 WebSocket 推送延迟测试（方案已完成，见 `documents/06-性能测试方案.md`）。
2. 执行真实浏览器三角色 E2E 和 MySQL 全量回归（场景设计已完成，见 `documents/07-E2E测试场景设计.md`）。
3. 补齐其他成员学号、贡献说明、Review 确认人、会议实际出席和签字。

## 一键复现

```powershell
.\run_tests.ps1
```

执行后，新日志写入 `testing/evidence/logs/`，覆盖率报告刷新到 `testing/evidence/coverage/index.html`。
