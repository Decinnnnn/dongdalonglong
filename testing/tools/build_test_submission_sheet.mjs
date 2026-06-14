import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outDir = "E:/软工实验/dongdalonglong/testing/course-submission";
await fs.mkdir(outDir, { recursive: true });

const wb = Workbook.create();
const navy = "#1F4E78";
const lightBlue = "#D9EAF7";
const lightGray = "#F2F4F7";
const green = "#E2F0D9";
const red = "#FCE4D6";

function title(sheet, text, range) {
  sheet.mergeCells(range);
  const cell = sheet.getRange(range.split(":")[0]);
  cell.values = [[text]];
  cell.format = {
    fill: navy,
    font: { bold: true, color: "#FFFFFF", size: 18 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
}

function styleHeader(range) {
  range.format = {
    fill: navy,
    font: { bold: true, color: "#FFFFFF", size: 10 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { all: { color: "#B7C9D6", style: "thin" } },
  };
}

function styleBody(range) {
  range.format = {
    font: { size: 9, color: "#222222" },
    verticalAlignment: "center",
    wrapText: true,
    borders: { all: { color: "#D9E1F2", style: "thin" } },
  };
}

const summary = wb.worksheets.add("执行摘要");
title(summary, "全栈外卖管理系统 - 测试设计及结果报告书", "A1:F2");
summary.getRange("A4:B10").values = [
  ["项目", "内容"],
  ["被测版本", "98e3939（test 分支）"],
  ["执行日期", "2026-06-14"],
  ["测试结论", "有条件不通过，不建议直接发布"],
  ["后端自动化", "30 passed，1 xfailed，覆盖率 76%"],
  ["前端测试与构建", "4 passed；生产构建通过"],
  ["已知缺陷", "DEF-001（P1）、DEF-002（P0）"],
  ["测试负责人", "李艺涵（学号：2023302628，计划书成员 C）"],
];
styleHeader(summary.getRange("A4:B4"));
styleBody(summary.getRange("A5:B11"));
summary.getRange("A13:F13").merge();
summary.getRange("A13").values = [["课程与计划书完成判定"]];
styleHeader(summary.getRange("A13:F13"));
summary.getRange("A14:F21").values = [
  ["要求", "计划标准", "实际情况", "判定", "证据", "备注"],
  ["接口测试", "22 个端点，100%通过", "31 个后端测试，30通过、1预期失败", "部分满足", "后端日志", "并发缺陷未满足"],
  ["10 个集成场景", "核心跨角色场景", "已覆盖核心状态机、越权和事务场景", "满足", "test_orders.py", ""],
  ["WebSocket", "推送与连接", "完成鉴权连接测试，未测延迟", "部分满足", "WebSocket测试", ""],
  ["安全测试", "越权、Token过期", "完成无/非法Token和跨角色；未单测过期边界", "部分满足", "test_auth.py", ""],
  ["并发性能", "100用户下单<500ms", "仅双骑手并发风险测试", "未完成", "DEF-001", "需补性能测试"],
  ["压力测试", "持续高负载稳定", "未执行", "未完成", "-", "需补测"],
  ["品质记录", "报告、缺陷、日志归档", "已完成", "满足", "testing/", ""],
];
styleHeader(summary.getRange("A14:F14"));
styleBody(summary.getRange("A15:F21"));
summary.getRange("A4:A11").format.fill = lightBlue;
summary.getRange("D15:D21").format.horizontalAlignment = "center";
summary.getRange("A1:F30").format.rowHeight = 24;
summary.getRange("A1:F30").format.columnWidth = 18;
summary.getRange("B1:B30").format.columnWidth = 34;
summary.getRange("C1:C30").format.columnWidth = 38;
summary.getRange("E1:E30").format.columnWidth = 24;
summary.freezePanes.freezeRows(3);

const cases = [
  ["FT-001","认证","健康检查","GET / 返回服务状态","200 且版本正确","OK","-"],
  ["FT-002","认证","三角色注册登录","分别注册 customer/merchant/courier 并登录","均成功并返回 JWT","OK","-"],
  ["FT-003","认证","同角色重名","重复注册同名顾客","返回 409","OK","-"],
  ["FT-004","认证","跨角色同名","顾客与商家使用同名注册","均可成功","OK","-"],
  ["FT-005","认证","注册字段校验","非法角色、空名、短密码","返回 422","OK","-"],
  ["FT-006","认证","错误登录","未知用户和错误密码登录","返回业务错误","OK","-"],
  ["FT-007","安全","未认证访问","不带 Token 访问地址接口","拒绝访问","OK","-"],
  ["FT-008","安全","RBAC 越权","商家 Token 请求顾客地址","返回 403","OK","-"],
  ["FT-009","安全","非法 Token","伪造 Token 请求受保护接口","返回 401","OK","-"],
  ["FT-010","菜品","顾客菜品过滤","下架/删除后查询菜品","不可见","OK","-"],
  ["FT-011","菜品","价格与状态校验","负数价格、非法上下架状态","返回 422","OK","-"],
  ["FT-012","菜品","菜品所有权","商家修改他人菜品","返回 404","OK","-"],
  ["FT-013","地址","地址软删除与所有权","他人删除、本人删除、重复删除","符合权限和软删除规则","OK","-"],
  ["FT-014","商家","有效商家列表","仅有在售菜品商家显示","列表正确","OK","-"],
  ["FT-015","订单","下单金额与历史","套餐18.50×2下单并查历史","总额37.00且明细正确","OK","-"],
  ["FT-016","订单","删除地址回滚","使用已删除地址下单","失败且无僵尸订单","OK","-"],
  ["FT-017","订单","不可用/跨商家菜品回滚","下架菜品或跨商家菜品下单","失败且无僵尸订单","OK","-"],
  ["FT-018","订单","他人地址越权","使用其他顾客地址下单","返回 403","OK","-"],
  ["FT-019","订单","取消订单状态机","他人取消、本人取消、重复取消","仅本人待处理订单可取消","OK","-"],
  ["FT-020","集成","完整配送闭环","下单→接单→抢单→送达","状态0→1→2→3且骑手释放","OK","-"],
  ["FT-021","订单","拒单后状态约束","拒单后再接单/抢单","均失败","OK","-"],
  ["FT-022","订单","商家指派空闲骑手","接单后指派骑手","订单派送中且骑手可查","OK","-"],
  ["FT-023","并发","并发抢单唯一性","两个骑手同时抢一个订单","仅一个成功","NG","DEF-001"],
  ["FT-024","单元","密码哈希与 JWT","哈希验证、JWT往返","结果正确","OK","-"],
  ["FT-025","单元","订单状态文本","核对0-5状态映射","全部正确","OK","-"],
  ["FT-026","WebSocket","合法身份连接","正确角色、用户、Token连接","连接成功","OK","-"],
  ["FT-027","WebSocket","非法身份连接","缺失Token或身份不匹配","以4001/4003关闭","OK","-"],
  ["FT-028","前端","API 服务契约","检查所有业务 API 封装","全部存在","OK","-"],
  ["FT-029","前端","路由守卫","检查三角色路由和鉴权判断","全部存在","OK","-"],
  ["FT-030","前端","认证状态存储","检查 set/remove 全部认证字段","全部存在","OK","-"],
  ["FT-031","前端","401 处理","检查 Bearer Token 和过期退出","契约正确","OK","-"],
  ["FT-032","构建","生产构建","执行 npm run build","构建成功","OK","-"],
  ["FT-033","环境","原始依赖干净安装认证","安装最新 bcrypt 后注册","注册登录可用","NG","DEF-002"],
  ["FT-034","性能","100用户并发下单","并发请求并统计响应","<500ms","未执行","-"],
  ["FT-035","压力","持续高负载稳定性","持续请求并监控内存","无崩溃/泄漏","未执行","-"],
  ["FT-036","WebSocket","推送延迟","下单并计时到商家收到消息","<1秒","未执行","-"],
];

const testSheet = wb.worksheets.add("测试项目一览");
title(testSheet, "功能测试设计书 / 报告书 - 测试项目一览", "A1:G2");
testSheet.getRange("A4:G4").values = [["项号","大分类项目","确认事项","测试条件/操作内容","预期结果","实施结果","故障序号"]];
styleHeader(testSheet.getRange("A4:G4"));
testSheet.getRange(`A5:G${cases.length + 4}`).values = cases;
styleBody(testSheet.getRange(`A5:G${cases.length + 4}`));
testSheet.getRange("F5:F40").format.horizontalAlignment = "center";
testSheet.getRange("A5:A40").format.horizontalAlignment = "center";
testSheet.getRange("G5:G40").format.horizontalAlignment = "center";
testSheet.getRange("A1:G45").format.rowHeight = 30;
for (const [col, width] of [["A:A",12],["B:B",14],["C:C",23],["D:D",42],["E:E",35],["F:F",12],["G:G",14]]) {
  testSheet.getRange(col).format.columnWidth = width;
}
testSheet.getRange("F27").format.fill = red;
testSheet.getRange("F37").format.fill = red;
testSheet.getRange("F38:F40").format.fill = lightGray;
testSheet.freezePanes.freezeRows(4);

const bugs = wb.worksheets.add("Bug一览");
title(bugs, "Bug 一览表", "A1:J2");
bugs.getRange("A4:J4").values = [["序号","状态","重要度","发生模块","问题点","复现步骤","对应测试项目","Bug原因","修改处理/建议","回归确认"]];
styleHeader(bugs.getRange("A4:J4"));
bugs.getRange("A5:J6").values = [
  ["DEF-001","待修复","P1/高","骑手抢单","并发抢单缺少原子条件更新，可能双成功并造成骑手状态不一致","创建已接单订单；两个空闲骑手并发调用 pickup","FT-023","读后写流程无条件更新、无行锁","使用 UPDATE ... WHERE status=1 AND courier_id IS NULL 并检查影响行数；MySQL回归","未通过"],
  ["DEF-002","待修复","P0/阻断","认证/依赖","Passlib 1.7.4 与 bcrypt 5.0.0 不兼容，干净安装后注册登录异常","按原 requirements 安装；调用注册接口","FT-033","生产依赖未锁定兼容 bcrypt 版本","在 requirements.txt 固定 bcrypt==4.0.1，并从干净环境全量回归","未通过"],
];
styleBody(bugs.getRange("A5:J6"));
bugs.getRange("A5:J6").format.fill = red;
bugs.getRange("A1:J15").format.rowHeight = 42;
for (const [col, width] of [["A:A",12],["B:B",12],["C:C",12],["D:D",18],["E:E",38],["F:F",34],["G:G",16],["H:H",34],["I:I",42],["J:J",14]]) {
  bugs.getRange(col).format.columnWidth = width;
}
bugs.freezePanes.freezeRows(4);

const review = wb.worksheets.add("Review记录");
title(review, "测试阶段 Review 记录表", "A1:H2");
review.getRange("A4:B11").values = [
  ["项目","内容"],
  ["项目名","全栈外卖管理系统"],
  ["Review 工程","测试设计与结果评审"],
  ["日期","2026-06-14"],
  ["记录者","李艺涵（学号：2023302628，测试负责人）"],
  ["参与者","贾孟祺、丁雨晨、李艺涵、王烜铱（出席情况待确认）"],
  ["Check List 项数","8"],
  ["Review 总问题数","4"],
];
styleHeader(review.getRange("A4:B4"));
styleBody(review.getRange("A5:B11"));
review.getRange("A13:H13").values = [["序号","评审项目","指出事项","BUG判定","程度","处理内容","责任角色","完了确认"]];
styleHeader(review.getRange("A13:H13"));
review.getRange("A14:H17").values = [
  ["1","依赖与环境","原始依赖可安装但最新 bcrypt 会阻断认证","Yes","高","锁定兼容版本并做干净环境回归","后端开发","待处理"],
  ["2","并发测试","并发抢单不能保证唯一成功","Yes","高","实现原子条件更新并在 MySQL 回归","后端开发","待处理"],
  ["3","非功能测试","计划要求的100用户性能、压力、WebSocket延迟尚未执行","No","中","补充性能与延迟测试报告","测试人员","待处理"],
  ["4","课程提交信息","其他成员学号、贡献度和签字尚待补全","No","低","李艺涵信息已确认；提交前补齐其余信息","项目组长","待处理"],
];
styleBody(review.getRange("A14:H17"));
review.getRange("A14:H17").format.fill = lightGray;
review.getRange("A1:H25").format.rowHeight = 34;
for (const [col, width] of [["A:A",12],["B:B",22],["C:C",42],["D:D",12],["E:E",10],["F:F",38],["G:G",18],["H:H",14]]) {
  review.getRange(col).format.columnWidth = width;
}
review.freezePanes.freezeRows(13);

const trace = wb.worksheets.add("需求追踪");
title(trace, "测试需求追踪矩阵", "A1:E2");
const traces = [
  ["R-AUTH-01","三角色注册登录与JWT","FT-002～FT-006、FT-024、FT-033","部分满足","DEF-002待修复"],
  ["R-AUTH-02","RBAC权限隔离","FT-007～FT-009","满足",""],
  ["R-CUS-01","浏览有效商家和菜品","FT-010、FT-014","满足",""],
  ["R-CUS-02","地址管理与软删除","FT-013","满足",""],
  ["R-CUS-03","原子下单与金额计算","FT-015～FT-018","满足",""],
  ["R-CUS-04","顾客取消待处理订单","FT-019","满足",""],
  ["R-MER-01","商家菜品管理与所有权","FT-011～FT-012","满足",""],
  ["R-MER-02","接单、拒单、指派骑手","FT-020～FT-022","满足",""],
  ["R-COU-01","抢单和确认送达","FT-020、FT-023","部分满足","DEF-001待修复"],
  ["R-DATA-01","异常下单无僵尸订单","FT-016～FT-018","满足",""],
  ["R-DATA-02","并发抢单唯一性","FT-023","不满足","DEF-001"],
  ["R-WS-01","WebSocket身份鉴权","FT-026～FT-027","满足","未量测延迟"],
  ["R-FE-01","前端路由/API/认证契约","FT-028～FT-031","满足",""],
  ["R-FE-02","前端生产构建","FT-032","满足","存在大包警告"],
  ["R-NFR-01","100用户并发下单<500ms","FT-034","未完成","需补测"],
  ["R-NFR-02","持续压力稳定性","FT-035","未完成","需补测"],
  ["R-NFR-03","WebSocket推送延迟<1秒","FT-036","未完成","需补测"],
];
trace.getRange("A4:E4").values = [["需求编号","需求","对应测试项目","状态","备注"]];
styleHeader(trace.getRange("A4:E4"));
trace.getRange(`A5:E${traces.length + 4}`).values = traces;
styleBody(trace.getRange(`A5:E${traces.length + 4}`));
trace.getRange("A1:E30").format.rowHeight = 30;
for (const [col, width] of [["A:A",16],["B:B",34],["C:C",28],["D:D",14],["E:E",28]]) {
  trace.getRange(col).format.columnWidth = width;
}
trace.freezePanes.freezeRows(4);

const files = wb.worksheets.add("材料对应");
title(files, "课程要求与提交材料对应表", "A1:D2");
const mappings = [
  ["测试设计及结果报告书","测试记录汇总.xlsx","FT设计、结果、Bug、Review、追踪","已完成"],
  ["详细测试计划与结果","testing/documents/01-测试计划.md、04-测试报告.md","范围、策略、环境、结论","已完成"],
  ["测试项目编写与执行","backend/tests/、frontend/tests/","个人测试贡献与可重复执行证据","已完成"],
  ["缺陷跟踪单","03-缺陷记录.md、Bug一览","缺陷复现、原因、建议、状态","已完成"],
  ["测试用例评审","Review记录、测试评审会议纪要","评审指出事项与处理决定","已完成"],
  ["品质记录","testing/evidence/logs、coverage","原始日志与覆盖率","已完成"],
  ["小组总结报告测试方法页","测试要求落实与材料对应说明.docx","总结PPT内容来源","可引用"],
];
files.getRange("A4:D4").values = [["课程/计划要求","对应成果物","用途","状态"]];
styleHeader(files.getRange("A4:D4"));
files.getRange(`A5:D${mappings.length + 4}`).values = mappings;
styleBody(files.getRange(`A5:D${mappings.length + 4}`));
files.getRange("A1:D20").format.rowHeight = 34;
for (const [col, width] of [["A:A",28],["B:B",42],["C:C",42],["D:D",14]]) {
  files.getRange(col).format.columnWidth = width;
}

const output = await SpreadsheetFile.exportXlsx(wb);
await output.save(`${outDir}/测试记录汇总.xlsx`);
console.log(`${outDir}/测试记录汇总.xlsx`);
