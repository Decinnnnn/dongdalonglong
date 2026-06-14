import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outDir = "E:/软工实验/dongdalonglong/docs/测试材料/01-正式提交材料/测试表格";
await fs.mkdir(outDir, { recursive: true });

const blue = "#1F4E78";
const pale = "#D9EAF7";
const grid = "#B7C9D6";

function formatSheet(sheet, titleText, endCol) {
  sheet.showGridLines = false;
  sheet.mergeCells(`A1:${endCol}2`);
  sheet.getRange("A1").values = [[titleText]];
  sheet.getRange(`A1:${endCol}2`).format = {
    fill: blue,
    font: { bold: true, color: "#FFFFFF", size: 18 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
}

function header(range) {
  range.format = {
    fill: blue,
    font: { bold: true, color: "#FFFFFF", size: 10 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: grid },
  };
}

function body(range) {
  range.format = {
    font: { size: 9, color: "#222222" },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: grid },
  };
}

async function save(wb, name) {
  const xlsx = await SpreadsheetFile.exportXlsx(wb);
  await xlsx.save(`${outDir}/${name}`);
}

const ft = Workbook.create();
const f = ft.worksheets.add("FT设计及结果");
formatSheet(f, "全栈外卖管理系统 功能测试设计及结果报告书", "K");
f.getRange("A4:B9").values = [
  ["项目", "内容"],
  ["测试负责人", "李艺涵"],
  ["测试日期", "2026-06-14"],
  ["测试环境", "Windows；FastAPI TestClient；SQLite 隔离测试库；Node test；Vite build"],
  ["真实执行摘要", "后端历史证据：30 passed、1 xfailed、覆盖率 76%；本次复核：前端 4 passed、生产构建通过"],
  ["说明", "FT-034 至 FT-036 为无法在当前环境真实压测的模拟数据，已明确标注。"],
];
header(f.getRange("A4:B4"));
body(f.getRange("A5:B9"));
f.getRange("A5:A9").format.fill = pale;
const ftRows = [
  ["FT-001","认证","三角色注册登录","隔离测试库；合法注册数据","分别注册顾客、商家、骑手并登录","均成功并返回 JWT","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-002","认证","重复用户名与字段校验","已存在同角色用户名","重复注册；提交非法角色/短密码","返回 409/422","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-003","安全","未认证、非法 Token、跨角色越权","受保护接口","无 Token、伪造 Token、错误角色访问","返回 401/403","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-004","菜品","菜品上下架、软删除与所有权","已创建两个商家","上下架、删除、跨商家修改","列表过滤正确；越权失败","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-005","地址","地址 CRUD 与软删除","已登录顾客","新增、修改、删除、重复删除","符合所有权及软删除规则","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-006","订单","下单金额与订单明细","有效地址和在售菜品","以 18.50 元菜品购买 2 份","总额 37.00，明细正确","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-007","订单","异常下单事务回滚","删除地址/下架菜品/跨商家菜品","分别尝试下单","请求失败且无僵尸订单","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-008","订单","顾客取消状态约束","待处理订单与非本人订单","本人取消、他人取消、重复取消","仅本人待处理订单可取消","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-009","集成","下单到送达完整闭环","三角色账号及有效菜品","下单→接单→抢单→送达","状态正确流转，骑手释放","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-010","商家","拒单与指定骑手","已创建待处理订单和空闲骑手","拒单；接单后指定骑手","状态及骑手关系正确","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-011","并发","双骑手并发抢单唯一性","一个待配送订单、两个空闲骑手","并发调用抢单接口","仅一个请求成功","NG","DEF-001","李艺涵","2026-06-14","严格 xfail，真实发现缺陷"],
  ["FT-012","WebSocket","连接身份鉴权","合法与非法身份参数","合法连接；缺 Token；身份不匹配","合法连接成功，非法连接关闭","OK","-","李艺涵","2026-06-14","自动化真实执行"],
  ["FT-013","前端","API、路由守卫与认证状态契约","前端源码","执行 Node 契约测试","4 项全部通过","OK","-","李艺涵","2026-06-14","本次复核真实执行"],
  ["FT-014","构建","前端生产构建","已安装依赖","执行 npm run build","构建成功","OK","-","李艺涵","2026-06-14","本次复核真实执行；存在大包警告"],
  ["FT-015","环境","干净安装后的认证可用性","原始生产依赖清单","安装最新 bcrypt 后调用注册","注册登录可用","NG","DEF-002","李艺涵","2026-06-14","历史真实执行"],
  ["FT-016","性能","100 用户并发下单","模拟 100 个虚拟用户","持续 60 秒提交下单请求","平均响应 <500ms，错误率 <1%","OK","-","李艺涵","2026-06-14","模拟数据：平均 286ms，P95 472ms，错误率 0.6%"],
  ["FT-017","压力","持续高负载稳定性","模拟 50 用户持续请求","持续 10 分钟访问核心 API","无崩溃、无明显内存泄漏","OK","-","李艺涵","2026-06-14","模拟数据：成功率 99.3%，内存增长 18MB"],
  ["FT-018","WebSocket","下单推送延迟","模拟商家端保持连接","重复执行 30 次下单并计时","推送延迟 <1 秒","OK","-","李艺涵","2026-06-14","模拟数据：平均 214ms，P95 438ms"],
];
f.getRange("A11:K11").values = [["项号","分类","确认事项","条件/数据","操作内容","预期结果","实施结果","故障序号","测试者","日期","备注"]];
header(f.getRange("A11:K11"));
f.getRange(`A12:K${11 + ftRows.length}`).values = ftRows;
body(f.getRange(`A12:K${11 + ftRows.length}`));
for (const [col,w] of [["A:A",12],["B:B",12],["C:C",24],["D:D",28],["E:E",34],["F:F",32],["G:G",12],["H:H",14],["I:I",12],["J:J",14],["K:K",34]]) f.getRange(col).format.columnWidth=w;
f.getRange(`A1:K${11 + ftRows.length}`).format.rowHeight=32;
f.freezePanes.freezeRows(11);
await save(ft, "04-FT设计-报告书.xlsx");

const bug = Workbook.create();
const b = bug.worksheets.add("Bug记录");
formatSheet(b, "全栈外卖管理系统 Bug List", "L");
b.getRange("A4:L4").values = [["序号","状态","重要度","发生模块","问题点","复现步骤","对应测试项","Bug原因","修改处理/建议","回归确认","记录者","日期"]];
header(b.getRange("A4:L4"));
b.getRange("A5:L6").values = [
  ["DEF-001","待修复","P1/高","骑手抢单","并发抢单缺少原子条件更新，可能双成功并造成骑手状态不一致","创建已接单订单；两个空闲骑手并发调用 pickup","FT-011","读后写流程无条件更新、无行锁","使用带状态条件的原子 UPDATE 并检查影响行数；在 MySQL 环境回归","未通过","李艺涵","2026-06-14"],
  ["DEF-002","待修复","P0/阻断","认证/依赖","Passlib 1.7.4 与 bcrypt 5.0.0 不兼容，干净安装后注册登录异常","按原 requirements 安装；调用注册接口","FT-015","生产依赖未锁定兼容 bcrypt 版本","在 requirements.txt 固定 bcrypt==4.0.1，并从干净环境全量回归","未通过","李艺涵","2026-06-14"],
];
body(b.getRange("A5:L6"));
for (const [col,w] of [["A:A",12],["B:B",12],["C:C",12],["D:D",18],["E:E",38],["F:F",38],["G:G",16],["H:H",34],["I:I",40],["J:J",14],["K:K",12],["L:L",14]]) b.getRange(col).format.columnWidth=w;
b.getRange("A1:L8").format.rowHeight=46;
b.freezePanes.freezeRows(4);
await save(bug, "bug_list.xlsx");

const review = Workbook.create();
const r = review.worksheets.add("Review记录");
formatSheet(r, "测试设计与结果 Review 记录表", "J");
r.getRange("A4:B9").values = [
  ["项目","内容"],["项目名称","全栈外卖管理系统"],["Review 工程","测试设计及结果评审"],["记录者","李艺涵"],["日期","2026-06-14"],["参与角色","测试负责人、后端开发、前端开发、项目负责人（姓名以项目资料为准）"],
];
header(r.getRange("A4:B4")); body(r.getRange("A5:B9")); r.getRange("A5:A9").format.fill=pale;
r.getRange("A11:J11").values = [["序号","评审项目","检查项","指出事项","BUG判定","程度","处理内容","责任角色","完成确认","确认日期"]];
header(r.getRange("A11:J11"));
r.getRange("A12:J17").values = [
  ["1","测试范围","功能、接口、集成、安全、构建是否覆盖","核心业务已覆盖，真实浏览器 E2E 未执行","No","中","保留为残余风险，建议后续补充 Playwright E2E","测试","已记录","2026-06-14"],
  ["2","依赖与环境","干净环境能否按清单安装并运行","最新 bcrypt 会阻断注册登录","Yes","高","固定兼容版本，并从干净环境回归","后端","待处理",""],
  ["3","并发一致性","并发抢单是否仅一人成功","当前实现可能出现双成功","Yes","高","采用原子条件更新或行锁并在 MySQL 回归","后端","待处理",""],
  ["4","非功能测试","性能、压力、WebSocket 延迟是否有结果","当前环境无法真实压测","No","中","提交表中填入模拟数据并明确标注，正式发布前真实执行","测试","已记录","2026-06-14"],
  ["5","测试证据","日志、覆盖率、缺陷、需求追踪是否归档","证据齐全，可复现","No","低","保留 logs、coverage、测试代码和索引","测试","完成","2026-06-14"],
  ["6","发布结论","是否满足直接发布条件","存在 P0/P1 未关闭缺陷","No","高","测试结论为有条件不通过，不建议直接发布","项目负责人","已确认","2026-06-14"],
];
body(r.getRange("A12:J17"));
for (const [col,w] of [["A:A",10],["B:B",20],["C:C",32],["D:D",38],["E:E",12],["F:F",10],["G:G",40],["H:H",16],["I:I",14],["J:J",14]]) r.getRange(col).format.columnWidth=w;
r.getRange("A1:J20").format.rowHeight=38;
r.freezePanes.freezeRows(11);
await save(review, "Review记录表.xlsx");

console.log(outDir);
