# PSI 周更 Agent 需求文档

> 版本：2026-11-02（v2：补充 PSI 表结构分析、GTM 实习 SOP 里的节奏与下游）· 作者：胡佳琪（小米波兰 GTM）
> 用途：给 Claude Code / Codex 等 AI 编程 agent 读，作为开发「Poland PSI DB+Channel」每周自动更新工具的需求说明。
> 在线版（Claude Docs）：https://claude.ai/artifact/RiyoGUKn2NUtuqLrSHZrVp（v1，未同步 v2）
> 状态：需求梳理阶段，原始样本文件还没提供（放到 `samples/` 后再开发）。
> 参考：`GTM实习工作梳理.md` 第 2 章（PSI 表维护）、第 3 章（SO 预实分析）。

## 0. 术语

| 术语 | 含义 |
|---|---|
| PSI | Purchase（进货）/ Sales（销售）/ Inventory（库存），按周跟踪每个渠道每个机型的进、销、存 |
| SI | Sell in，小米卖给经销商（DB）的量。运营商没有 DB 层，所以 OP 的 ST 就算 SI |
| ST | Sell through，经销商（或小米直供）卖给 KA 的量，也就是 KA 的进货 |
| SO | Sell out，KA 卖给消费者的量 |
| Channel Inventory | 理论库存 = 上周库存 + ST − SO（公式） |
| Channel Inventory Actual | KA 周报里报的实际库存（手填） |
| Channel DOS | Days of Stock，渠道库存可卖天数 = Channel Inventory Actual ÷ 当周 SO × 7 |
| ST FCST / SO FCST | ST / SO 的周预测，由 KAM 填。未来的周里 ST / SO 行直接等于 FCST；过去的周用实际数覆盖 |
| KAM | Key Account Manager，负责某几个客户的销售（见 §2 联系人表） |
| KA | Key Account，重点客户（EURO、MEX、MSH、X-KOM、Play 等） |
| DB | Distributor，经销商：AB、ALSO、IM |
| OM / OP | Open Market 公开市场 / Operator 运营商 |
| 预实分析 | 预测 vs 实际：SO 实际和 SO FCST 的对比 |

---

## 1. 背景与目标

现在每周靠 Excel 手工更新 PSI 表（DB + 运营商 OP + 公开市场 OM），主要有三个问题：
- 各家数据格式不一样，产品命名和 PSI 表的叫法对不上；
- 对照关系散落在好几张工具表和"上周的文件"里；
- 步骤多，容易漏、容易错。

**目标**
1. 公邮来了周报，自动下载并归档（只要周报，只要手机和平板）。
2. 自动解析、清洗，把产品映射成 PSI 代码，按 PSI 的行顺序算出 ST / SO / Stock，生成**待粘贴表**。
3. 自动校验总数、找出没匹配上的型号和新品，输出一份核对报告。
4. 型号映射只维护一张统一的对照表，人工确认过一次的，以后自动命中。

**原则**：固定脚本负责解析、映射和计算（确定性，可复现）；AI 只处理例外，比如新命名、格式变了、异常数值。AI 不直接改 PSI 里的数字。

## 2. 范围与节奏

| 模块 | 更新到 PSI 哪里 | 数据来源 | 截止时间 |
|---|---|---|---|
| DB（经销商） | `2026 DB ST`、`2026 DB STOCK` 两张 sheet | AB、ALSO、IM 的邮件 | 每周一上午，完成后在 GTM 群说一声 |
| OP（四大运营商） | `Play` `Plus` `Orange` `TMO` 四张 sheet 的 ST / SO / Channel Inventory Actual 行 | ST 从国际 BI 查；SO 和库存从运营商周报取 | **周三下班前** |
| OM（公开市场） | `EURO` `MEX` `MSH` `X-KOM` `Michain` `OM Others` 等 sheet | ST 用公式从 DB ST 取；SO 和库存从各 KA 周报取 | 收到就更新，**周四 15:00 前** |
| 汇总 sheet | `ALL` `OM` `OP` | 全是公式，一般不动 | — |
| 下游：SO 预实分析 | 另一张飞书表「波兰公开市场 SO预实分析」，数据从 PSI 拉 | PSI | 周四 14:00–15:00 发 GTM 群，注明缺哪家数据 |

**为什么周四 15:00 是硬截止**：SO 预实分析直接引用 PSI，PSI 里没更新的 KA，预实分析就不完整。所以 agent 的核心 KPI 是"周四 14:00 前，哪些 KA 已就绪、哪些还缺"。

**缺数找谁**（agent 起草催数消息，人工确认后发；实在没收到可以后补，在预实分析里注明）：

| KA | 联系人 |
|---|---|
| ROLV / GG / EURO / Auchan | Dominika Wichowska |
| MSH / X-KOM | Agata Kozak Piotrowska |
| MEX / KTR / MAX | Maciej Świątek |
| TMO / Play | 盛俊 Sheng Jun |
| Plus / Orange | 吴江涛 Jet Wu |
| IM 库存 | 缪鹏飞 Davis Miao |
| 新品 SKU 信息查不到 | Pope Peng |

- 公邮：`pl-psi@support.mi.com`（Outlook / Microsoft 365）。
- 只要**周报**，不要日报；只要**手机（SMARTFONY / Telefony GSM）和平板（TABLETY / Tablety z systemem Android）**。手环、手表、耳机、电视、配件都丢掉。
- 周的定义：ISO 周，周一到周日（例：2026-W41 = 10-05 到 10-11）。⚠ 待确认。
- 指标含义：ST = sell through（经销商或渠道卖给 KA 的量）；SO = sell out（KA 卖给消费者的量）；Stock / Channel Inventory Actual = 库存。
- Michain = GG + ROLV；OM Others = MAX + Auchan + KTR（Komputronik）。

## 3. 总体架构

```
[公邮 Outlook] --Power Automate--> [OneDrive/本地 inbox/W{周}/{来源}/ 附件+正文]
                                            |
[国际 BI 妥投量] --浏览器导出(半自动)-->       v
[Jet Wu 的 Orange 表 / Davis Miao 的 IM 库存] -> [解析器: 每个来源一个 parser]
                                            |
                                            v
                              [统一映射表 mapping.xlsx] --未命中--> [AI 给候选 → 人工确认 → 回写映射表]
                                            |
                                            v
                              [按 PSI 行顺序汇总: ST / SO / Stock / Others]
                                            |
                                            v
                 [输出: 待粘贴表 output/W{周}_paste.xlsx + 核对报告 W{周}_report.md]
                                            |
                                            v
                         [人工粘贴进飞书 PSI 表 + 标颜色]（第二期再考虑 API 自动写入）
```

| 步骤 | 工具 |
|---|---|
| 收邮件 | Power Automate：收到指定发件人或标题的邮件，就把附件和正文存到 OneDrive，再同步到本地。备选：Microsoft Graph API（要 IT 授权） |
| 解析 | Python 3 + pandas + openpyxl；每个来源一个 parser，配置写在 `config/sources.yaml` |
| 映射 | `mapping/mapping.xlsx` + 规则匹配（EAN 精确匹配 → 物料码 / SKU 精确匹配 → 名称规则） |
| BI 查询 | 用浏览器自动化（Claude in Chrome），在已登录状态下按周和客户导出妥投量 |
| 输出与校验 | Python 生成 xlsx 和 md 报告 |
| 调度 | Claude 桌面端的定时任务或 cron：周一早上跑 DB，周三到周四跑 KA |

### 3.1 每周流水线（agent 的状态机）

| 时间 | agent 做什么 | 人做什么 |
|---|---|---|
| 周一早上 | 拉 AB/ALSO/IM 邮件 → 生成 DB ST / DB STOCK 待粘贴行 + 报告；起草"DB 已更新"群消息 | 粘贴、发群 |
| 周一–周三 | 每次有新周报进公邮就解析，更新"本周就绪看板"（每个 KA：已收到 / 已解析 / 有待确认映射 / 未收到） | 确认新映射 |
| 周三中午 | OP 四家还没就绪 → 起草催数消息给盛俊 / Jet Wu；提示去 BI 查四家 ST | 发消息、查 BI |
| 周三下班 | OP 待粘贴表定稿 | 粘贴 OP |
| 周四 10:00 | OM 还没就绪的 KA → 按联系人表起草催数消息 | 发消息 |
| 周四 14:00 | 生成 OM 待粘贴表定稿 + 核对报告 + GTM 摘要；起草预实分析群消息（含"缺 XX 数据"） | 粘贴、更新预实分析、发群 |
| 下周补数 | 晚到的周报照常解析，报告里标"补 W{N}" | 粘贴 |

## 3.5 PSI 表结构（从 2026-11-02 导出分析得出）

> 飞书导出的 xlsx **不带公式计算结果**，只能读公式和手填值。agent 每次运行前都要重新读一份导出来确定行号和列号，不要写死。

**列**：每张 KA / OM / OP sheet 第 5 行是周数。2026 年 W1 在 `CI` 列，第 N 周 = `CI` 往右 N−1 列（W43 = `DY`）。ALL sheet 表头在第 11 行。
**行**：第 6–47 行是品类 / 品牌小计（SUMIFS），第 48–54 行是 Others，第 55 行起是产品。每个产品 8 行，I 列是指标名：

| I 列 | 过去的周 | 当周 | 未来的周 |
|---|---|---|---|
| ST | 数值 | OM：`SUMIFS('2026 DB ST'!J, F=代码, H=KA名(A列), A=周)`；OP：手填 BI 妥投量 | `=ST FCST` |
| Promo Price | 手填 | | |
| SO | 数值 | **粘贴数值**（覆盖原来的 `=SO FCST`） | `=SO FCST` |
| Channel Inventory | `=上周 + ST − SO` | 同左 | 同左 |
| Channel Inventory Actual | 数值 | **粘贴数值** | `=上周 + ST FCST − SO FCST` |
| Channel DOS | `=Actual ÷ SO × 7` | 同左 | 同左 |
| ST FCST / SO FCST | KAM 填 | | |

**匹配键**：所有汇总都按 F 列 = G 列项目代码 & H 列内存（例 `C3NN8+256`）+ I 列指标名 SUMIFS。代码在任何一张 sheet 里拼错、多空格、或者汇总表里没有这一行，数就会**静默丢失**。

**汇总链**：
- `2026 DB ST`：贴 A 周 / B 经销商 / C 产品名 / I 客户名 / J 数量 / K（AB 物料码）或 L（ALSO、IM 的 EAN）；D–H、M–O 是 XLOOKUP（工具表1：EAN 或物料码 → 代码、内存、颜色；工具表2：客户名 → KA）
- 6 个 OM 渠道 sheet 的 ST ← `2026 DB ST`
- `OM` = EURO + MEX + X-KOM + MSH + Michain + OM Others；`OP` = Orange + Play + TMO + Plus
- `📌ALL` = OM + OP；ALL 的 SI = OM 的 SI + OP 的 ST；ALL 的 DB Inventory = 上周 + SI − ST
- `2026 DB STOCK` **没有被任何 sheet 引用**；OM 的 SI 行 2026 年全空

**已知结构问题（2026-11-02 导出）**，agent 每周自动复查：
1. KA 里有、OM / OP 里没有的代码，数据丢失。已确认有量的：MEX `P82 inbox8+256`（Pad 8），EURO `P83 inbox6+128`、`P85 inbox4+64`。运营商 sheet 另有约 10 个代码在 OP 缺行（如 `Q194+256`、`P25AE3+64`），目前还没有量。
2. MEX 小计区间和其他 sheet 不一致：SMP TTL 从第 55 行开始（漏了 Others），品牌小计从第 62 行开始（漏了第一个产品）。
3. TMO 的 `P15AE4+128` 有两个块（第 169、177 行），需要确认是不是有意分开的。
4. DB 库存没接进汇总，DB DOS 没有数据可依。

## 4. DB ST / DB STOCK（AB、ALSO、IM）

写入规则：每行填 `WEEK`、`DB`（经销商名），再贴入 产品名 / 数量 / reference / ST time。其他列都是公式，从上一周的行往下拉。
**AB 的 reference 是小米物料码（如 MZB0NOMEU），ALSO 和 IM 的 reference 是 EAN**，所以两者的公式和列位置不一样。

### 4.1 AB
- 邮件：发件人 `vrep@ab.pl`，同一封邮件里有两个附件：
  - `xiaomi_smart_sell_through_{起始YYYYMMDD}_{结束YYYYMMDD}.xlsx` → ST
  - `xiaomi_smart_stock_{日期}.xlsx` → Stock
- ST：按 `Data faktury`（日期）+ `Nazwa produktu` + `Nazwa klienta`（KA）+ `Kod produktu dostawcy`（物料码）分组，对 `Ilość` **求和（Sum）**。
- Stock：取 `Opis`（产品名）、`Stan`（库存）、`EAN` 三列。

### 4.2 ALSO
- 文件里有 `Sellout_PL`、`Stock_PL`、`IMEI_Sellout`、`IMEI_Returns` 四张 sheet。
- ST：`Sellout_PL` 的列是 CustomerName / VendorPartNumber / ProductDescription / EAN / Suma，**没有日期**。原来的做法是用 `XLOOKUP(EAN&客户, IMEI_Sellout!M:M, IMEI_Sellout!J:J)` 补日期。
  - 已知局限：XLOOKUP 只返回第一个匹配的日期。不过 PSI 是按 WEEK 汇总的，只要文件不跨周，结果就不受影响。
  - 脚本做法：直接用 `IMEI_Sellout` 明细按 日期 + 客户 + EAN 计数，再和 `Sellout_PL` 的 Suma 核对。⚠ 要确认 `IMEI_Returns` 的退货要不要扣。
- Stock：`Stock_PL`（VendorPartNumber / ProductDescription / EAN / Qty），里面有大量非手机平板的产品，要过滤。

### 4.3 IM（Ingram Micro）
- ST：原始数据不需要处理，直接用。⚠ 发件人和文件名待补。
- Stock：在另一封邮件里，`Internal W{周} 2026 Ingram SP split.xlsx`，由 Davis Miao 转发（公邮里可能没有）。

### 4.4 例外
- 公式拉完显示 0 或查不到 → 看产品名：配件（手环、手机壳、平板笔）直接删掉；新品要补进 `工具表1_SKUID对应表`，SKU 信息去 PL GTM All-In-One 查，还查不到就问 Pope Peng。
- 更新完在 GTM 群里通知一声（agent 只起草，发送前要确认）。

## 5. OP 四大运营商（Play / Plus / Orange / TMO）

### 5.1 ST：以国际 BI 为准
- 地址：`https://iwork.mioffice.cn/idi/Main/Shujing/PSI?id=all_sale_dataset` → KPI(Sale) Dataset。
- 配置：维度 = 项目 + 内存；指标 = 妥投量；筛选 = 周（静态区间，比如 2026-W41）+ 客户名称。
- 客户主体名：Play = `P4 Sp. z o.o.`；Plus / Orange / TMO ⚠ 待补。
- 写入规则：
  - 按 项目代码 + 内存 填到当周列，并标颜色；妥投量为 0 的不标。
  - 表上原来有数但 BI 查不到 → 改成 0。
  - 当周完全没有 ST → 任选一格 0 标上颜色，表示已更新。

### 5.2 SO 与 Channel Inventory Actual
在 PSI 里一起筛选 SO 和 Channel Inventory Actual 两行填写。重点跟着上周标过颜色的行更新，并留意新品。

| 运营商 | 数据来源 | 怎么匹配 | SO 列 | 库存列 | 特殊规则 |
|---|---|---|---|---|---|
| Orange | Jet Wu 整理好的表（W 列 Sales / Delivery / Stock） | 产品名 | Sales | Stock | 来源不是公邮 |
| Plus | 周报，sheet 名形如 `Stany YYYY-MM-DD` | 产品名（Model） | Sprzedaż PLK | Magazyn + sieć PLK | 是 0 的不填；非手机平板丢掉 |
| Play | 周报 `Stock and sales - week {N}.xlsx` | SKU（如 TE-XM-14T5G-PL1-BL）→ code | SPRZEDAŻ | STOCK | 原来用上周的文件 VLOOKUP；#N/A 的非手机产品不管，手机或平板按名称补 code，注意区分内存 |
| TMO | 周报 `kw_cpfr_main_Xiaomi.xlsx` | Material（如 XA-H-R121010）→ code | WydaniaTTL | ZapasTTL | 名称带 `Outlet` 的 code 记为 Outlet，带 `DEMO` 的记为 Demos，都不计入；别的后缀拿不准问盛俊 |

匹配完后按 code 汇总（相当于透视表求和），再按 PSI 的行填入并标颜色。

## 6. OM 公开市场

通用流程：
1. 只保留手机和平板。
2. 用 EAN（69 码）去 **TTL ID 工具表**（`TTL ID {日期}更新.xlsx`，列 ID / EAN / SKU，K 列 = code+内存，比如 `P12U12+512`）查出 PSI 代码。
3. 检查查不到的（#N/A）：非手机平板丢掉；手机或平板要补进 TTL 表。
4. 按 PSI 对应 sheet 里 **SO 行（从 Others 下面开始）的代码顺序** SUMIF，分别算出 SO 和 Stock。
5. **Others = 原始数据总数 − 已匹配到 PSI 行的总数**，SO 和 Stock 各算一次。
6. 在 PSI 筛选 SO 和 Channel Inventory Actual，手填 Others，下面整列粘贴数值，标颜色和字体。
7. 复核：整列求和应该等于原始数据总数。

ST：OM 的 ST 是 PSI 里的 SUMIFS 公式，从 `2026 DB ST` 按 代码 + KA + 周 取数。每周把上周那列公式复制到新的一周，再把上周那列粘贴成数值。

| KA | 怎么认出邮件 / 文件 | SO 列 | 库存列 | 特殊处理 |
|---|---|---|---|---|
| EURO | 手机和平板分两封：Iwona.Pawelek@euro.com.pl 发「tsma w{N}」附件 `Xiaomi_tele_w{N}.xlsx`；Justyna.Lukasiewicz@euro.com.pl 发「XIAOMI TABLETY W{N}」附件 `Xiaomi_TAB_W{N}.xlsx` | Sprzedaż | Stan | 两个文件合并；删掉手环、手表（OPASKA / SMARTWATCH） |
| MEX | 发件人域名带 TERG；表头有 Data od / Data do | SPRZEDAŻ | STAN | 只保留 GRUPA = SMARTFONY / TABLETY。**ST 要另加 BI 的妥投量**：先删掉上周公式末尾手动加的 `+N`，再把本周 BI 的数加到公式后面 |
| MSH | 发件人域名 media-saturn | Ilość | Stan magazynowy | 看 `Total` sheet，按 Grupa towarowa 只保留手机和平板 |
| X-KOM | ⚠ 待补 | SELL_OUT_QTY | FINAL_INVENTORY_QTY | 文件里有所有周（Week 列，如 2026W40），只取当周 |
| OM Others：KTR | 域名 komputronik | sprzedaż W{N} | stan W{N} | **EAN 开头少了 6**：`IF(LEFT(EAN,1)="6",EAN,"6"&EAN)`；品类（kategoria_nazwa）只保留 Telefony GSM、Tablety z systemem Android |
| OM Others：MAX | 域名 gtpoland | Sprzedaż | Stan magazynowy | 三家匹配后合并 |
| OM Others：Auchan | 文件名 `Xiaomi T{N} 日期.xlsx` | ILOŚĆ R | Zapas Bieżący + Zapas Magazynowy（门店 + 仓库） | 2026-11-04 确认相加；历史周用 `tools.auchan_backfill` 补算 |
| Michain：GG、ROLV | ⚠ 发件人待补 | Qty | stock | 用 Product ID 查代码；筛选 E/T = T；SO 和 stock 分在两张表，汇总时分开贴 |

## 7. 统一映射表与新品

现在的对照关系分散在：`工具表1_SKUID对应表`（PSI 内，DB 用）、TTL ID 工具表（EAN → code，OM 用）、Play 上周的文件（SKU → code）、TMO 上周的文件（Material → code）、Orange 和 Plus 靠产品名人工对。

**改成一张 `mapping/mapping.xlsx`**：

| 列 | 说明 |
|---|---|
| source | 来源类型：EAN / AB_MATERIAL / PLAY_SKU / TMO_MATERIAL / PLUS_NAME / ORANGE_NAME |
| key | 识别码，例如 6932554483654、MZB0NOMEU、TE-XM-14T5G-PL1-BL |
| psi_code | PSI 代码（F 列 = 项目代码 + 内存，例如 MOSCOW6+256） |
| category | phone / pad / exclude（配件、Outlet、Demo 等） |
| product_name | 产品名（方便人看） |
| added_on / added_by | 什么时候、由谁确认 |

规则：
- 项目代码全部用系统叫法，不用别名，例如 Redmi Note 17 Pro 5G 写 MOSCOW，不写 Q16。
- 代码和内存前后不能有空格。

**新品加进 PSI（2.3）**：
- 每个产品插 8 行（ST FCST、SO FCST、ST、Promo Price、SO、Channel Inventory、Channel Inventory Actual、Channel DOS 等）。
- F 列是公式 `=G&H`，不要手写；后面月份的公式复制上一个产品的。
- 除了各 KA 的 sheet，**也要同步加到 OM / OP / ALL**。
- 新品加不加由 GTM 决定，agent 只提示"发现疑似新品"，不自己加行。

## 8. 写回 PSI 的方式

PSI 表是**飞书在线表格**（owner Pope Peng，L3-Confidential），多人同时编辑。

| 方案 | 说明 | 结论 |
|---|---|---|
| A. 待粘贴表 | agent 输出列顺序和 PSI 一致的 xlsx 和核对报告，人工筛选后整列粘贴、标颜色 | **第一期用这个** |
| B. 浏览器自动化 | 在已登录飞书的浏览器里逐格填写 | 备选，速度慢，结构一变容易错位 |
| C. 飞书开放平台 API | 按单元格写入，不会覆盖别人的修改 | 第二期，需要管理员审批 + owner 授权 |
| D. 下载、本地改、再上传 | 会覆盖别人的修改、弄坏格式 | 禁止 |

## 9. 校验与异常

- 每个来源：原始总数（只算手机平板）= 匹配总数 + Others。
- 映射查不到的列一份清单：产品名、识别码、来源、AI 给的候选 code、是否疑似新品。
- 环比异常：某个 code 的 SO 或库存比上周变化超过 ±50%，并且绝对值大于 50 台（阈值待定），就标出来。
- 文件的时间段要和目标周一致（例：MEX 的 Data od / Data do、AB 文件名里的日期）。
- 缺了某家的周报 → 报告里列出"未收到"，并附上起草好的催数消息（收件人按 §2 联系人表）。
- 发群通知、改映射表、往 PSI 写数，都要先经人确认。

**结构校验**（读 PSI 导出，每周跑一次）：
- 每个 KA sheet 的代码集合 ⊆ OM 或 OP 的代码集合 ⊆ ALL 的代码集合；不满足就列出缺的代码，以及它今年累计的 ST / SO / 库存。
- 同一张 sheet 里同一个代码 + 指标只出现一次。
- G / H 列前后没有空格；F 列是 `=G&H` 公式，不是手写。
- 小计 SUMIFS 的起始行 = Others 起始行（TTL）或第一个产品行（品牌），各 sheet 一致。
- 待粘贴的代码在目标 sheet 里一定有对应的行；没有就归入"疑似新品"，不进 Others。

**数值校验**：
- 库存恒等式：|本周 Actual −（上周 Actual + ST − SO）| 超过阈值的代码，列出来。常见原因是映射错位、退货、调拨，或者周报口径变了。
- 待粘贴表各列合计 = 原始数据（手机 + 平板）合计。

## 9.5 GTM 摘要（报告最后一页）

agent 在核对报告最后附一页，给 GTM 做决策用：

| 指标 | 算法 | 提示条件（阈值待定） |
|---|---|---|
| Channel DOS | Actual ÷ 近 4 周平均 SO × 7，按 KA × 机型 | > 60 天压货；< 14 天断货风险 |
| SO 预实 | SO ÷ SO FCST − 1，按 KA 和 TTL | 偏差 > ±20% |
| 渠道净进货 | ST − SO，连续 3 周为正 | 压货前兆 |
| SO 周环比 | 本周 ÷ 上周 − 1 | ±50% 且 > 50 台 |
| 新品上量 | 上市后前 8 周 SO 曲线 | 低于 FCST |
| EOL 库存 | 停产机型剩余渠道库存 + DOS | 需要清货促销 |
| DB DOS | DB STOCK ÷ DB 卖给 KA 的周 ST × 7（数据来自 DB STOCK sheet） | 决定 SI |

## 10. 待补材料与待确认问题

**材料**（放到 `samples/W{周}/`，可以先脱敏）
- [ ] 同一周所有来源的原始文件：AB×2、ALSO、IM（ST + SP split）、Play、Plus、TMO、Orange、EURO×2、MEX、MSH、X-KOM、KTR、MAX、Auchan、GG、ROLV
- [ ] TTL ID 工具表 xlsx
- [ ] PSI 表导出成 xlsx（带公式），最好有同一周更新前后两版
- [ ] `工具表1_SKUID对应表` 导出

**问题**
- [ ] ALSO、IM、Play、Plus、TMO、X-KOM、Auchan、GG、ROLV 的发件人和标题格式，以及日报和周报怎么区分
- [ ] Plus / Orange / TMO 在国际 BI 里的客户主体名
- [ ] ALSO 的 ST 要不要扣 IMEI_Returns
- [ ] 周定义是否为 ISO 周
- [ ] Power Automate 能不能用；Orange（Jet Wu）和 IM 库存（Davis Miao）能不能改成直接发到公邮
- [ ] 公司是否允许用 AI 处理 L3 数据
- [ ] GTM 群在飞书还是企业微信

## 11. 实施路线

1. **第一期 AB 跑通**：下载 → 解析 → 映射 → 待粘贴表 → 报告，拿一周真实数据对数。
2. 加 ALSO、IM，完成 DB 部分。
3. OM：EURO → MSH → MEX → X-KOM → Others → Michain。
4. OP：Play / TMO（SKU / Material 映射）→ Plus / Orange（按名称映射）→ BI 半自动。
5. 定时运行 + 就绪看板 + 催数草稿；评估飞书 API 写回（方案 C）。
6. GTM 摘要；SO 预实分析的新一周列也由 agent 生成待粘贴版。

**OM + OP 周报处理（已完成 2026-11-02）**：`psi_tool/`，配置在 `config/sources.yaml`，用法见 `使用说明.md`。用 W39–W41 的真实文件和 PSI 已填的数逐格比对：EURO、MEX、MSH、X-KOM、Michain、Play、Orange（库存）一致；剩下的差异都已查明，是 PSI 原来填错了，见 `docs/各家周报列说明.md`。DB（AB / ALSO / IM）也已完成，W39–W41 和 PSI 逐行比对，只剩冲正"计数 vs 求和"和三处零星 1 台的手工改动。

**第 0 步（已完成 2026-11-02）**：`checks/structure_check.py` 实现了 §9 的结构校验和库存恒等式。每次拿到新的 PSI 导出就跑一次，报告写到 `output/structure_check_{日期}.md`（2026-11-02 那份已移到 `前期资料/分析报告/`）；有 🔴 级问题时退出码为 1，方便以后接进定时任务。

```
.venv/bin/python -m checks.structure_check
```

## 12. Agent 清单（2026-11-03 更新）

| # | 功能 | 状态 | 怎么用 |
|---|---|---|---|
| 1 | 收邮件：公邮周报自动下载到一个文件夹 | 待做（W44、W45 并行跑通后再做；方案：Power Automate 存附件到 OneDrive） | — |
| 2 | **整理下载文件夹**：按内容认来源和周，归档到 `W{周}原始文件/`；邮件附件、zip 自动取出；其他周、同名不同内容、认不出的单独放并记录 | ✅ 已完成 | `tools.sort_downloads` |
| 3 | 解析 + 映射 + 汇总：DB / OM / OP 共 15 个来源 → 待粘贴表 + 核对报告 | ✅ 已完成（W39–W41 验证） | `psi_tool.run` |
| 4 | 人工抽查辅助：附上周 PSI 的数，变化大的行标黄 | ✅ 已完成 | 待粘贴表右侧几列 |
| 5 | **和 Excel 手工结果对数**：生成模板 → 填结果 → 逐格列出差异 | ✅ 已完成 | `tools.compare_manual` |
| 6 | 给 Excel 用的 PSI 码对照表（含去 DEMO 公式） | ✅ 已完成 | `tools.export_mapping` |
| 7 | PSI 表结构校验（漏行、重复代码、小计区间） | ✅ 已完成 | `checks.structure_check` |
| 8 | 新型号：AI 给候选代码 → 人确认 → 写入 mapping.xlsx | 半自动（报告列出未匹配，候选靠名称规则） | 报告"未匹配"一节 |
| 9 | 写回飞书 PSI（API） | 第二期 | — |
| 10 | 群通知 / 催数消息草稿 | 待做 | — |

## 附：建议目录结构

```
~/Desktop/PSI_agent/
├── PSI周更Agent需求文档.md   ← 本文件
├── samples/                  ← 原始样本（不进 git）
├── inbox/W{周}/{来源}/        ← 每周下载的附件和邮件正文
├── config/sources.yaml       ← 每个来源：发件人、标题规则、sheet、列名、过滤条件
├── mapping/mapping.xlsx      ← 统一映射表
├── parsers/                  ← 每个来源一个解析器
├── output/W{周}_paste.xlsx   ← 待粘贴表
└── output/W{周}_report.md    ← 核对报告
```
