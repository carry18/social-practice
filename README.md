# 实践部经费助手（网页版）

清华大学学生就业服务协会 · 实践部。纯静态网页，浏览器打开即用，把
**预算编制 → 决算打款 → 报名推送 → 总览监控** 串成一条线，另含实践地图。

- 现网：<https://carry18.github.io/social-practice/>（仓库 `carry18/social-practice`，`main` 根目录发布）
- 预算的计算、xlsx 与 Word 生成全部在**浏览器本地**完成；决算打款、报名推送、预决算汇总需经腾讯云函数转发到金山文档 / 金数据。
- 没有数据库、没有登录：数据留在金山文档的表里，网页只负责读写与生成。

## 模块

| 模块 | 做什么 |
| --- | --- |
| 📊 总览监控 | 从金山读「预算汇总」「决算汇总」→ 年 × 期次汇总表、预决算趋势图、内嵌实践地图 |
| 📝 预算编制 | 手动填写或上传院系申报表 xlsx → 生成预览 → 下载预算表 xlsx / 公文版 Word → 同步到金山；也能从金山读回继续改 |
| 💰 决算打款 | 上传决算表 docx → 提取「院系联合组织」支队 → 在金山打款汇总表里新建工作表；再据此生成《经费拨付事宜的说明》Word 并归档到金山 |
| 📮 报名推送 | 按期次创建/复用金数据报名表 → 生成报名二维码 → 下载含二维码的推送 Word；可把支队与答卷同步进实践台账 |
| 🗺️ 实践地图 | `map.html` 独立页（总览页内嵌），ECharts 中国地图展示各期实践地点，可拖入统计表 xlsx 即时预览 |

期次统一用「寒假 / 暑期 / 春季学期中 / 秋季学期中」，年份单独选择。

## 预算口径

| 项 | 口径 |
| --- | --- |
| 交通费 | 行程链逐段 12306 二等座最高票价之和（北京 → A → B → 北京）。默认用 `fares.json` 快照；勾选「优先实时查询」则逐段查 12306 |
| 住宿 | 150 元/人·晚 × 人数 × (天数-1) × 比例；点位或院系栏含「承担住宿」「住宿自理」→ 记 0 |
| 保险 | 30 元/人·周 × 人数 × ⌈天数/7⌉ × 比例 |
| 比例 | 承办单位含「就协」「就业协会」→ 全额；院系联合 → 2/3 |

生成的 xlsx 与工作区 Python 工具同口径：双行表头，I–S 为呈现列，T–AC 为计算辅助列，AD 记逐段取价依据；公式带缓存值，Excel / WPS 打开可继续编辑。
未识别的地名不会静默放过——预览、警告和同步前都会拦下，并提示补票价。

## 外部依赖

- **CDN（jsDelivr）**：xlsx-js-style、docx、mammoth、echarts、pizzip、qrcode-generator。离线环境页面打不开。
- **腾讯云函数**（地址写在 js 里，函数内不存敏感信息）：
  - 金山预算表读写与预决算汇总（`app.js`、`overview.js`、`signup.js`）
  - 金山决算与打款表读写（`overview.js`、`payout.js`、`note.js`）
  - 金数据报名表创建 / 同步（`signup.js`）
  - 12306 查价代理（`scf_fare_proxy.js`）
- **金山文档 AirScript**：预算表写入、打款表新建、拨付说明归档；实践台账的建表 / 入库脚本是 `ledger-template.js`。

## 部署与更新

GitHub Pages 从 `main` 分支根目录发布，push 后 1–2 分钟生效：

```bash
git add -A && git commit -m "..." && git push
```

- 改完 `app.js` / `signup.js` 等文件，记得同步改 `index.html` 里对应的 `?v=` 版本号，否则浏览器会吃缓存。
- 本地预览：`python3 -m http.server 8000`，必须走 http（`file://` 读不到 `fares.json`）。
- 仓库须公开，免费版才有 Pages 和 Actions。

## 票价自动刷新

`.github/workflows/refresh-fares.yml` 只有两种触发：

- **定时**：每周一 09:00（北京时间）跑 `fare_query.py --update`，查 12306 预售期 14 天窗口的二等座最高价，有变化才自动提交；
- **手动**：仓库 Actions → 「刷新12306票价库」→ Run workflow（可勾 `force` 忽略「今天已刷新」的判断）。

已取消 push 触发（每次推送都跑会和 Pages 构建抢队列）；当天刷新过则自动跳过；查询失败不影响网页，继续用上次快照。

新增 / 维护票价分段：

```bash
python3 fare_query.py --add 北京-新城市 新城市-北京 --days 14   # 有直达车
python3 fare_query.py --set 北京-某县=1234 --note "中转方案"     # 无直达车，人工录入
git add fares.json && git commit -m "新增分段" && git push
```

## 文件说明

| 文件/目录 | 作用 |
| --- | --- |
| `index.html` | 页面骨架、样式与四个模块的 DOM |
| `app.js` | 预算编制：录入解析、票价匹配、xlsx / Word 生成、金山同步 |
| `overview.js` | 总览监控：预决算汇总与趋势图 |
| `payout.js` | 决算 docx → 打款院系表 |
| `note.js` | 拨付说明 Word + 院系汇总表 |
| `signup.js` | 报名推送：建表、二维码、推送 Word、台账同步 |
| `enterprise.js` / `enterprise_index.json` | 企业参考侧栏（按地点筛历年合作单位，一键写入点位列） |
| `ledger-template.js` | 金山 AirScript：实践台账建表 / 增量入库 / 淘汰捞起归位 |
| `map.html` / `map.js` / `map_data.json` | 实践地图页与数据 |
| `fares.json` | 票价库快照（Actions 每周刷新，网页与 Python 工具共用） |
| `city_province.json` / `china.json` | 城市–省份与地图边界数据 |
| `dept_alias.json` | 院系别名库（新增院系只改这里，不用改代码） |
| `ledger-map.json` | 每期台账配置：hook、报名截止日、是否启用 |
| `fare_query.py` / `station_codes.json` | 供 Actions 刷新票价用（网页本身不需要） |
| `scf_fare_proxy.js` | 云函数端 12306 查价代理源码 |
| `.github/workflows/refresh-fares.yml` | 票价自动刷新 |
| `templates/` | 报名推送 Word 模板 |
| `docs/` | 设计文档：自动化蓝图、实践预决算自动化流程、12306 票价查询 |
| `tools/` | 本地工具链（见下） |

## 内部工具 tools/

| 脚本 | 作用 |
| --- | --- |
| `rebuild_budget.py` | 由只含 A–H 的预算表重建 I–S / T–AC 列（`--refresh` 查实时票价） |
| `gen_budget_docx.py` | 预算 xlsx →《实践预算》公文 Word（只生成就协部分） |
| `gen_budget_template.py` | 生成空白预算表模板 `预算表模板.xlsx` |
| `budget_lookup.py` | 决算端按支队取预算封顶基准，供审核比对 |
| `build_enterprise_index.py` | `自动化/企业台账/companies.csv` → `enterprise_index.json`（企业侧栏数据） |
| `fare_query.py` / `station_codes.json` | 12306 票价查询与城市站名表 |

```bash
python3 tools/rebuild_budget.py 院系申报表.xlsx -o 预算表.xlsx --refresh --fares fares.json
python3 tools/gen_budget_docx.py 预算表.xlsx -o 预算公文.docx
python3 tools/budget_lookup.py 预算表.xlsx 合肥 上海
python3 tools/build_enterprise_index.py
```

## 实践台账（金山 AirScript）

1. 金山文档新建一个空白表格 → 效率 → AirScript → 新建「文档共享脚本」；
2. 贴入 `ledger-template.js` → 保存；
3. 更多 → 复制 webhook 链接；
4. 登记到 `ledger-map.json`（期次 → `hook` / `until` / `enabled`），页面「报名推送 → 🧾 建/同步台账」即可用。
