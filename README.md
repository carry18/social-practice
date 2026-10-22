# 实践支队预算表生成器（网页版）

纯静态网页：院系/支队同学在浏览器里填写支队信息（或上传院系申报表 xlsx），
页面在**本地浏览器内**计算并生成完整预算表 xlsx 下载——数据不经过任何服务器。

- 交通费 = 行程链逐段 12306 二等座最高票价之和（随页面打包 fares.json 快照）
- 住宿 = 300 标间 × ⌈人数/2⌉ × (天数-1) × 比例；保险 = 30 元/人/周
- 比例：承办单位含"就协"/"就业协会" → 全额，否则 2/3
- 生成的 xlsx 含公式 + 缓存值，Excel/WPS 打开可继续编辑，口径与工作区 Python 工具一致

## 部署到 GitHub Pages（一次性，约 5 分钟）

1. 在 GitHub 新建**公开**仓库（如 `practice-budget`）；
2. 把本目录（`web/`）下的**全部内容**上传到仓库根目录，两种方式任选：
   - 网页上传：仓库页 → Add file → Upload files → 拖入本目录所有文件（含 `.github/` 隐藏目录，可用命令行方式）；
   - 命令行：
     ```bash
     cd web
     git init && git add -A && git commit -m "预算表生成器"
     git branch -M main
     git remote add origin git@github.com:<你的用户名>/<仓库名>.git
     git push -u origin main
     ```
3. 仓库 → **Settings → Pages** → Source 选 `Deploy from a branch`，
   Branch 选 `main` / `/ (root)` → Save；
4. 等 1-2 分钟，访问 `https://<你的用户名>.github.io/<仓库名>/` 即可使用。

> 注意：必须**公开**仓库免费版才有 Pages 和 Actions 定时任务；
> 页面必须经 http(s) 访问（直接双击 index.html 的 file:// 方式读不到 fares.json）。

## 票价自动刷新（完全智能，无需人工）

仓库内含 `.github/workflows/refresh-fares.yml`，三种触发方式：
- **定时**：每周一 09:00（北京时间）自动运行 `fare_query.py --update`，
  联网查 12306 预售期 14 天窗口的二等座最高票价，有变化才自动提交；
- **上传触发**：你每次 push 更新网页/其他文件时顺带刷新（票价文件本身的提交不触发，防循环）；
- **手动**：每期正式预算前到仓库 Actions 页 → "刷新12306票价库" → Run workflow；
- 某次查询失败不影响网页，继续使用上次的票价。

## 新增/维护票价分段

新实践地城市直接在仓库 Actions 手动跑之前，先本地补充分段再提交：

```bash
python3 fare_query.py --add 北京-新城市 新城市-北京 --days 14      # 有直达车
python3 fare_query.py --set 北京-某县=1234 --note "中转方案"        # 无直达车人工录入
git add fares.json && git commit -m "新增分段" && git push
```

## 本地预览

```bash
cd web && python3 -m http.server 8000
# 浏览器打开 http://localhost:8000
```

## 文件说明

| 文件/目录 | 作用 |
|------|------|
| `index.html` / `app.js` | 网页本体（计算逻辑与 Python 工具一致，已对拍验证） |
| `fares.json` | 票价库快照（Actions 每周自动更新，网页与内部工具共用） |
| `fare_query.py` / `station_codes.json` | 供 Actions 刷新票价用（网页本身不需要） |
| `.github/workflows/refresh-fares.yml` | 自动刷新配置 |
| `docs/` | 设计文档：实践预决算自动化流程设计、12306 票价查询设计 |
| `tools/` | 内部工具链：rebuild_budget.py（重建预算表）、gen_budget_docx.py（公文生成）、budget_lookup.py（决算取数）、gen_budget_template.py（模板）、预算表模板.xlsx |

**内部工具用法**（在仓库根目录执行，票价库指向根目录 fares.json）：

```bash
python3 tools/rebuild_budget.py 院系申报表.xlsx -o 预算表.xlsx --refresh --fares fares.json
python3 tools/gen_budget_docx.py 预算表.xlsx -o 预算公文.docx
python3 tools/budget_lookup.py 预算表.xlsx 合肥 上海
```
