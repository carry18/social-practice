#!/usr/bin/env python3
"""12306 高铁二等座最高票价查询脚本（配套 fares.json 票价缓存）。

用法：
    python3 fare_query.py 北京 上海 [--date 2026-10-24]   # 查单日最高二等座价
    python3 fare_query.py --check fares.json              # 对比缓存价与实时价
    python3 fare_query.py --update fares.json             # 用实时价刷新 segments
    python3 fare_query.py --add 北京-合肥 合肥-上海        # 查询并补入新分段

接口（kyfw.12306.cn，无需登录）：
    1. GET /otn/leftTicket/init                        建立会话 cookie（每次会话先访问）
    2. GET /otn/leftTicketPrice/queryAllPublicPrice    票价查询（一次返回当日全部车次各席别价格）
       参数: leftTicketDTO.train_date=YYYY-MM-DD
             leftTicketDTO.from_station=三字码
             leftTicketDTO.to_station=三字码
             purpose_codes=ADULT
    3. GET /otn/resources/js/framework/station_name.js  全国车站名→三字码表（本地缓存）

注意：
    - 价格字段 ze_price 单位为"角"（如 "8200" = 820.0 元），末位为小数；
    - 12306 预售期约 15 天，超出预售期的日期自动回退到最近可售日（票价按车次定价，
      与日期基本无关，回退不影响"最高价"口径，浮动票价线路以查询日为准）；
    - 出发/到达按城市聚合：查"北京(BJP)→上海(SHH)"会返回北京南→上海虹桥等全部同城车次，
      正好满足"城市到城市取最高价"的预算口径。
"""
import argparse, json, re, sys, time, warnings
from datetime import date, timedelta
from pathlib import Path

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings("ignore", message="Unverified HTTPS request")

HERE = Path(__file__).resolve().parent
STATION_CACHE = HERE / "station_codes.json"
STATION_JS_URL = "https://kyfw.12306.cn/otn/resources/js/framework/station_name.js"
INIT_URL = "https://kyfw.12306.cn/otn/leftTicket/init"
PRICE_URL = "https://kyfw.12306.cn/otn/leftTicketPrice/queryAllPublicPrice"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Referer": "https://kyfw.12306.cn/otn/leftTicket/init",
    "X-Requested-With": "XMLHttpRequest",
    "Accept": "application/json, text/javascript, */*; q=0.01",
}
PRESALE_DAYS = 15          # 12306 预售期（含当天）
REQUEST_GAP = 1.0          # 相邻请求间隔（秒），防反爬
TIMEOUT = 12


# ---------- 车站名 → 三字码（含官方所属城市归属） ----------

def load_station_data(force=False):
    """返回 (codes, city_of, name_of)：
    codes   {站名: 三字码}
    city_of {站名: 所属城市}   ← station_name.js 官方字段（如 盛泽→苏州、福田→深圳、汉口→武汉）
    name_of {三字码: 站名}
    优先本地缓存，缺失/旧版缓存/强制时联网拉取 station_name.js。
    """
    if STATION_CACHE.exists() and not force:
        data = json.load(open(STATION_CACHE, encoding="utf-8"))
        if data.get("version") == 2:
            return data["codes"], data["city_of"], data["name_of"]
    text = requests.get(STATION_JS_URL, headers=HEADERS, timeout=TIMEOUT, verify=False).text
    codes, city_of = {}, {}
    for rec in text.split("@")[1:]:
        parts = rec.split("|")
        if len(parts) >= 8:
            name, code, city = parts[1], parts[2].upper(), parts[7]
            codes[name] = code
            city_of[name] = city
    name_of = {c: n for n, c in codes.items()}
    json.dump({"version": 2, "codes": codes, "city_of": city_of, "name_of": name_of},
              open(STATION_CACHE, "w", encoding="utf-8"), ensure_ascii=False, sort_keys=True)
    return codes, city_of, name_of


def to_telecode(name, codes):
    """站名/三字码 → 三字码。城市名精确匹配同名主站（如 上海→SHH）。"""
    name = name.strip()
    if re.fullmatch(r"[A-Za-z]{3}", name):
        return name.upper()
    if name in codes:
        return codes[name]
    # 兜底：前缀匹配（如 "北京" → 北京/北京南…），取站名最短者为主站
    cands = sorted((n for n in codes if n.startswith(name)), key=len)
    if cands:
        return codes[cands[0]]
    raise KeyError(f"未识别车站: {name}（可删除 {STATION_CACHE.name} 后重试刷新站表）")


# ---------- 票价查询 ----------

def parse_price(raw):
    """'8200' → 820.0（单位角，末位小数）。无效值返回 None。"""
    if not raw or raw == "--" or not raw.isdigit():
        return None
    return float(raw[:-1] + "." + raw[-1])


class FareClient:
    """12306 会话：先 init 取 cookie，再查价；自带重试与限速。"""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self.codes, self.city_of, self.name_of = load_station_data()
        self._inited = False
        self._last_req = 0.0

    def _query_city(self, name):
        """查询输入（城市名/站名/三字码）→ 官方所属城市名，用于同城过滤。"""
        stn = self.name_of.get(to_telecode(name, self.codes))
        return self.city_of.get(stn, name.strip())

    def _throttle(self):
        gap = REQUEST_GAP - (time.time() - self._last_req)
        if gap > 0:
            time.sleep(gap)
        self._last_req = time.time()

    def _init(self):
        self._throttle()
        self.session.get(INIT_URL, timeout=TIMEOUT, verify=False)
        self._inited = True

    def query_day(self, from_name, to_name, day):
        """查单日 (from→to) 全部车次二等座价。返回 [(price, train_code, 发站, 到站, 历时min)]。

        12306 按城市聚合返回同城所有车站。同城判定用 station_name.js 官方
        "所属城市"字段（如 盛泽属苏州、福田属深圳、汉口属武汉），比站名前缀
        匹配更准：不会漏掉不以城市名开头的同城站，也不会放进异城站。

        在此基础上剔除"绕路车"：历时超过 max(2×当日最短历时, 最短历时+60分钟)
        的车次不计入最高价（如 盛泽→无锡 570 元、苏州南→无锡 307 元的环线车，
        而直达仅 10-20 分钟）。
        """
        if not self._inited:
            self._init()
        params = {
            "leftTicketDTO.train_date": day.isoformat(),
            "leftTicketDTO.from_station": to_telecode(from_name, self.codes),
            "leftTicketDTO.to_station": to_telecode(to_name, self.codes),
            "purpose_codes": "ADULT",
        }
        last_err = None
        for attempt in range(3):
            self._throttle()
            try:
                resp = self.session.get(PRICE_URL, params=params, timeout=TIMEOUT, verify=False)
                if resp.status_code != 200 or "error" in str(resp.url):
                    raise RuntimeError(f"HTTP {resp.status_code} {resp.url}")
                data = resp.json().get("data") or []
                rows = []
                for item in data:
                    dto = item.get("queryLeftNewDTO", {})
                    p = parse_price(dto.get("ze_price", ""))
                    if p is None:
                        continue
                    h, _, m = (dto.get("lishi") or "99:59").partition(":")
                    mins = int(h) * 60 + int(m) if h.isdigit() and m.isdigit() else 9999
                    rows.append((p, dto.get("station_train_code", "?"),
                                 dto.get("from_station_name", ""), dto.get("to_station_name", ""),
                                 mins))
                fc, tc = self._query_city(from_name), self._query_city(to_name)
                main = [r for r in rows
                        if self.city_of.get(r[2]) == fc and self.city_of.get(r[3]) == tc]
                if not main:  # 官方城市归属缺失时回退到站名前缀匹配
                    main = [r for r in rows
                            if r[2].startswith(from_name) and r[3].startswith(to_name)] or rows
                if main:
                    shortest = min(r[4] for r in main)
                    cap = max(shortest * 2, shortest + 60)
                    direct = [r for r in main if r[4] <= cap]
                else:
                    direct = main
                return direct or rows
            except (requests.RequestException, ValueError, RuntimeError) as e:
                last_err = e
                self._inited = False  # 可能被反爬踢掉会话，下次重建
                time.sleep(2 * (attempt + 1))
        raise RuntimeError(f"查询失败（已重试3次）: {last_err}")

    def max_fare_range(self, from_name, to_name, start=None, days=1):
        """窗口内最高二等座价：从 start 起逐日查 days 天，取全窗口最高价。

        12306 预售期约 15 天：超出预售期的日子自动跳过（days_skipped 计数），
        因此可以直接传"拟定出行窗口"，查不到的日子自动落到可售日内。
        浮动票价按星期几波动，窗口 ≥7 天可覆盖全周口径。

        返回 dict: {price, train, from_station, to_station, query_date(最高价出现日),
                    window, days_queried, days_skipped, date_adjusted}
        """
        today = date.today()
        sale_end = today + timedelta(days=PRESALE_DAYS - 1)
        first = start or (today + timedelta(days=1))
        last = first + timedelta(days=days - 1)
        best = None          # (price, code, from_stn, to_stn, mins, count, day)
        queried = skipped = 0
        d = first
        while d <= last:
            if d > sale_end or d < today:
                skipped += 1
                d += timedelta(days=1)
                continue
            rows = self.query_day(from_name, to_name, d)
            if rows:
                queried += 1
                day_best = max(rows)
                if best is None or day_best[0] > best[0]:
                    best = day_best + (len(rows), d)
            d += timedelta(days=1)
        if best is None:
            # 整窗无数据（如全部落在预售期外）：回退到预售期最后一天单日查
            d = sale_end
            rows = self.query_day(from_name, to_name, d)
            if not rows:
                raise RuntimeError(f"{from_name}→{to_name} 未查到任何有二等座的车次（无直达车？）")
            day_best = max(rows)
            best = day_best + (len(rows), d)
            queried, skipped = 1, 0
        return {
            "price": best[0],
            "train": best[1],
            "from_station": best[2],
            "to_station": best[3],
            "train_count": best[5],
            "query_date": best[6].isoformat(),
            "window": f"{first.isoformat()}~{last.isoformat()}" if days > 1 else None,
            "days_queried": queried,
            "days_skipped": skipped,
            "date_adjusted": skipped > 0 or (start is not None and best[6] != start),
        }

    def max_fare(self, from_name, to_name, day=None):
        """单日最高二等座价（max_fare_range 的单日包装）。"""
        return self.max_fare_range(from_name, to_name, start=day, days=1)


# ---------- fares.json 维护 ----------

def load_fares(path):
    return json.load(open(path, encoding="utf-8"))


def save_fares(path, fares):
    json.dump(fares, open(path, "w", encoding="utf-8"),
              ensure_ascii=False, indent=2, sort_keys=False)


def _meta_of(r):
    """从查询结果提取 fares.json meta 留痕。"""
    m = {"query_date": r["query_date"], "train": r["train"]}
    if r.get("window"):
        m["window"] = r["window"]
        m["days_queried"] = r["days_queried"]
    return m


def _fmt_tag(r):
    """输出用来源标签，如 'G3937 @2026-10-22' 或 'G3937 窗口10-22~11-04·最高日10-28'。"""
    if r.get("window"):
        w0, w1 = r["window"].split("~")
        return (f"{r['train']} 窗口{w0[5:]}~{w1[5:]}·最高日{r['query_date'][5:]}"
                f"·查{r['days_queried']}天")
    return f"{r['train']} @{r['query_date']}"


def cmd_check_update(args, update=False):
    fares = load_fares(args.fares)
    client = FareClient()
    segs = fares.get("segments", {})
    changed, failed = 0, []
    meta = fares.setdefault("meta", {})
    start = date.fromisoformat(args.date) if args.date else None
    for key in list(segs):
        frm, to = key.split("-", 1)
        manual = meta.get(key, {}).get("manual")
        try:
            r = client.max_fare_range(frm, to, start=start, days=args.days)
        except Exception as e:
            failed.append((key, str(e)))
            if manual:
                print(f"⚑ {key:12} = {segs[key]} [人工录入"
                      f"{('·' + meta[key]['note']) if meta[key].get('note') else ''}"
                      f"·仍无直达，{'待人工复核' if update else ''}]")
            else:
                print(f"✗ {key:12} 查询失败: {e}")
            continue
        old = segs[key]
        mark = "" if abs(r["price"] - old) < 0.01 else f"  ← 缓存 {old}"
        if manual:
            mark += "  ← 人工录入价，现已有直达车实时价"
        adj = "（部分日期超预售期）" if r["days_skipped"] else ""
        print(f"{'→' if update else ' '} {key:12} 实时最高 {r['price']:>7.1f} "
              f"[{_fmt_tag(r)}]{adj}{mark}")
        if update and (manual or abs(r["price"] - old) >= 0.01):
            segs[key] = r["price"]
            changed += 1
        meta[key] = _meta_of(r)
    if update:
        meta["updated_at"] = date.today().isoformat()
        save_fares(args.fares, fares)
        print(f"\n已更新 {changed} 条 → {args.fares}")
    if failed:
        print(f"\n{len(failed)} 条查询失败，缓存价未动。", file=sys.stderr)


def cmd_add(args):
    fares = load_fares(args.fares)
    segs = fares.setdefault("segments", {})
    meta = fares.setdefault("meta", {})
    client = FareClient()
    start = date.fromisoformat(args.date) if args.date else None
    for seg in args.add:
        frm, to = seg.split("-", 1)
        try:
            r = client.max_fare_range(frm, to, start=start, days=args.days)
        except Exception as e:
            print(f"✗ {seg}: {e}", file=sys.stderr)
            continue
        segs[seg] = r["price"]
        meta[seg] = _meta_of(r)
        print(f"✓ {seg} = {r['price']} [{_fmt_tag(r)}]"
              + ("（部分日期超预售期）" if r["days_skipped"] else ""))
    meta["updated_at"] = date.today().isoformat()
    save_fares(args.fares, fares)
    print(f"已写入 {args.fares}")


def cmd_set(args):
    """人工录入票价（无直达车等联网查不到的段）：记入 fares.json 并打 manual 标记，
    预算 --refresh 时会自动重试联网查询，仍查不到则提示"人工价待确认"。"""
    fares = load_fares(args.fares)
    segs = fares.setdefault("segments", {})
    meta = fares.setdefault("meta", {})
    for item in args.set:
        seg, _, price = item.partition("=")
        if not price:
            sys.exit(f"--set 格式应为 起-止=价格，收到: {item}")
        segs[seg] = float(price)
        meta[seg] = {"manual": True, "note": args.note or "",
                     "updated_at": date.today().isoformat()}
        print(f"✓ {seg} = {price} [人工录入{('·' + args.note) if args.note else ''}·预算时需复核]")
    meta["updated_at"] = date.today().isoformat()
    save_fares(args.fares, fares)
    print(f"已写入 {args.fares}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("stations", nargs="*", help="出发站 到达站（单段查询）")
    ap.add_argument("--date", help="窗口起始日期 YYYY-MM-DD（默认明天；超预售期的日子自动跳过）")
    ap.add_argument("--days", type=int, default=1,
                    help="查询窗口天数（默认 1 天；≥7 天可覆盖全周浮动票价，预售期内最多可查约 14 天）")
    ap.add_argument("--fares", default=str(HERE / "fares.json"))
    ap.add_argument("--check", action="store_true", help="对比缓存价与实时价")
    ap.add_argument("--update", action="store_true", help="用实时价刷新 segments")
    ap.add_argument("--add", nargs="+", metavar="起-止", help="查询并补入新分段")
    ap.add_argument("--set", nargs="+", metavar="起-止=价格",
                    help="人工录入票价（无直达车等场景），打 manual 标记、预算时复核")
    ap.add_argument("--note", help="--set 的备注（如中转方案构成）")
    ap.add_argument("--json", action="store_true", help="单段查询以 JSON 输出")
    args = ap.parse_args()

    if args.check or args.update:
        cmd_check_update(args, update=args.update)
    elif args.set:
        cmd_set(args)
    elif args.add:
        cmd_add(args)
    elif len(args.stations) == 2:
        day = date.fromisoformat(args.date) if args.date else None
        r = FareClient().max_fare_range(args.stations[0], args.stations[1],
                                        start=day, days=args.days)
        if args.json:
            print(json.dumps(r, ensure_ascii=False))
        else:
            skip = f"（{r['days_skipped']}天超预售期跳过）" if r["days_skipped"] else ""
            win = f"，窗口 {r['window']} 共查{r['days_queried']}天" if r["window"] else ""
            print(f"{args.stations[0]}→{args.stations[1]} 二等座最高价: {r['price']} 元 "
                  f"[{r['train']} {r['from_station']}→{r['to_station']} "
                  f"最高日 {r['query_date']}, 当日共{r['train_count']}车次有二等座]{win}{skip}")
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
