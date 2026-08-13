# -*- coding: utf-8 -*-
"""
uploads/ 文件夹 → docs/data.json 构建脚本（拖文件即更新方案）
读取：
  uploads/ 下全部的 学生答题记录*.xlsx（兼容单张“全量学习数据表”和旧版双sheet，自动过滤数学）
  uploads/ 下全部的 学生课次明细*.csv（暑期→plans，秋季→plansFall；同名学生取文件名日期最新的）
输出：docs/data.json（records/fullLearned/plans/plansFall/exportStart/generatedAt）
依赖：pip install pandas openpyxl
"""
import os, re, json, glob, time
import warnings; warnings.filterwarnings("ignore")
import pandas as pd

UP, OUT = "uploads", os.path.join("docs", "data.json")

def nd(v):
    m = re.match(r"(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})", str(v))
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else None

def read_full_sheet(f):
    """读取答题记录中的全量数据表；优先表名含“全量”，并用必需列校验。"""
    required = {"学生名称", "题模名称", "是否过关"}
    book = pd.ExcelFile(f)
    preferred = [s for s in book.sheet_names if "全量" in s]
    candidates = preferred + [s for s in book.sheet_names if s not in preferred]
    seen = []
    for sheet in candidates:
        data = pd.read_excel(book, sheet_name=sheet)
        data.columns = [str(c).lstrip("\ufeff").strip() for c in data.columns]
        seen.append(f"{sheet}({len(data)}行)")
        if required.issubset(data.columns):
            return data, sheet
    raise ValueError("没有找到学习数据表（需要列：学生名称 / 题模名称 / 是否过关）；已检查：" + "、".join(seen))

def parse_one(f):
    """解析一个答题记录xlsx → (全量recs, fl, 文件内容最大日期, 窗口起始日期)"""
    # 课次起点由网页端按学生设置，因此这里直接读取全量学习数据，不再用导出窗口裁剪历史学习。
    full, sheet = read_full_sheet(f)
    print(f"读取 {os.path.basename(f)} / {sheet}")
    cur = full.drop_duplicates(subset=["学生名称","题模名称"], keep="last")
    if "学科" in cur.columns:
        cur = cur[cur["学科"].astype(str).str.contains("数学")]
    new_cols = [c for c in ["新知学日期1","新知学日期2"] if c in cur.columns]
    reinforce_cols = [c for c in ["巩固学日期1","巩固学日期2","巩固学日期3"] if c in cur.columns]
    recs = []
    for _, r in cur.iterrows():
        stu, mod = str(r["学生名称"]).strip(), str(r["题模名称"]).strip()
        if not stu or not mod or stu == "nan":
            continue
        new_days = {d for d in (nd(r.get(c)) for c in new_cols) if d}
        reinforce_days = {d for d in (nd(r.get(c)) for c in reinforce_cols) if d}
        explicit_study = new_days | reinforce_days
        ans_all, ans_ok = [], []
        for i in range(1, 8):
            d = nd(r.get(f"答题时间{i}"))
            if not d: continue
            ans_all.append(d)
            ce = str(r.get(f"消错题结果{i}", "")).strip()
            if not ce or ce in ("/", "nan"):
                ans_ok.append(d)                      # 非消错作答（二遍学判定）
        # 前测日期：答对的计入（判会时点）；答错的仅当当天另有学习活动或再无其他日期时计入
        base = explicit_study | set(ans_ok)
        pre_d = nd(r.get("前测日期"))
        pre_ok = str(r.get("前测是否答对", "")).strip() == "答对"
        test_days = set()
        if pre_d and (pre_ok or pre_d in base or not (base or ans_all)):
            base.add(pre_d)
            test_days.add(pre_d)
        # 用于区分一遍学习的性质：新知学=预习，巩固学/前测=提升。
        # 非前测日的普通作答只作为有效活动，不擅自判断预习或提升。
        normal_days = explicit_study | {d for d in ans_ok if d != pre_d}
        acts = sorted(base)
        dt = max(ans_all) if ans_all else (max(acts) if acts else None)
        recs.append({"n": stu, "m": mod,
                     "p": 1 if str(r["是否过关"]).strip() == "过关" else 0,
                     "dt": dt, "pre": str(r.get("前测是否答对", "/")).strip() or "/",
                     "acts": acts, "studyActs": sorted(normal_days),
                     "newActs": sorted(new_days), "reinforceActs": sorted(reinforce_days),
                     "testActs": sorted(test_days)})
    if "学科" in full.columns:
        full = full[full["学科"].astype(str).str.contains("数学")]
    fl = {}
    for stu, g in full.groupby("学生名称"):
        fl[str(stu).strip()] = sorted({str(x).strip() for x in g["题模名称"]})
    mx = max((r["dt"] for r in recs if r["dt"]), default="0000")
    all_days = sorted({d for r in recs for d in (r["acts"] or ([r["dt"]] if r["dt"] else []))})
    win = all_days[0] if all_days else "0000"
    return recs, fl, mx, win

def build_records():
    xs = glob.glob(os.path.join(UP, "*学生答题记录*.xlsx"))
    assert xs, "uploads/ 下没有 学生答题记录*.xlsx"
    parsed = []
    for f in xs:
        try:
            parsed.append((f, *parse_one(f)))
        except Exception as e:
            print(f"跳过 {os.path.basename(f)}: {e}")
    # 按文件内容最大日期从旧到新，按（学生,题模）合并：
    # 窗口起点之前的活动保留旧文件的，窗口内以新文件为准（支持任意窗口的导出）
    parsed.sort(key=lambda x: (x[3], x[4]))   # 先按截至日期，再按窗口起点（窗口晚的裁剪导出最后合并）
    by_key, fl_by_stu = {}, {}
    for f, recs, fl, mx, win in parsed:
        print(f"合并 {os.path.basename(f)}（{win}→{mx}）")
        for r in recs:
            key = (r["n"], r["m"])
            if key in by_key:
                old_r = by_key[key]
                keep = [d for d in (old_r["acts"] or ([old_r["dt"]] if old_r["dt"] else [])) if d < win]
                r["acts"] = sorted(set(keep) | set(r["acts"]))
                for field in ("studyActs", "newActs", "reinforceActs", "testActs"):
                    old_days = [d for d in old_r.get(field, []) if d < win]
                    r[field] = sorted(set(old_days) | set(r.get(field, [])))
                if old_r["dt"] and (not r["dt"] or old_r["dt"] > r["dt"]):
                    r["dt"] = old_r["dt"]
                if r["pre"] in ("", "/") and old_r.get("pre") not in ("", "/", None):
                    r["pre"] = old_r["pre"]
            by_key[key] = r
        for s, mods in fl.items():
            fl_by_stu[s] = sorted(set(fl_by_stu.get(s, [])) | set(mods))
    # 保留合并后的全部活动。每位学生的“第1次课日期”由网页端手动设置，
    # 该日期只决定课次编号起点；更早活动仍作为历史已学数据参与进度判断。
    return list(by_key.values()), fl_by_stu

def build_plans():
    plans, plans_fall, stamp = {}, {}, {}
    def fdate(p):
        m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", os.path.basename(p))
        return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else "0000"
    for path in sorted(glob.glob(os.path.join(UP, "*学生课次明细*.csv")), key=fdate):
        try:
            d = pd.read_csv(path)
        except Exception:
            d = pd.read_csv(path, encoding="gbk")
        for stage_key, store in [("暑期", plans), ("秋季", plans_fall)]:
            s = d[d["阶段"].astype(str).str.contains(stage_key)].copy()
            if not len(s): continue
            s["no"] = s["课次"].astype(str).str.extract(r"(\d+)").astype(int)
            s["sub"] = s["阶段"].astype(str).str.contains("期末前").astype(int)
            for stu, g in s.groupby("学生姓名"):
                stu = str(stu).strip()
                key = (stu, id(store))
                if stamp.get(key, "") > fdate(path):    # 同学生取文件名日期最新
                    continue
                stamp[key] = fdate(path)
                lessons, dur, types, renum = [], {}, {}, 0
                for (sub, no, date), gg in sorted(g.groupby(["sub", "no", "上课日期"])):
                    renum += 1
                    lessons.append({"no": renum, "date": nd(date),
                                    "names": [str(x).strip() for x in gg["题模名称"]],
                                    "stage": ("期末前" if sub else "期中前") if stage_key == "秋季" else "暑期"})
                    for _, r in gg.iterrows():
                        n = str(r["题模名称"]).strip()
                        dur[n] = dur.get(n, 0) + float(r["学习时长(min·系数后)"])
                        types[n] = str(r["类型"]).strip()
                first = g.iloc[0]
                meta = {"score": int(first["目标分"]) if pd.notna(first.get("目标分")) else None,
                        "band": str(first.get("分数段")), "grade": str(first.get("年级"))}
                store[stu] = {"lessons": lessons, "dur": dur, "types": types, "meta": meta}
    return plans, plans_fall

def main():
    plans, plans_fall = build_plans()
    recs, fl = build_records()
    dts = sorted(r["dt"] for r in recs if r["dt"])
    out = {"records": recs, "fullLearned": fl,
           "plans": plans, "plansFall": plans_fall,
           "exportStart": dts[0] if dts else None,
           "generatedAt": time.strftime("%Y-%m-%d %H:%M:%S")}
    os.makedirs("docs", exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    # file:// 本地预览无法稳定 fetch data.json，因此同步更新 HTML 内嵌数据。
    index_path = os.path.join("docs", "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            html = f.read()
        old_match = re.search(r'<script id="EMBED_DATA" type="application/json">(.*?)</script>', html, re.S)
        old_payload = json.loads(old_match.group(1)) if old_match else {}
        embedded_out = dict(out)
        # 知识树只存于单文件网页中，更新学习数据时必须原样保留。
        if old_payload.get("tree"):
            embedded_out["tree"] = old_payload["tree"]
        payload = json.dumps(embedded_out, ensure_ascii=False, separators=(",", ":"))
        html, n = re.subn(r'(<script id="EMBED_DATA" type="application/json">).*?(</script>)',
                          lambda m: m.group(1) + payload + m.group(2), html,
                          count=1, flags=re.S)
        if n != 1:
            raise ValueError("docs/index.html 中未找到唯一的 EMBED_DATA")
        with open(index_path, "w", encoding="utf-8") as f:
            f.write(html)
    print(f"OK → {OUT}: 记录{len(recs)}条({out['exportStart']}→{dts[-1] if dts else '-'})，"
          f"暑期规划{len(plans)}人，秋季规划{len(plans_fall)}人")

if __name__ == "__main__":
    main()
