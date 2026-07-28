# -*- coding: utf-8 -*-
"""
uploads/ 文件夹 → docs/data.json 构建脚本（拖文件即更新方案）
读取：
  uploads/ 下最新的 学生答题记录*.xlsx（当次+全量两个sheet，自动过滤数学）
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

def parse_one(f):
    """解析一个答题记录xlsx → (recs, fl, 文件内容最大日期)"""
    cur = pd.read_excel(f, sheet_name="当次统计(按答题时间)")
    if "学科" in cur.columns:
        cur = cur[cur["学科"].astype(str).str.contains("数学")]
    study_cols = [c for c in ["前测日期","新知学日期1","新知学日期2","巩固学日期1","巩固学日期2","巩固学日期3"] if c in cur.columns]
    recs = []
    for _, r in cur.iterrows():
        stu, mod = str(r["学生名称"]).strip(), str(r["题模名称"]).strip()
        if not stu or not mod or stu == "nan":
            continue
        study = sorted({d for d in (nd(r.get(c)) for c in study_cols) if d})
        ans_all, ans_ok = [], []
        for i in range(1, 8):
            d = nd(r.get(f"答题时间{i}"))
            if not d: continue
            ans_all.append(d)
            ce = str(r.get(f"消错题结果{i}", "")).strip()
            if not ce or ce in ("/", "nan"):
                ans_ok.append(d)                      # 非消错作答（二遍学判定）
        dt = max(ans_all) if ans_all else (max(study) if study else None)
        recs.append({"n": stu, "m": mod,
                     "p": 1 if str(r["是否过关"]).strip() == "过关" else 0,
                     "dt": dt, "pre": str(r.get("前测是否答对", "/")).strip() or "/",
                     "acts": sorted(set(study) | set(ans_ok))})
    full = pd.read_excel(f, sheet_name="全量统计(不限时间)")
    if "学科" in full.columns:
        full = full[full["学科"].astype(str).str.contains("数学")]
    fl = {}
    for stu, g in full.groupby("学生名称"):
        fl[str(stu).strip()] = sorted({str(x).strip() for x in g["题模名称"]})
    mx = max((r["dt"] for r in recs if r["dt"]), default="0000")
    return recs, fl, mx

def build_records():
    xs = glob.glob(os.path.join(UP, "*学生答题记录*.xlsx"))
    assert xs, "uploads/ 下没有 学生答题记录*.xlsx"
    parsed = []
    for f in xs:
        try:
            parsed.append((f, *parse_one(f)))
        except Exception as e:
            print(f"跳过 {os.path.basename(f)}: {e}")
    # 按文件内容最大日期从旧到新，逐文件按学生整体替换（每个学生以其最新一次导出为准）
    parsed.sort(key=lambda x: x[3])
    by_stu, fl_by_stu = {}, {}
    for f, recs, fl, mx in parsed:
        print(f"合并 {os.path.basename(f)}（截至{mx}）")
        stus = {r["n"] for r in recs}
        for s in stus:
            by_stu[s] = [r for r in recs if r["n"] == s]
        for s, mods in fl.items():
            fl_by_stu[s] = mods
    all_recs = [r for rs in by_stu.values() for r in rs]
    return all_recs, fl_by_stu

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
    recs, fl = build_records()
    plans, plans_fall = build_plans()
    dts = sorted(r["dt"] for r in recs if r["dt"])
    out = {"records": recs, "fullLearned": fl,
           "plans": plans, "plansFall": plans_fall,
           "exportStart": dts[0] if dts else None,
           "generatedAt": time.strftime("%Y-%m-%d %H:%M:%S")}
    os.makedirs("docs", exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"OK → {OUT}: 记录{len(recs)}条({out['exportStart']}→{dts[-1] if dts else '-'})，"
          f"暑期规划{len(plans)}人，秋季规划{len(plans_fall)}人")

if __name__ == "__main__":
    main()
