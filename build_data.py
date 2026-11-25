"""엑셀 4개 -> data/*.json 변환 (한 번만 실행). 사용: python build_data.py"""
import json, glob, os
from openpyxl import load_workbook
LV = {"대분류": "통합대분류", "중분류": "통합중분류", "소분류": "통합소분류", "기본부문": "기본부문"}
def num(x): return float(x) if isinstance(x, (int, float)) else 0.0
for key, tag in LV.items():
    f = glob.glob(f"*{tag}.xlsx") + glob.glob(f"../**/*{tag}.xlsx", recursive=True) + glob.glob(f"/mnt/user-data/uploads/*{tag}.xlsx")
    wb = load_workbook(f[0], read_only=True, data_only=True)
    S = lambda n: list(wb[n].iter_rows(values_only=True))
    tr = S("A표_총거래표(생산자)")
    n = next(i for i, r in enumerate(tr[6:]) if str(r[0]) == "9590")
    hc = {str(c): i for i, c in enumerate(tr[4]) if c is not None}
    rc = {str(r[0]): r for r in tr[6 + n:] if r[0] is not None}
    d = tr[6:6 + n]
    col = lambda c: [num(r[hc[c]]) for r in d]
    cols = lambda cs: [sum(x) for x in zip(*[col(c) for c in cs])]
    sec = lambda c: [num(x) for x in rc[c][2:2 + n]]
    mat = lambda name, k=5: [[round(num(x), k) for x in r[2:2 + n]] for r in S(name)[6:6 + n]]
    csum = lambda name, k=5: [round(sum(c), k) for c in zip(*[[num(x) for x in r[2:2 + n]] for r in S(name)[6:6 + n]])]
    out = {
        "names": [str(r[1]).strip() for r in d], "codes": [str(r[0]) for r in d],
        "Z": [[round(num(x)) for x in r[2:2 + n]] for r in d],
        "out": col("9310"), "exp": col("9140"), "cons": cols(["9111", "9112"]),
        "inv": cols(["9121", "9122", "9131", "9132"]), "imp": col("9329"),
        "cost": {"mid": sec("9590"), "comp": sec("9610"), "op": sec("9621"), "dep": sec("9622"), "tax": sec("9630"), "va": sec("9690")},
        "Ld": mat("생산유발계수"), "impind": csum("수입유발계수"), "vaind": csum("부가가치유발계수"), "amsum": csum("수입투입계수(Am)"),
    }
    json.dump(out, open("data/"+{"대분류":"large","중분류":"mid","소분류":"small","기본부문":"basic"}[key]+".json", "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    print(key, n, 1 // 1024, "KB")
