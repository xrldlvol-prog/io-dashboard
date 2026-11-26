"""엑셀 12개 시트 -> s_<분류>_<번호>.json 변환. 사용: python build_sheets.py (엑셀 4개를 같은 폴더에 두고 실행)"""
import json, glob, os
from openpyxl import load_workbook
LV = {"large": "통합대분류", "mid": "통합중분류", "small": "통합소분류", "basic": "기본부문"}
def cell(x):
    if isinstance(x, bool) or x is None: return None
    if isinstance(x, (int, float)): return int(x) if float(x).is_integer() else round(x, 6)
    return str(x).replace("\n", " ").strip()
names0 = None
for key, tag in LV.items():
    path = (glob.glob(f"*{tag}.xlsx") + glob.glob(f"/mnt/user-data/uploads/*{tag}.xlsx"))[0]
    wb = load_workbook(path, read_only=True, data_only=True)
    if names0 is None: names0 = wb.sheetnames
    assert wb.sheetnames == names0, "시트 구성이 분류마다 다릅니다"
    for k, ws in enumerate(wb.worksheets, 1):
        rows = list(ws.iter_rows(values_only=True))
        hc, hn = rows[4], rows[5]
        nc = max(j for j in range(2, len(hc)) if hc[j] is not None or hn[j] is not None) + 1
        data = [r for r in rows[6:] if r[0] is not None or r[1] is not None]
        out = {"title": str(rows[0][0]).strip(), "cc": [cell(x) for x in hc[2:nc]], "cn": [cell(x) for x in hn[2:nc]],
               "rc": [cell(r[0]) or "" for r in data], "rn": [cell(r[1]) or "" for r in data],
               "v": [[cell(x) for x in (list(r) + [None] * nc)[2:nc]] for r in data]}
        fn = f"s_{key}_{k:02d}.json"
        json.dump(out, open(fn, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
        print(fn, len(data), "x", nc - 2, os.path.getsize(fn) // 1024, "KB", ws.title)
