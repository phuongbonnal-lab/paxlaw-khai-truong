#!/usr/bin/env python3
"""Tạo link ngắn theo tên khách cho thư mời Paxlaw.

Cách dùng:
    python3 tools/make_links.py danh_sach_khach.csv      (hoặc .xlsx)

File vào cần các cột: "Mã khách", "Họ tên", "Chức danh" (tên cột không phân biệt hoa thường).
Kết quả:
    guests.json   – danh sách khách để trang thư mời tra cứu (đưa lên repo)
    links.csv     – Mã khách, Họ tên, Chức danh, Link  (dán ngược vào Google Sheet)
"""
import csv, json, re, sys, unicodedata
from pathlib import Path

BASE = "https://phuongbonnal-lab.github.io/paxlaw-khai-truong/#"
ROOT = Path(__file__).resolve().parent.parent


HONORIFIC = re.compile(r"^(ông|bà|anh|chị|cô|chú)\s+")


def slugify(name):
    """Giống công thức cột E trong Google Sheet: bỏ danh xưng đầu tên, bỏ dấu, nối bằng gạch ngang."""
    s = unicodedata.normalize("NFC", name.strip()).lower()
    s = HONORIFIC.sub("", s)
    s = s.replace("đ", "d")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "khach"


def read_rows(path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        ws = load_workbook(path, read_only=True, data_only=True).active
        rows = [["" if v is None else str(v) for v in r] for r in ws.iter_rows(values_only=True)]
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
    head = [h.strip().lower() for h in rows[0]]

    def col(*names):
        for n in names:
            if n in head:
                return head.index(n)
        return None

    ci, ni, ti = col("mã khách", "ma khach", "id", "mã"), col("họ tên", "họ và tên", "ho ten", "tên"), col("chức danh", "chức danh · đơn vị", "chuc danh", "đơn vị")
    if ni is None:
        sys.exit('Không thấy cột "Họ tên" trong dòng tiêu đề: %s' % rows[0])
    for r in rows[1:]:
        get = lambda i: (r[i].strip() if i is not None and i < len(r) and r[i] else "")
        if get(ni):
            yield get(ci), get(ni), get(ti)


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    rows = list(read_rows(sys.argv[1]))
    count = {}
    for _, name, _ in rows:
        count[name.strip()] = count.get(name.strip(), 0) + 1
    guests, out = {}, []
    for code, name, title in rows:
        slug = slugify(name)
        if count[name.strip()] > 1 and code:      # trùng tên: thêm mã khách, ví dụ #nguyen-van-anh-k012
            slug += "-" + code.lower()
        if slug in guests:
            sys.exit("Link bị trùng: #%s (%s) — hãy thêm mã khách cho dòng này." % (slug, name))
        guests[slug] = {k: v for k, v in (("c", code), ("n", name), ("t", title)) if v}
        out.append((code, name, title, BASE + slug))
    (ROOT / "guests.json").write_text(json.dumps(guests, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    with open(ROOT / "links.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Mã khách", "Họ tên", "Chức danh", "Link"])
        w.writerows(out)
    print("Đã tạo %d link → guests.json, links.csv" % len(out))


if __name__ == "__main__":
    main()
