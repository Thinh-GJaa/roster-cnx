from ortools.sat.python import cp_model
import pandas as pd
import random
from copy import deepcopy
from datetime import date, timedelta
import calendar

# --- Giữ tính ngẫu nhiên nhưng tái lập được ---
random.seed(42)  # đổi 42 để thay đổi chuỗi ngẫu nhiên

# ==== Excel engine: ưu tiên xlsxwriter để có highlight; nếu thiếu thì openpyxl (không highlight) ====
try:
    import xlsxwriter  # noqa: F401
    _EXCEL_ENGINE = "xlsxwriter"
    _HAVE_HIGHLIGHT = True
except Exception:
    _EXCEL_ENGINE = "openpyxl"
    _HAVE_HIGHLIGHT = False


# ==== Tính các khung cuối tuần của một tháng năm 2025 ====
def compute_month_weekend_slots(year: int, month: int):
    """
    Trả về list 'slots' (theo thứ tự thời gian). Mỗi slot là dict:
      {
        'has_sat': bool,
        'has_sun': bool,
        'date_sat': date | None,
        'date_sun': date | None
      }
    Gộp theo tuần ISO (Mon..Sun). Nếu tháng bắt đầu bằng CN => slot đầu chỉ có CN.
    Nếu tháng kết thúc bằng T7 => slot cuối chỉ có T7.
    """
    first_weekday, days_in_month = calendar.monthrange(year, month)
    # Liệt kê toàn bộ ngày trong tháng
    all_days = [date(year, month, d) for d in range(1, days_in_month + 1)]

    # Gom theo tuần ISO
    by_week = {}
    for d in all_days:
        iso_week = d.isocalendar().week
        if iso_week not in by_week:
            by_week[iso_week] = {"has_sat": False, "has_sun": False, "date_sat": None, "date_sun": None}
        if d.weekday() == 5:  # Saturday
            by_week[iso_week]["has_sat"] = True
            by_week[iso_week]["date_sat"] = d
        elif d.weekday() == 6:  # Sunday
            by_week[iso_week]["has_sun"] = True
            by_week[iso_week]["date_sun"] = d

    # Sắp xếp theo tuần ISO (trong cùng năm)
    week_keys = sorted(by_week.keys())
    slots = [by_week[k] for k in week_keys if by_week[k]["has_sat"] or by_week[k]["has_sun"]]
    return slots


def solve_once(USERS, slots):
    """
    Giải bài toán 1 lần với danh sách USERS và danh sách slots đã cho.
    Trả về (status, df_long) với status thuộc {cp_model.OPTIMAL, cp_model.FEASIBLE} nếu có nghiệm, ngược lại trả (status, None).
    df_long dùng để in ra console; các hàm save sẽ tự tạo dạng wide với tên cột là ngày thực.
    """
    W = len(slots)
    weeks = list(range(1, W + 1))
    SITES = sorted({u["site"] for u in USERS})

    # Generate sets from USERS
    PEOPLE = [u["name"] for u in USERS]
    site_of = {u["name"]: u["site"] for u in USERS}
    cannot_sun = {u["name"] for u in USERS if not u["can_sun"]}
    must_sun = {u["name"] for u in USERS if u["must_sun"]}
    worked_last_month = {u["name"] for u in USERS if u["worked_last_month"]}
    cannot_work_alone = {u["name"] for u in USERS if not u["can_alone"]}

    # Site có 1 người
    single_site = {s: (sum(1 for u in USERS if u["site"] == s) == 1) for s in SITES}

    # Target days theo số slot của tháng (giữ như quy tắc cũ)
    target_days = 2 if W == 4 else 3

    model = cp_model.CpModel()
    Sun = {(p, w): model.NewBoolVar(f"Sun[{p},{w}]") for p in PEOPLE for w in weeks}
    Sat = {(p, w): model.NewBoolVar(f"Sat[{p},{w}]") for p in PEOPLE for w in weeks}

    # =========================
    #         CONSTRAINTS
    # =========================
    for w in weeks:
        has_sat = slots[w-1]["has_sat"]
        has_sun = slots[w-1]["has_sun"]

        # Chủ nhật: nếu có CN trong tháng thì bắt buộc 2 người và >=1 must_sun
        if has_sun:
            model.Add(sum(Sun[(p, w)] for p in PEOPLE) == 2)
            if len(must_sun) > 0:
                model.Add(sum(Sun[(p, w)] for p in must_sun) >= 1)
        else:
            # Không có CN ở slot này → tất cả Sun=0
            for p in PEOPLE:
                model.Add(Sun[(p, w)] == 0)

        # Cấm CN cho những ai không thể trực CN (chỉ có ý nghĩa nếu has_sun)
        for p in cannot_sun:
            model.Add(Sun[(p, w)] == 0)

        # Không thể vừa Sat vừa Sun trong cùng slot
        for p in PEOPLE:
            model.Add(Sat[(p, w)] + Sun[(p, w)] <= 1)

        # Thứ bảy (chỉ áp nếu tháng có T7 ở slot này)
        if has_sat:
            for s in SITES:
                members = [p for p in PEOPLE if site_of[p] == s]
                if single_site[s]:
                    continue
                model.Add(sum(Sat[(p, w)] for p in members) >= 1)

            for p in cannot_work_alone:
                model.Add(sum(Sat[(q, w)] for q in PEOPLE if site_of[q] == site_of[p]) >= 2)
        else:
            # Không có T7 → tất cả Sat=0
            for p in PEOPLE:
                model.Add(Sat[(p, w)] == 0)

        # Mỗi tuần phải có ít nhất 1 người từ OH hoặc TV trực (Sat hoặc Sun) nếu slot đó có ngày làm
        site_oh = [p for p in PEOPLE if site_of[p] == "OH"]
        site_tv = [p for p in PEOPLE if site_of[p] == "TV"]
        # Nếu slot có T7 hoặc CN thì áp dụng
        if has_sat or has_sun:
            model.Add(sum(Sat[(p, w)] + Sun[(p, w)] for p in site_oh) +
                      sum(Sat[(p, w)] + Sun[(p, w)] for p in site_tv) >= 1)

    # Tuần 1: ai đã trực tháng trước thì không trực slot đầu tiên của tháng
    for p in worked_last_month:
        model.Add(Sat[(p, 1)] + Sun[(p, 1)] == 0)

    # ==== Cách A: must_sun chỉ trực CN và đủ target_days CN ====
    for p in must_sun:
        model.Add(sum(Sun[(p, w)] for w in weeks) == target_days)
        for w in weeks:
            model.Add(Sat[(p, w)] == 0)

    # Những người không thuộc must_sun: tối đa 1 lần trực CN trong tháng
    for p in PEOPLE:
        if p not in must_sun:
            model.Add(sum(Sun[(p, w)] for w in weeks) <= 1)

    # Tổng ngày trực (Sat+Sun) của mỗi người đúng target_days
    for p in PEOPLE:
        model.Add(sum(Sat[(p, w)] + Sun[(p, w)] for w in weeks) == target_days)

    # Chu kỳ phân bố (như cũ)
    if W == 4:
        for p in PEOPLE:
            for i in range(W - 1):
                w1, w2 = weeks[i], weeks[i + 1]
                model.Add(Sat[(p, w1)] + Sun[(p, w1)] + Sat[(p, w2)] + Sun[(p, w2)] <= 1)
            model.Add(Sat[(p, 1)] + Sun[(p, 1)] == Sat[(p, 3)] + Sun[(p, 3)])
            model.Add(Sat[(p, 2)] + Sun[(p, 2)] == Sat[(p, 4)] + Sun[(p, 4)])
    elif W >= 5:
        for p in PEOPLE:
            for i in range(W - 2):
                w1, w2, w3 = weeks[i], weeks[i + 1], weeks[i + 2]
                model.Add(
                    Sat[(p, w1)] + Sun[(p, w1)] +
                    Sat[(p, w2)] + Sun[(p, w2)] +
                    Sat[(p, w3)] + Sun[(p, w3)] <= 2
                )

    # =========================
    #         SOLVE
    # =========================
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return status, None

    # =========================
    #        OUTPUT (long)
    # =========================
    # Tạo map: w -> nhãn hiển thị (Sat/Sun)
    labels_sat = []
    labels_sun = []
    for w in weeks:
        s = slots[w-1]
        labels_sat.append(s["date_sat"].isoformat() + "_T7" if s["has_sat"] else None)
        labels_sun.append(s["date_sun"].isoformat() + "_CN" if s["has_sun"] else None)

    data = []
    for p in PEOPLE:
        record = {"Site": site_of[p], "Name": p}
        sat_count = 0
        sun_count = 0
        for w in weeks:
            sat = solver.Value(Sat[(p, w)])
            sun = solver.Value(Sun[(p, w)])
            sat_count += sat
            sun_count += sun
            # Ghi nhãn theo ngày thực để in ra (ưu tiên hiển thị dễ đọc)
            tag = []
            if sat and labels_sat[w-1]:
                tag.append(labels_sat[w-1])
            if sun and labels_sun[w-1]:
                tag.append(labels_sun[w-1])
            record[f"Slot{w}"] = ", ".join(tag) if tag else ""
        record["SatCount"] = sat_count
        record["SunCount"] = sun_count
        record["Total"] = sat_count + sun_count
        data.append(record)

    df = pd.DataFrame(data).sort_values(["Site", "Name"])
    return status, df, labels_sat, labels_sun


def to_wide_for_export(df_long, labels_sat, labels_sun):
    """
    Chuyển df_long sang dạng wide: mỗi ngày có một cột theo thứ tự thời gian, ví dụ:
      2025-05-03_T7, 2025-05-04_CN, ...
    Giá trị 1 = trực, 0 = nghỉ.
    """
    # Xây danh sách cột ngày theo thứ tự: từng slot -> (T7 nếu có) -> (CN nếu có)
    day_cols = []
    for ls, ln in zip(labels_sat, labels_sun):
        if ls: day_cols.append(ls)
        if ln: day_cols.append(ln)

    rows = []
    for _, r in df_long.iterrows():
        out = {"Site": r["Site"], "Name": r["Name"]}
        # init 0
        for c in day_cols:
            out[c] = 0
        # parse các đánh dấu trong từng Slot
        for key, val in r.items():
            if not key.startswith("Slot"):
                continue
            if not val:
                continue
            for token in val.split(", "):
                if token in out:
                    out[token] = 1
        out["SatCount"] = int(r["SatCount"])
        out["SunCount"] = int(r["SunCount"])
        out["Total"] = int(r["Total"])
        rows.append(out)

    df_wide = pd.DataFrame(rows).sort_values(["Site", "Name"])
    cols = ["Site", "Name"] + day_cols + ["SatCount", "SunCount", "Total"]
    return df_wide[cols], day_cols


def print_schedule(df, labels_sat, labels_sun, note=""):
    if note:
        print(note)
    print("\n=== Detailed Schedule by Slot (ngày thực) ===")
    print("Mỗi 'Slot' là cuối tuần theo tuần ISO trong tháng; cột hiển thị các ngày thực bạn trực.")
    current_site = None
    for _, row in df.iterrows():
        if row['Site'] != current_site:
            current_site = row['Site']
            print(f"\n────────────── {current_site} ──────────────")
        # Thu gọn dòng hiển thị
        slots = [row.get(f"Slot{i+1}", "") for i in range(sum(1 for _ in labels_sat))]
        print(f"{row['Name']:<10} | " + " | ".join(s or "-" for s in slots) +
              f" | Sat:{row['SatCount']:<2} | Sun:{row['SunCount']:<2} | Total:{row['Total']}")


def save_csv_and_xlsx(df_long, labels_sat, labels_sun, csv_name="schedule.csv", xlsx_name="schedule.xlsx"):
    """
    Lưu CSV và Excel với cột theo ngày thực (YYYY-MM-DD_T7 / YYYY-MM-DD_CN).
    Excel có conditional formatting tô màu ô = 1 nếu có xlsxwriter.
    """
    df_wide, day_cols = to_wide_for_export(df_long, labels_sat, labels_sun)

    # CSV
    df_wide.to_csv(csv_name, index=False)
    print(f"Đã lưu CSV: {csv_name}")

    # Excel
    with pd.ExcelWriter(xlsx_name, engine=_EXCEL_ENGINE) as writer:
        df_wide.to_excel(writer, sheet_name="Schedule", index=False)
        if _HAVE_HIGHLIGHT and len(day_cols) > 0:
            wb = writer.book
            ws = writer.sheets["Schedule"]
            fmt_on = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
            n_rows = len(df_wide)
            # Vị trí cột day_cols trong df_wide: sau 2 cột đầu (Site, Name)
            start_col = 2
            end_col = start_col + len(day_cols) - 1
            ws.conditional_format(1, start_col, n_rows, end_col, {
                "type": "cell",
                "criteria": "==",
                "value": 1,
                "format": fmt_on
            })
            print(f"Đã lưu Excel (có highlight): {xlsx_name}")
        else:
            msg = "không highlight (openpyxl)" if not _HAVE_HIGHLIGHT else "không có cột ngày"
            print(f"Đã lưu Excel ({msg}): {xlsx_name}")


def build_and_solve():
    # ---- INPUT: tháng của năm 2025 ----
    try:
        month = int(input("Nhập tháng của năm 2025 (1-12): ").strip())
        if month < 1 or month > 12:
            print("Giá trị không hợp lệ, mặc định dùng tháng 1.")
            month = 1
    except Exception:
        month = 1

    year = 2025
    slots = compute_month_weekend_slots(year, month)
    if not slots:
        print("Tháng không có ngày T7/CN trong nội bộ tháng (rất hiếm).")
        return

    # ---- USER DATA CONFIG ----
    USERS = [
        {"name": "Cong",    "site": "QTSC1", "can_sun": True,  "can_alone": True,  "must_sun": False, "worked_last_month": False},
        {"name": "Bao",     "site": "QTSC1", "can_sun": True,  "can_alone": True,  "must_sun": False, "worked_last_month": True},
        {"name": "Huy",     "site": "QTSC1", "can_sun": True,  "can_alone": True,  "must_sun": False, "worked_last_month": False},
        {"name": "Thang",   "site": "QTSC9", "can_sun": True,  "can_alone": True,  "must_sun": False, "worked_last_month": True},
        {"name": "Lam",     "site": "QTSC9", "can_sun": True,  "can_alone": True,  "must_sun": False, "worked_last_month": False},
        {"name": "Liem",    "site": "QTSC9", "can_sun": True,  "can_alone": True,  "must_sun": False, "worked_last_month": False},
        {"name": "Thinh",   "site": "QTSC9", "can_sun": False, "can_alone": False, "must_sun": False, "worked_last_month": True},
        {"name": "Mi",      "site": "QTSC9", "can_sun": False, "can_alone": False, "must_sun": False, "worked_last_month": False},
        {"name": "KietDinh","site": "QTSC9", "can_sun": True,  "can_alone": True,  "must_sun": True,  "worked_last_month": False},
        {"name": "Hoa",     "site": "FLE",   "can_sun": True,  "can_alone": True,  "must_sun": False, "worked_last_month": True},
        {"name": "Khai",    "site": "FLE",   "can_sun": True,  "can_alone": True,  "must_sun": False, "worked_last_month": False},
        {"name": "Duong",   "site": "FLE",   "can_sun": True,  "can_alone": True,  "must_sun": True,  "worked_last_month": True},
        {"name": "Hoang",   "site": "TV",    "can_sun": False, "can_alone": True,  "must_sun": False, "worked_last_month": True},
        {"name": "KietLat", "site": "OH",    "can_sun": False, "can_alone": True,  "must_sun": False, "worked_last_month": False},
    ]

    # 1) Thử giải với cấu hình gốc
    status, df, labels_sat, labels_sun = solve_once(USERS, slots)
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print(f"Tháng {month:02d}/{year}: có {len([s for s in slots if s['has_sat']])} ngày T7 và {len([s for s in slots if s['has_sun']])} ngày CN.")
        print_schedule(df, labels_sat, labels_sun)
        csv_name  = f"schedule_{year}-{month:02d}.csv"
        xlsx_name = f"schedule_{year}-{month:02d}.xlsx"
        save_csv_and_xlsx(df, labels_sat, labels_sun, csv_name, xlsx_name)
        return

    # 2) Nếu thất bại: random thứ tự các ứng viên worked_last_month=True, thử nới dần
    idx_candidates = [i for i, u in enumerate(USERS) if u["worked_last_month"]]
    if not idx_candidates:
        print("Không có kết quả và không có ai đang 'worked_last_month=True' để nới ràng buộc.")
        return

    random.shuffle(idx_candidates)  # theo seed ở đầu file
    for k in range(1, len(idx_candidates) + 1):
        USERS_try = deepcopy(USERS)
        relaxed_idxs = idx_candidates[:k]
        for i in relaxed_idxs:
            USERS_try[i]["worked_last_month"] = False

        names_relaxed = [USERS[i]["name"] for i in relaxed_idxs]
        status, df, labels_sat, labels_sun = solve_once(USERS_try, slots)
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            note = "ĐÃ NỚI RÀNG BUỘC: tạm thời bỏ 'worked_last_month=True' cho: " + ", ".join(names_relaxed)
            print(f"Tháng {month:02d}/{year}: có {len([s for s in slots if s['has_sat']])} ngày T7 và {len([s for s in slots if s['has_sun']])} ngày CN.")
            print_schedule(df, labels_sat, labels_sun, note=note)
            csv_name  = f"schedule_{year}-{month:02d}.csv"
            xlsx_name = f"schedule_{year}-{month:02d}.xlsx"
            save_csv_and_xlsx(df, labels_sat, labels_sun, csv_name, xlsx_name)
            return

    # 3) Nếu vẫn không tìm được nghiệm
    print("Không có kết quả: đã thử bỏ dần 'worked_last_month=True' theo thứ tự ngẫu nhiên nhưng vẫn vô nghiệm.")


if __name__ == '__main__':
    build_and_solve()
