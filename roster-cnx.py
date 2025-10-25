from ortools.sat.python import cp_model
import pandas as pd
import random
from copy import deepcopy

# --- Giữ tính ngẫu nhiên nhưng tái lập được ---
random.seed(42)  # đổi 42 để thay đổi chuỗi ngẫu nhiên

def solve_once(USERS, W):
    """
    Giải bài toán 1 lần với danh sách USERS và W đã cho.
    Trả về (status, df) với status thuộc {cp_model.OPTIMAL, cp_model.FEASIBLE} nếu có nghiệm, ngược lại trả (status, None).
    """
    weeks = list(range(1, W + 1))
    SITES = sorted({u["site"] for u in USERS})

    # Generate sets from USERS
    PEOPLE = [u["name"] for u in USERS]
    site_of = {u["name"]: u["site"] for u in USERS}
    cannot_sun = {u["name"] for u in USERS if not u["can_sun"]}
    must_sun = {u["name"] for u in USERS if u["must_sun"]}
    worked_last_month = {u["name"] for u in USERS if u["worked_last_month"]}
    cannot_work_alone = {u["name"] for u in USERS if not u["can_alone"]}

    single_site = {s: (sum(1 for u in USERS if u["site"] == s) == 1) for s in SITES}

    model = cp_model.CpModel()
    target_days = 2 if W == 4 else 3

    Sun = {(p, w): model.NewBoolVar(f"Sun[{p},{w}]") for p in PEOPLE for w in weeks}
    Sat = {(p, w): model.NewBoolVar(f"Sat[{p},{w}]") for p in PEOPLE for w in weeks}

    # =========================
    #         CONSTRAINTS
    # =========================
    for w in weeks:
        # Mỗi tuần đúng 2 người trực Chủ nhật
        model.Add(sum(Sun[(p, w)] for p in PEOPLE) == 2)

        # Ít nhất 1 người thuộc must_sun trực Chủ nhật mỗi tuần (Cách A)
        if len(must_sun) > 0:
            model.Add(sum(Sun[(p, w)] for p in must_sun) >= 1)

        # Cấm những ai không thể trực CN
        for p in cannot_sun:
            model.Add(Sun[(p, w)] == 0)

        # Không thể vừa Sat vừa Sun trong cùng tuần
        for p in PEOPLE:
            model.Add(Sat[(p, w)] + Sun[(p, w)] <= 1)

        # Mỗi site (nếu có >=2 người) phải có ít nhất 1 người trực T7
        for s in SITES:
            members = [p for p in PEOPLE if site_of[p] == s]
            if single_site[s]:
                continue
            model.Add(sum(Sat[(p, w)] for p in members) >= 1)

        # Với người không thể trực một mình: site đó tối thiểu 2 người trực T7
        for p in cannot_work_alone:
            model.Add(sum(Sat[(q, w)] for q in PEOPLE if site_of[q] == site_of[p]) >= 2)

        # Ít nhất 1 người từ OH hoặc TV trực (Sat hoặc Sun) mỗi tuần
        site_oh = [p for p in PEOPLE if site_of[p] == "OH"]
        site_tv = [p for p in PEOPLE if site_of[p] == "TV"]
        model.Add(sum(Sat[(p, w)] + Sun[(p, w)] for p in site_oh) +
                  sum(Sat[(p, w)] + Sun[(p, w)] for p in site_tv) >= 1)

    # Tuần 1: ai đã trực tháng trước thì không trực tuần 1
    for p in worked_last_month:
        model.Add(Sat[(p, 1)] + Sun[(p, 1)] == 0)

    # ==== Cách A: must_sun chỉ trực CN và đủ target_days CN ====
    for p in must_sun:
        # Tổng số CN đúng bằng target_days
        model.Add(sum(Sun[(p, w)] for w in weeks) == target_days)
        # Cấm trực Thứ bảy
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
    elif W == 5:
        for p in PEOPLE:
            for i in range(W - 2):
                w1, w2, w3 = weeks[i], weeks[i + 1], weeks[i + 2]
                model.Add(
                    Sat[(p, w1)] + Sun[(p, w1)] + Sat[(p, w2)] + Sun[(p, w2)] + Sat[(p, w3)] + Sun[(p, w3)] <= 2
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
    #        OUTPUT DF
    # =========================
    data = []
    for p in PEOPLE:
        record = {"Site": site_of[p], "Name": p}
        sat_count = sum(solver.Value(Sat[(p, w)]) for w in weeks)
        sun_count = sum(solver.Value(Sun[(p, w)]) for w in weeks)
        total = sat_count + sun_count
        for w in weeks:
            sat = solver.Value(Sat[(p, w)])
            sun = solver.Value(Sun[(p, w)])
            if sat and sun:
                record[f"Week{w}"] = 'Sat+Sun'
            elif sat:
                record[f"Week{w}"] = 'Sat'
            elif sun:
                record[f"Week{w}"] = 'Sun'
            else:
                record[f"Week{w}"] = ''
        record["SatCount"] = sat_count
        record["SunCount"] = sun_count
        record["Total"] = total
        data.append(record)

    df = pd.DataFrame(data)
    df = df.sort_values(["Site", "Name"])
    return status, df


def print_schedule(df, W, note=""):
    weeks = list(range(1, W + 1))
    if note:
        print(note)
    print("\n=== Detailed Weekly Schedule (All Sites) ===")
    current_site = None
    for _, row in df.iterrows():
        if row['Site'] != current_site:
            current_site = row['Site']
            print(f"\n────────────── {current_site} ──────────────")
        print(f"{row['Name']:<10}", end=' | ')
        for w in weeks:
            print(f"{row.get(f'Week{w}', '') or '-':<7}", end=' | ')
        print(f" Sat:{row['SatCount']:<2} | Sun:{row['SunCount']:<2} | Total:{row['Total']}")


def save_csv(df, filename="schedule.csv"):
    df.to_csv(filename, index=False)
    print(f"Đã lưu: {filename}")


def build_and_solve():
    # ---- INPUT: số tuần trong tháng ----
    try:
        W = int(input("Nhập số tuần trong tháng (4 hoặc 5): ").strip())
        if W not in [4, 5]:
            print("Giá trị không hợp lệ, mặc định là 4 tuần.")
            W = 4
    except:
        W = 4

    # ---- USER DATA CONFIG ----
    USERS = [
        {"name": "Cong", "site": "QTSC1", "can_sun": True, "can_alone": True, "must_sun": False, "worked_last_month": False},
        {"name": "Bao", "site": "QTSC1", "can_sun": True, "can_alone": True, "must_sun": False, "worked_last_month": True},
        {"name": "Huy", "site": "QTSC1", "can_sun": True, "can_alone": True, "must_sun": False, "worked_last_month": False},
        {"name": "Thang", "site": "QTSC9", "can_sun": True, "can_alone": True, "must_sun": False, "worked_last_month": True},
        {"name": "Lam", "site": "QTSC9", "can_sun": True, "can_alone": True, "must_sun": False, "worked_last_month": False},
        {"name": "Liem", "site": "QTSC9", "can_sun": True, "can_alone": True, "must_sun": False, "worked_last_month": False},
        {"name": "Thinh", "site": "QTSC9", "can_sun": False, "can_alone": False, "must_sun": False, "worked_last_month": True},
        {"name": "Mi", "site": "QTSC9", "can_sun": False, "can_alone": False, "must_sun": False, "worked_last_month": False},
        {"name": "KietDinh", "site": "QTSC9", "can_sun": True, "can_alone": True, "must_sun": True, "worked_last_month": False},
        {"name": "Hoa", "site": "FLE", "can_sun": True, "can_alone": True, "must_sun": False, "worked_last_month": True},
        {"name": "Khai", "site": "FLE", "can_sun": True, "can_alone": True, "must_sun": False, "worked_last_month": False},
        {"name": "Duong", "site": "FLE", "can_sun": True, "can_alone": True, "must_sun": True, "worked_last_month": True},
        {"name": "Hoang", "site": "TV", "can_sun": False, "can_alone": True, "must_sun": False, "worked_last_month": True},
        {"name": "KietLat", "site": "OH", "can_sun": False, "can_alone": True, "must_sun": False, "worked_last_month": False},
    ]

    # 1) Thử giải ngay với cấu hình gốc
    status, df = solve_once(USERS, W)
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print_schedule(df, W)
        save_csv(df, "schedule.csv")
        return

    # 2) Nếu thất bại: random thứ tự các ứng viên worked_last_month=True, thử nới dần (bỏ thêm 1 người mỗi lần)
    idx_candidates = [i for i, u in enumerate(USERS) if u["worked_last_month"]]
    if not idx_candidates:
        print("Không có kết quả và không có ai đang 'worked_last_month=True' để nới ràng buộc.")
        return

    random.shuffle(idx_candidates)  # sẽ theo seed đã đặt ở đầu file
    # Thử tuần tự: lần 1 bỏ 1 người đầu, lần 2 bỏ 2 người đầu, ... (cộng dồn)
    for k in range(1, len(idx_candidates) + 1):
        USERS_try = deepcopy(USERS)
        relaxed_idxs = idx_candidates[:k]
        for i in relaxed_idxs:
            USERS_try[i]["worked_last_month"] = False

        names_relaxed = [USERS[i]["name"] for i in relaxed_idxs]
        status, df = solve_once(USERS_try, W)
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            note = "ĐÃ NỚI RÀNG BUỘC: tạm thời bỏ 'worked_last_month=True' cho các bạn: " + ", ".join(names_relaxed)
            print_schedule(df, W, note=note)
            save_csv(df, "schedule.csv")
            return

    # 3) Nếu vẫn không tìm được nghiệm
    print("Không có kết quả: đã thử bỏ dần 'worked_last_month=True' cho tất cả ứng viên theo thứ tự ngẫu nhiên nhưng vẫn vô nghiệm.")

if __name__ == '__main__':
    build_and_solve()
