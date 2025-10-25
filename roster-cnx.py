from ortools.sat.python import cp_model
import pandas as pd

def build_and_solve():
    # ---- INPUTS ----
    W = 5
    weeks = list(range(1, W + 1))
    SITES = ["QTSC1", "QTSC9", "FLE", "TV", "OH"]
    PEOPLE = [
        "Cong", "Bao", "Huy",
        "Thang", "Lam", "Liem", "Thinh", "Mi", "KietDinh",
        "Hoa", "Khai", "Duong",
        "Hoang", "KietLat",
    ]
    site_of = {
        "Cong": "QTSC1", "Bao": "QTSC1", "Huy": "QTSC1",
        "Thang": "QTSC9", "Lam": "QTSC9", "Liem": "QTSC9", "Thinh": "QTSC9", "Mi": "QTSC9", "KietDinh": "QTSC9",
        "Hoa": "FLE", "Khai": "FLE", "Duong": "FLE",
        "Hoang": "TV", "KietLat": "OH",
    }
    single_site = {"QTSC1": False, "QTSC9": False, "FLE": False, "TV": True, "OH": True}

    cannot_sun = {"Thinh", "Mi", "Hoang", "KietLat"}
    must_sun = {"Duong", "KietDinh"}

    model = cp_model.CpModel()
    target_days = 2 if W == 4 else 3

    Sun = {(p, w): model.NewBoolVar(f"Sun[{p},{w}]") for p in PEOPLE for w in weeks}
    Sat = {(p, w): model.NewBoolVar(f"Sat[{p},{w}]") for p in PEOPLE for w in weeks}

    # Constraints
    for w in weeks:
        model.Add(sum(Sun[(p, w)] for p in PEOPLE) == 2)
        model.Add(sum(Sun[(p, w)] for p in must_sun) == 1)
        eligible_other = [p for p in PEOPLE if (p not in must_sun) and (p not in cannot_sun)]
        model.Add(sum(Sun[(p, w)] for p in eligible_other) == 1)

        for p in cannot_sun:
            model.Add(Sun[(p, w)] == 0)

        for p in PEOPLE:
            model.Add(Sat[(p, w)] + Sun[(p, w)] <= 1)

        for s in SITES:
            members = [p for p in PEOPLE if site_of[p] == s]
            if len(members) == 1 and single_site[s]:
                continue
            model.Add(sum(Sat[(p, w)] for p in members) >= 1)

        # Ensure Thinh and Mi never alone in their site on Sat
        for p in ["Thinh", "Mi"]:
            model.Add(sum(Sat[(q, w)] for q in PEOPLE if site_of[q] == site_of[p]) >= 2)

        # New rule: if OH off then TV must work, and vice versa
        site_oh = [p for p in PEOPLE if site_of[p] == "OH"]
        site_tv = [p for p in PEOPLE if site_of[p] == "TV"]
        model.Add(sum(Sat[(p, w)] + Sun[(p, w)] for p in site_oh) +
                  sum(Sat[(p, w)] + Sun[(p, w)] for p in site_tv) >= 1)

    m = max(1, len(must_sun))
    base = len(weeks) // m
    extra = len(weeks) % m
    for i, p in enumerate(sorted(must_sun)):
        model.Add(sum(Sun[(p, w)] for w in weeks) == base + (1 if i < extra else 0))

    for p in PEOPLE:
        if p not in must_sun:
            model.Add(sum(Sun[(p, w)] for w in weeks) <= 1)

    for p in PEOPLE:
        model.Add(sum(Sat[(p, w)] + Sun[(p, w)] for w in weeks) == target_days)

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

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 30
    solver.parameters.num_search_workers = 8
    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print('No feasible solution with given constraints.')
        return

    # Build unified table
    data = []
    for p in PEOPLE:
        record = {"Site": site_of[p], "Name": p}
        total = 0
        for w in weeks:
            sat = solver.Value(Sat[(p, w)])
            sun = solver.Value(Sun[(p, w)])
            if sat and sun:
                record[f"Week{w}"] = 'Sat+Sun'
                total += 2
            elif sat:
                record[f"Week{w}"] = 'Sat'
                total += 1
            elif sun:
                record[f"Week{w}"] = 'Sun'
                total += 1
            else:
                record[f"Week{w}"] = ''
        record["Total"] = total
        data.append(record)

    df = pd.DataFrame(data)
    df = df.sort_values(["Site", "Name"])

    # Print combined table with site separators
    print("\n=== Detailed Weekly Schedule (All Sites) ===")
    current_site = None
    for idx, row in df.iterrows():
        if row['Site'] != current_site:
            current_site = row['Site']
            print(f"\n────────────── {current_site} ──────────────")
        print(f"{row['Name']:<10}", end=' | ')
        for w in weeks:
            print(f"{row[f'Week{w}'] or '-':<7}", end=' | ')
        print(f" {row['Total']}")

if __name__ == '__main__':
    build_and_solve()