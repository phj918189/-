import csv, pathlib

from . import db

BASE = pathlib.Path(__file__).resolve().parents[1]

def _load_rules():
    rules=[]
    p = BASE / "sql" / "item_rules.csv"
    with p.open(encoding="utf-8") as f:
        for r in csv.DictReader(f):
            rules.append(r)
    return rules

def _load_people():
    people=[]
    p = BASE / "sql" / "researchers.csv"
    with p.open(encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r.get("active","1") == "1":
                people.append({"name": r["name"], "email": r["email"]})
    return people

def run_assign(df):
    rules = _load_rules()
    people = _load_people()
    loads = db.today_loads()  # {name: count}

    # 이름 전용 리스트
    names = [p["name"] for p in people]

    assigned=[]
    for _, row in df.iterrows():
        item = str(row["item"])
        chosen = None

        # 1) 규칙 우선
        for rule in sorted(rules, key=lambda x: int(x.get("priority","1"))):
            if rule["item_pattern"] and rule["item_pattern"] in item:
                pref = rule.get("preferred")
                if pref in names:
                    chosen = pref
                    break

        # 2) 라운드로빈(오늘자 작업량 최소)
        if not chosen:
            pool = sorted([(n, loads.get(n,0)) for n in names], key=lambda x: x[1])
            chosen = pool[0][0]

        assigned.append({"sample_no": row["sample_no"], "item": item, "researcher": chosen, "method":"rule+rr"})
        loads[chosen] = loads.get(chosen,0) + 1

    db.save_assignments(assigned)
    return assigned
