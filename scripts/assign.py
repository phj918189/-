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
    
    # 이미 배정된 작업 조회
    existing_assignments = db.get_existing_assignments()

    # 이름 전용 리스트
    names = [p["name"] for p in people]
    
    # 디버깅: 로드된 데이터 확인
    print(f"DEBUG: 로드된 연구원들: {names}")
    print(f"DEBUG: 로드된 규칙 수: {len(rules)}")
    for i, rule in enumerate(rules):
        print(f"DEBUG: 규칙 {i+1}: {rule}")

    assigned=[]
    skipped=0
    for _, row in df.iterrows():
        sample_no = str(row.get("sample_no", ""))
        item = str(row["item"])
        
        # 중복 체크: 이미 배정된 작업은 스킵
        assignment_key = f"{sample_no}_{item}"
        if assignment_key in existing_assignments:
            print(f"DEBUG: 이미 배정된 작업 스킵 - sample_no: {sample_no}, item: {item}")
            skipped += 1
            continue
        
        chosen = None
        print(f"DEBUG: 처리 중인 항목: '{item}'")

        # 규칙에 따른 배정 (라운드로빈 제거)
        for rule in sorted(rules, key=lambda x: int(x.get("priority","1"))):
            if rule["item_pattern"]:
                # 파이프 구분자로 분리된 패턴들 확인
                patterns = rule["item_pattern"].split("|")
                print(f"DEBUG: 패턴들: {patterns}")
                for pattern in patterns:
                    if pattern.strip() in item:
                        pref = rule.get("preferred")
                        print(f"DEBUG: 매칭된 패턴: '{pattern.strip()}' -> 선호 연구원: '{pref}'")
                        if pref in names:
                            chosen = pref
                            print(f"DEBUG: 배정됨: '{chosen}'")
                            break
                if chosen:
                    break

        # 규칙에 해당하지 않는 경우 첫 번째 연구원에게 배정
        if not chosen:
            chosen = names[0] if names else "system"
            print(f"DEBUG: 규칙 없음, 기본 배정: '{chosen}'")

        assigned.append({"sample_no": row["sample_no"], "item": item, "researcher": chosen, "method":"rule_only"})
        loads[chosen] = loads.get(chosen,0) + 1
    
    print(f"DEBUG: 총 {skipped}개 작업 스킵 (이미 배정됨)")
    print(f"DEBUG: 새로 배정된 작업: {len(assigned)}개")

    db.save_assignments(assigned)
    return assigned
