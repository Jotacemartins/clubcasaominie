import requests
from datetime import datetime, timedelta

TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb21wYW55X2lkIjo2NDk5MSwidXNlcl9pZCI6MTEzMjJ9.V6lojDviSyMH9mTl_NPCVdc8PghEWenFspueXJmZGVg"
HEADERS = {"accept": "application/json", "Authorization": f"Bearer {TOKEN}"}

PATIENT_ID = 110308717

ranges = [
    ("2026-06-01", "2026-06-30"),
    ("2026-07-01", "2026-07-31"),
    ("2026-08-01", "2026-08-31"),
]

print("=" * 70)
print(f"Agendamentos do paciente {PATIENT_ID}")
print("=" * 70)

for start, end in ranges:
    r = requests.get(
        "https://amigobot-api.amigoapp.com.br/attendances",
        headers=HEADERS,
        params={"start_date": start, "end_date": end, "status": "ALL", "place_id": 74999},
        timeout=15
    )
    todos = r.json().get("data", [])
    meus  = [a for a in todos if a.get("patient_id") == PATIENT_ID]

    if not meus:
        continue

    print(f"\n{start} → {end}:")
    for a in meus:
        start_utc = a.get("start_date", "")
        dt_utc    = datetime.fromisoformat(start_utc.replace("Z", ""))
        cancelado = a.get("canceled", False)
        medico    = a.get("user", {}).get("name", "?")
        print(f"  ID: {a.get('id')} | Salvo UTC: {dt_utc.strftime('%d/%m %H:%M')} | Cancelado: {cancelado} | Médico: {medico}")