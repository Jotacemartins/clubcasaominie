import requests
from datetime import datetime

TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb21wYW55X2lkIjo2NDk5MSwidXNlcl9pZCI6MTEzMjJ9.V6lojDviSyMH9mTl_NPCVdc8PghEWenFspueXJmZGVg"
HEADERS = {"accept": "application/json", "Authorization": f"Bearer {TOKEN}"}

# CPFs de teste
PACIENTES_TESTE = [110308717]  # patient_id do JBN 1 — adiciona o ID do JBN 2 depois de criar

ranges = [
    ("2026-06-01", "2026-06-30"),
    ("2026-07-01", "2026-07-31"),
    ("2026-08-01", "2026-08-31"),
    ("2026-09-01", "2026-09-30"),
]

print("Buscando agendamentos de teste...")
ativos = []

for start, end in ranges:
    r = requests.get(
        "https://amigobot-api.amigoapp.com.br/attendances",
        headers=HEADERS,
        params={"start_date": start, "end_date": end, "status": "ALL", "place_id": 74999},
        timeout=15
    )
    todos = r.json().get("data", [])
    meus  = [a for a in todos if a.get("patient_id") in PACIENTES_TESTE and not a.get("canceled")]
    ativos.extend(meus)

print(f"\nAtivos encontrados: {len(ativos)}")
for a in ativos:
    print(f"  ID: {a.get('id')} | Data: {a.get('start_date','')[:10]} | Médico: {a.get('user',{}).get('name','?')}")

if not ativos:
    print("Nenhum agendamento ativo encontrado.")
else:
    print(f"\nCancelando {len(ativos)} agendamentos...")
    for a in ativos:
        rc = requests.put(
            f"https://amigobot-api.amigoapp.com.br/attendances/cancel/{a.get('id')}",
            headers=HEADERS,
            timeout=15
        )
        status = "✅" if rc.status_code == 200 else "❌"
        print(f"  {status} ID {a.get('id')}: {rc.status_code}")