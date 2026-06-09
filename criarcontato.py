import requests

TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb21wYW55X2lkIjo2NDk5MSwidXNlcl9pZCI6MTEzMjJ9.V6lojDviSyMH9mTl_NPCVdc8PghEWenFspueXJmZGVg"
HEADERS = {
    "accept": "application/json",
    "Content-Type": "application/json",
    "Authorization": f"Bearer {TOKEN}"
}

r = requests.post(
    "https://amigobot-api.amigoapp.com.br/patients",
    headers=HEADERS,
    json={
        "name": "Teste time JBN 2",
        "born": "1990-05-15",
        "gender": "Masculino",
        "contact_cellphone": "81988887777",
        "email": "teste2@amigo.com.br",
        "cpf": "52998224725",
        "address_city": "Recife",
        "address_state": "Pernambuco"
    },
    timeout=15
)

print(f"Status: {r.status_code}")
print(f"Response: {r.json()}")