import httpx, painel
r = httpx.post(painel.URL, headers={"x-agente-token": painel.TOKEN},
               json={"acao": "iniciar_campanha", "nome": "Teste", "arquivo_origem": "teste.xlsx", "total_empresas": 0})
print("URL:", painel.URL)
print("Status:", r.status_code)
print("Tipo:", r.headers.get("content-type"))
print("Resposta:", r.text[:500])