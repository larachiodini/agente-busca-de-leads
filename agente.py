"""
Agente de Leads – Simplifica+

Como usar (no terminal, com o .venv ativado):
    python agente.py "distribuidoras de alimentos" lista.xlsx

Aceita o export do Driva (.xlsx com várias abas) ou do Speedio (.csv/.xlsx).
Para testar com poucas empresas primeiro:
    python agente.py "distribuidoras de alimentos" lista.xlsx --limite 5
"""

import argparse
import asyncio
import json
import re
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

from claude_agent_sdk import (
    ClaudeAgentOptions,
    ClaudeSDKClient,
    ResultMessage,
    create_sdk_mcp_server,
    tool,
)

import ferramentas as f
import painel

PASTA = Path(__file__).parent
INSTRUCOES = (PASTA / "instrucoes_agente.md").read_text(encoding="utf-8")
MINIMO_EMPRESAS = 60

# Telefones que as ferramentas REALMENTE encontraram para a empresa atual.
# O agente só pode aprovar telefones desta lista (evita números inventados).
telefones_vistos = {}


def registrar(telefones, fonte):
    for telefone in telefones:
        telefones_vistos.setdefault(telefone, fonte)


def resposta_ferramenta(dados):
    return {"content": [{"type": "text", "text": json.dumps(dados, ensure_ascii=False)}]}


def dominio(url):
    return urlparse(url if url.startswith("http") else "https://" + url).netloc.replace("www.", "")


# ----------------------------------------------------------------------
# Ferramentas que o agente pode usar
# ----------------------------------------------------------------------
@tool("consultar_cnpj", "Consulta dados e telefones oficiais na Receita Federal.", {"cnpj": str})
async def ferramenta_cnpj(args):
    dados = await asyncio.to_thread(f.consultar_cnpj, args["cnpj"])
    registrar(dados.get("telefones_fixos", []), "Receita Federal")
    return resposta_ferramenta(dados)


@tool("buscar_google", "Pesquisa no Google. Use para achar site, CNPJ Biz, redes sociais e contatos.",
      {"consulta": str})
async def ferramenta_google(args):
    dados = await asyncio.to_thread(f.buscar_google, args["consulta"])
    for r in dados.get("resultados", []):
        registrar(r["telefones_no_resumo"], f"Google ({dominio(r['link'] or '')})")
    if dados.get("painel_google"):
        registrar(dados["painel_google"]["telefones"], "Google (perfil da empresa)")
    return resposta_ferramenta(dados)


@tool("ler_site", "Abre uma página (site da empresa, página de contato, CNPJ Biz) e extrai telefones.",
      {"url": str})
async def ferramenta_site(args):
    dados = await asyncio.to_thread(f.ler_site, args["url"])
    registrar(dados.get("telefones_fixos", []), dominio(dados.get("url", args["url"])))
    return resposta_ferramenta(dados)


SERVIDOR = create_sdk_mcp_server(
    name="leads", version="1.0.0",
    tools=[ferramenta_cnpj, ferramenta_google, ferramenta_site],
)

OPCOES = ClaudeAgentOptions(
    system_prompt=INSTRUCOES,
    model="sonnet",
    mcp_servers={"leads": SERVIDOR},
    allowed_tools=["mcp__leads__consultar_cnpj", "mcp__leads__buscar_google", "mcp__leads__ler_site"],
    disallowed_tools=["Bash", "Write", "Edit", "Read", "Glob", "Grep", "WebFetch",
                      "WebSearch", "NotebookEdit", "Task", "TodoWrite"],
    max_turns=15,
    max_budget_usd=0.50,  # trava de segurança: no máximo US$ 0,50 por empresa
)

PEDIDO = """Campanha: {campanha}

Enriqueça esta empresa:
- Nome: {nome}
- CNPJ: {cnpj}
- Nome fantasia: {nome_fantasia}
- Cidade: {cidade}
- Atividade: {atividade}
- Setor/Subsetor: {setor} / {subsetor}
- Faturamento (estimado): {faturamento}
- Telefones já conhecidos ({origem}): {telefone}
- Site informado: {site}
- Sites possíveis (podem incluir sites errados, confirme qual é o oficial): {sites_candidatos}
- LinkedIn: {linkedin}
- Facebook: {facebook}
- Instagram: {instagram}

Passos:
1. consultar_cnpj para confirmar os dados oficiais e a situação.
2. Ache e confirme o site oficial. Se há sites possíveis, confira qual é o da empresa
   (lendo a página e comparando nome/CNPJ). Se não veio nenhum, busque no Google.
3. Leia o site e as páginas de contato.
4. Busque no Google: CNPJ Biz, LinkedIn, Instagram, Facebook e "telefone" da empresa.
5. Pare quando já tiver olhado as fontes principais (no máximo umas 10 consultas).

No final, responda SOMENTE com um JSON neste formato:
{{"dentro_do_perfil": true, "motivo_perfil": "indústria de ...",
  "telefones": [{{"numero": "(19) 3577-9705", "fonte": "site oficial"}}],
  "site_confirmado": "empresa.com.br", "observacao": ""}}
"""


def ler_json(texto):
    achado = re.search(r"\{.*\}", texto or "", re.DOTALL)
    try:
        return json.loads(achado.group(0)) if achado else {}
    except json.JSONDecodeError:
        return {}


# ----------------------------------------------------------------------
# Enriquecer uma empresa
# ----------------------------------------------------------------------
async def enriquecer(empresa, campanha):
    telefones_vistos.clear()
    campos = ("nome", "nome_fantasia", "cnpj", "cidade", "atividade", "setor", "subsetor",
              "faturamento", "origem", "telefone", "site", "sites_candidatos",
              "linkedin", "facebook", "instagram")
    dados = {}
    for campo in campos:
        valor = empresa.get(campo)
        if isinstance(valor, list):
            valor = ", ".join(valor)
        dados[campo] = valor or "não informado"
    pedido = PEDIDO.format(campanha=campanha, **dados)
    resposta, custo = "", 0.0
    async with ClaudeSDKClient(options=OPCOES) as cliente:
        await cliente.query(pedido)
        async for mensagem in cliente.receive_response():
            if isinstance(mensagem, ResultMessage):
                resposta = mensagem.result or ""
                custo = mensagem.total_cost_usd or 0.0

    resultado = ler_json(resposta)

    # Junta os telefones: Speedio + Receita + os aprovados pelo agente que foram REALMENTE vistos
    telefones, fontes = [], []

    def adicionar(numero, fonte):
        numero = f.padronizar_telefone_fixo(numero or "")
        if numero and numero not in telefones:
            telefones.append(numero)
            fontes.append(f"{numero} – {fonte}")

    for numero in f.extrair_telefones_fixos(str(empresa.get("telefone") or "")):
        adicionar(numero, empresa.get("origem", "lista de origem"))
    for numero, fonte in telefones_vistos.items():
        if fonte == "Receita Federal":
            adicionar(numero, fonte)
    for item in resultado.get("telefones", []):
        numero = f.padronizar_telefone_fixo(item.get("numero", ""))
        if numero in telefones_vistos:
            adicionar(numero, item.get("fonte") or telefones_vistos[numero])

    empresa["telefones_encontrados"] = telefones
    empresa["fontes"] = fontes or ["não encontrado"]
    if resultado.get("site_confirmado"):
        empresa["site"] = resultado["site_confirmado"]
    empresa["dentro_do_perfil"] = resultado.get("dentro_do_perfil", True)
    empresa["observacao"] = resultado.get("observacao") or resultado.get("motivo_perfil", "")
    return custo


# ----------------------------------------------------------------------
# Programa principal
# ----------------------------------------------------------------------
async def principal():
    parser = argparse.ArgumentParser(description="Agente de Leads")
    parser.add_argument("campanha", help='Ex.: "distribuidoras de alimentos"')
    parser.add_argument("arquivo", help="Lista exportada do Driva ou Speedio (.csv ou .xlsx)")
    parser.add_argument("--limite", type=int, default=0, help="Processar só N empresas (teste)")
    args = parser.parse_args()

    empresas = f.carregar_lista(args.arquivo)
    if args.limite:
        empresas = empresas[: args.limite]
    print(f"\n{len(empresas)} empresas carregadas para a campanha '{args.campanha}'.")
    if len(empresas) < MINIMO_EMPRESAS and not args.limite:
        print(f"ATENÇÃO: a lista tem menos de {MINIMO_EMPRESAS} empresas. Exporte mais na plataforma.")

    nome_arquivo = re.sub(r"\W+", "-", args.campanha.lower()).strip("-")
    saida = PASTA / f"leads_{nome_arquivo}_{date.today()}.xlsx"
    custo_total = 0.0

    campanha_id = painel.iniciar_campanha(args.campanha, args.arquivo, len(empresas))
    if campanha_id:
        print("Painel do Lovable conectado: acompanhe o progresso por lá.")
    elif painel.ATIVO:
        print("Aviso: painel configurado, mas não respondeu. Seguindo só com a planilha.")

    for numero, empresa in enumerate(empresas, start=1):
        print(f"\n[{numero}/{len(empresas)}] {empresa.get('nome')} ({empresa.get('cnpj')})")
        try:
            custo_total += await enriquecer(empresa, args.campanha)
            print(f"   Telefones: {'; '.join(empresa['telefones_encontrados']) or 'nenhum'}")
        except Exception as erro:
            empresa["fontes"] = [f"erro: {erro}"]
            print(f"   ERRO: {erro}")
        f.salvar_planilha(empresas, saida)  # salva a cada empresa: nada se perde se travar
        painel.salvar_lead(campanha_id, empresa, numero, custo_total)

    painel.finalizar_campanha(campanha_id, custo_total)

    no_perfil = [e for e in empresas if e.get("dentro_do_perfil", True)]
    com_telefone = [e for e in empresas if e.get("telefones_encontrados")]
    print("\n========== RESUMO ==========")
    print(f"Empresas processadas: {len(empresas)}")
    print(f"Dentro do perfil: {len(no_perfil)}")
    print(f"Com pelo menos 1 telefone: {len(com_telefone)}")
    fora = [e.get("nome") for e in empresas if not e.get("dentro_do_perfil", True)]
    if fora:
        print("Possivelmente fora do perfil (confira):", "; ".join(fora))
    print(f"Custo aproximado da IA: US$ {custo_total:.2f}")
    print(f"Planilha: {saida}")


if __name__ == "__main__":
    asyncio.run(principal())
