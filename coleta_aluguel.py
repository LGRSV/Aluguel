"""
Coleta de imóveis para alugar no Facebook Marketplace (Palmas/TO), sem IA.

Reaproveita o robô de ../marketplace-iphone13/coleta.py (guias anônimas, ritmo lento, repescagem)
trocando só as buscas e o extrator, e grava em dados/ desta pasta.

Uso:
    python coleta_aluguel.py teste       # 2 buscas, 8 anúncios em dados/teste
    python coleta_aluguel.py coletar     # coleta, histórico, gerar_aluguel.py e publica no GitHub
    python coleta_aluguel.py coletar --guias 2
"""

import argparse
from datetime import date
import asyncio
import os
import subprocess
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(AQUI), "marketplace-iphone13"))
import coleta  # noqa: E402
import historico  # noqa: E402
from publicar import publicar  # noqa: E402

coleta.DADOS = os.path.join(AQUI, "dados")
coleta.LOG = os.path.join(coleta.DADOS, "coleta.log")

# A categoria Locações por faixa de preço (cada faixa devolve outros ~20 anúncios) + buscas por texto.
FAIXAS = [(0, 600), (600, 900), (900, 1200), (1200, 1600), (1600, 2200), (2200, 3200), (3200, 100000)]
BUSCAS = ["/propertyrentals"] + [f"/propertyrentals?minPrice={a}&maxPrice={b}" for a, b in FAIXAS] + [
    "aluguel casa", "aluguel apartamento", "kitnet aluguel", "casa para alugar", "apartamento para alugar",
    "aluga-se", "quarto para alugar", "sala comercial aluguel",
]

# Página de locação: "<título>" / "R$1.900/mês" / "Locações" / "<cidade>" / ... "Detalhes da unidade" / "Descrição".
coleta.EXTRATOR = r"""
async () => {
  await new Promise(r => setTimeout(r, 2500));
  [...document.querySelectorAll('div[role="button"],span')]
    .filter(e => /^Ver mais$/i.test(e.innerText.trim()) && !e.closest("a")).forEach(b => b.click());  // "Ver mais" que é link leva para outra página
  await new Promise(r => setTimeout(r, 800));
  const id = (location.pathname.match(/item\/(\d+)/) || [])[1];
  const L = document.body.innerText.split('\n').map(s => s.trim()).filter(Boolean);
  const limpa = a => a.map(s => s.replace(/\s*Ver (mais|menos)$/i, '')).filter(Boolean).join('\n');
  const fotos = [...new Set([...document.querySelectorAll('img')]
    .filter(im => /^Foto (de produto )?de /i.test(im.alt) && /scontent|fbcdn/.test(im.src)).map(im => im.src))];
  // Formato 1, locação: "<título>" / "R$1.900/mês" / "Locações" / "<cidade>" ... "Detalhes da unidade" ... "Descrição"
  const ip = L.findIndex(s => /^R\$\s?[\d.]+\s*\/\s*m[êe]s/i.test(s));
  if (ip > 0) {
    const ilo = L.findIndex((s, k) => k > ip && s === 'Localização do imóvel');
    let local = L.slice(ip + 1, ip + 5).find(s => /, [A-Z]{2}$/.test(s)) || (ilo > -1 ? L[ilo + 1] : '');
    const idu = L.findIndex((s, k) => k > ip && s === 'Detalhes da unidade');
    const detalhes = idu > -1 ? L.slice(idu + 1, ilo > idu ? ilo : idu + 8).join(' | ') : '';
    const ides = L.findIndex((s, k) => k > ip && s === 'Descrição');
    const fim = L.findIndex((s, k) => k > ides && /^Denuncie esse classificado|^Enviar mensagem$|^Seleções de hoje$/.test(s));
    return {id, formato: 'locacao', titulo: L[ip - 1], preco: (L[ip].match(/R\$\s?([\d.]+)/) || [])[1] || '',
            local, condicao: '', detalhes,
            descricao: ides > -1 ? limpa(L.slice(ides + 1, fim > ides ? fim : ides + 40)) : '', fotos_url: fotos};
  }
  // Formato 2, anúncio comum: igual ao de iPhone ("Anunciado ... em <cidade>", preço e título logo antes)
  const ia = L.findIndex(s => /^Anunciado\s+(?:h[áa]\s+.+?\s+)?em\s+.+$/.test(s));
  if (ia < 2) return {id, erro: 'indisponivel'};
  let p2 = ia - 1;
  while (p2 > 0 && !/^R\$\s?[\d.]+/.test(L[p2])) p2--;
  const ic = L.findIndex((s, k) => k > ia && s === 'Condição');
  const f2 = L.findIndex((s, k) => k > ia && /A localização é aproximada|^Informações do vendedor$|^Pesquisas relacionadas$/.test(s));
  const i2 = ic > -1 ? ic + 2 : L.findIndex((s, k) => k > ia && s === 'Detalhes') + 1;
  return {id, formato: 'comum', titulo: L[p2 - 1] || '', preco: (L[p2].match(/R\$\s?([\d.]+)/) || [])[1] || '',
          local: L[ia].replace(/^Anunciado\s+(?:h[áa]\s+.+?\s+)?em\s+/, '').trim(), condicao: '', detalhes: '',
          descricao: limpa(L.slice(i2, f2 > i2 ? f2 : i2 + 40).filter(s => !/^(Enviar mensagem|Detalhes|Salvar|Compartilhar)$/i.test(s))),
          fotos_url: fotos};
}
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("modo", choices=["teste", "coletar"])
    ap.add_argument("--guias", type=int, default=2)
    ap.add_argument("--max", type=int, default=450)
    a = ap.parse_args()
    coleta.GUIAS = max(1, a.guias)
    os.makedirs(coleta.DADOS, exist_ok=True)

    if a.modo == "teste":
        coleta.PAUSA_BUSCA, coleta.PAUSA_ANUNCIO, coleta.PAUSA_LONGA, coleta.POR_VEZ = (15, 25), (5, 9), (10, 20), 2
        coleta.DADOS = os.path.join(coleta.DADOS, "teste")
        return asyncio.run(coleta.coletar(["/propertyrentals", "kitnet aluguel"], 8))

    rc = asyncio.run(coleta.coletar(BUSCAS, a.max))
    if rc == 0:
        coleta.log("Atualizando o histórico e conferindo os anúncios que sumiram ...")
        historico.registrar_coleta(AQUI)
        historico.verificar(AQUI)
        coleta.log("Classificando com o Haiku e gerando planilha e página ...")
        if subprocess.run([sys.executable, os.path.join(AQUI, "gerar_aluguel.py")]).returncode == 0:
            publicar(AQUI, f"Coleta de {date.today():%d/%m/%Y}")
    return rc


if __name__ == "__main__":
    sys.exit(main())
