"""Monta o index.html do Painel de Licenças injetando os arquivos da marca.

Uso: python3 src/build.py src <pasta de saída>
"""
import base64
import sys
from pathlib import Path
from urllib.parse import quote

src = Path(sys.argv[1])
saida = Path(sys.argv[2])
marca = src / "marca"


def svg_uri(nome):
    texto = (marca / nome).read_text(encoding="utf-8")
    return "data:image/svg+xml," + quote(" ".join(texto.split()), safe=" =:/',.-")


def bin_uri(nome, tipo):
    return f"data:{tipo};base64," + base64.b64encode((marca / nome).read_bytes()).decode()


html = (src / "index.src.html").read_text(encoding="utf-8")
trocas = {
    "{{LOGO}}": svg_uri("logo-branco.svg"),
    "{{FAVICON}}": svg_uri("favicon.svg"),
    "{{CHEVRON}}": svg_uri("pattern-chevron.svg"),
    "{{ISO}}": bin_uri("pattern-isometrico.webp", "image/webp"),
}
for chave, valor in trocas.items():
    if chave not in html:
        sys.exit(f"Marcador {chave} não encontrado")
    html = html.replace(chave, valor)
if "{{" in html:
    sys.exit("Sobrou marcador sem substituir")

saida.mkdir(parents=True, exist_ok=True)
(saida / "index.html").write_text(html, encoding="utf-8")
print(f"index.html gerado ({len(html) // 1024} KB)")
