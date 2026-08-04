import argparse
import csv
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shared_code.dam_normalization import normalize_dam_filename, normalize_dam_relative_path


SAMPLES = [
    "9000602000001-DE-BRANDX-SAMPLE26-GRADE2-ASSESSMENT-1-16.pdf",
    "D2-9002-DEMO-NAT-GRADE2-OPENING.idml",
    "9000601000001-SE-BRAND-EI-DEMO-READING-GRADE2-PAGES-01-32.pdf",
    "QUESTOES E RESOLUCOES - PROVA 1 - 3 ANO.docx",
    "MATERIAIS_DIGITAIS_ALUNOS_PROFESSORES.zip",
    "S26-0-INT83-2A-01-LA-LIC-017a.psd",
    "PROVA ÚNICA - MÓDULO 1 - 2ª SÉRIE.tif",
    "BRANDX-SIMULADOS-PROFESSOR-6-ANO-PROVA-2.ai",
    "LIVRO DO ALUNO - CAPA - VOLUME 1.pdf",
    "LIVRO DO PROFESSOR - MIOLO - FUND 1.indd",
    "PORTUGUÊS MATEMÁTICA CIÊNCIAS HISTÓRIA GEOGRAFIA.docx",
    "ING - CN - CH - EDUCAÇÃO FÍSICA - SIMULADO 26.pdf",
]

PATH_SAMPLES = [
    "Links/PROVA ÚNICA - MÓDULO 1 - 2ª SÉRIE.tif",
    "Document Fonts/HelveticaLTStd-Cond.otf",
    "document-fonts/MyriadPro-BoldIt.otf",
    "09-DIAGRAMAÇÃO/09F-PROVA-FINAL/D2-9002-DEMO-NAT-GRADE2-OPENING.idml",
]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Testa a normalizacao semantica de nomes de arquivo DAM."
    )
    parser.add_argument("--input-json", help="Inventario JSON do Blob para extrair nomes reais.")
    parser.add_argument("--limit", type=int, default=200, help="Limite de nomes do inventario.")
    parser.add_argument("--output-dir", default="tmp/normalizacao", help="Pasta de saida.")
    args = parser.parse_args()

    names = load_names(args.input_json, args.limit) if args.input_json else SAMPLES
    rows = [
        {
            "original": name,
            "normalizado": normalize_dam_filename(name),
            "mudou": "sim" if normalize_dam_filename(name) != name else "nao",
        }
        for name in names
    ]

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_csv = output_dir / "teste_normalizacao_nomes.csv"

    with output_csv.open("w", newline="", encoding="cp1252", errors="replace") as file:
        writer = csv.DictWriter(file, fieldnames=["original", "normalizado", "mudou"], delimiter=";")
        writer.writeheader()
        writer.writerows(rows)

    changed = sum(1 for row in rows if row["mudou"] == "sim")
    print("CONCLUIDO")
    print(f"Nomes analisados: {len(rows)}")
    print(f"Nomes alterados: {changed}")
    print(f"CSV: {output_csv}")
    print()
    for row in rows[:20]:
        print(f"{row['original']} -> {row['normalizado']}")

    print()
    print("Caminhos:")
    for path in PATH_SAMPLES:
        print(f"{path} -> {normalize_dam_relative_path(path)}")


def load_names(input_json: str, limit: int) -> list[str]:
    payload = json.loads(Path(input_json).read_text(encoding="utf-8-sig"))
    names: list[str] = []
    seen: set[str] = set()
    for item in payload.get("items", []):
        if item.get("type") != "blob":
            continue
        name = str(item.get("name") or "").strip()
        if not name or name in seen:
            continue
        seen.add(name)
        names.append(name)
        if len(names) >= limit:
            break
    return names


if __name__ == "__main__":
    main()

