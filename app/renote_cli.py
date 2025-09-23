from __future__ import annotations

import argparse
import sys

from renote.processor import process_pptx, ProcessingOptions


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Renote PPTX transformer")
    parser.add_argument("input", help="Путь к исходному .pptx")
    parser.add_argument("output", help="Путь для сохранения результата .pptx")
    # Direct mode is the only mode now
    parser.add_argument("--title-min-font", type=float, default=36.0, help="Минимальный размер шрифта титула, pt")
    parser.add_argument("--title-min-width-ratio", type=float, default=1.2, help="Мин. ширина shape к ширине трети")
    parser.add_argument("--vlm", choices=["heuristic", "ollama"], default="heuristic", help="Режим классификации сценария")
    parser.add_argument("--soffice", default=None, help="Путь к soffice (LibreOffice) для экспорта PNG")
    parser.add_argument("--ollama-model", default="llava:latest", help="Имя модели Ollama (VLM)")
    parser.add_argument("--ollama-url", default="http://localhost:11434", help="URL Ollama API")
    return parser


def main(argv: list[str]) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    opts = ProcessingOptions(
        title_min_font_pt=args.title_min_font,
        title_min_width_ratio=args.title_min_width_ratio,
        vlm_mode=args.vlm,
        soffice_path=args.soffice,
        ollama_model=args.ollama_model,
        ollama_url=args.ollama_url,
    )

    process_pptx(
        args.input,
        args.output,
        options=opts,
        direct=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))


