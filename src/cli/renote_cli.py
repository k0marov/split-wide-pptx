import argparse
import sys

from src.renote.processor import process_pptx, ProcessingOptions


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Renote PPTX transformer")
    parser.add_argument("input", help="Путь к исходному .pptx")
    parser.add_argument("output", help="Путь для сохранения результата .pptx")
    # Direct mode is the only mode now
    parser.add_argument("--title-min-font", type=float, default=36.0, help="Минимальный размер шрифта титула, pt")
    parser.add_argument("--title-min-width-ratio", type=float, default=1.2, help="Мин. ширина shape к ширине трети")
    return parser


def main(argv: list[str]) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    opts = ProcessingOptions(
        title_min_font_pt=args.title_min_font,
        title_min_width_ratio=args.title_min_width_ratio,
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


