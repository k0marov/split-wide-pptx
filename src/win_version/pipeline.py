from src.win_version import win_prepare
from src.win_version import process


def pipeline_split_wide_pptx(input_path: str, output_path: str, title_font_size_threshold: float):
    triplicated_path = input_path + '.triplicated'
    # puts triplicated slides into the same path
    mapping = win_prepare.classify_and_triplicate(input_path, triplicated_path, title_font_size_threshold)
    process.split_slides_into_thirds(triplicated_path, mapping, output_path)
