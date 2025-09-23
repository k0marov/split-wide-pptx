from __future__ import annotations

from typing import Iterable, Optional, List
from io import BytesIO

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
from pptx.oxml.ns import qn

from .detectors import is_title_slide, pick_primary_title_shape


def _delete_shape(shape) -> None:
    element = shape._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _iter_shapes(slide):
    for shape in list(slide.shapes):
        yield shape


def _remove_placeholders(slide) -> None:
    """Remove placeholders from a slide to avoid duplicates on cloning."""
    for shp in list(slide.shapes):
        if getattr(shp, "is_placeholder", False):
            _delete_shape(shp)


def _fix_relationships_for_element(new_el, src_slide, dst_slide) -> None:
    """Rebind relationship ids for copied XML subtree to destination slide."""
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ATTRS = (f"{{{REL_NS}}}embed", f"{{{REL_NS}}}id")

    old_to_new = {}
    try:
        for old_rid, rel in src_slide.part.rels.items():
            try:
                new_rid = dst_slide.part.relate_to(rel._target, rel.reltype)
                old_to_new[old_rid] = new_rid
            except Exception:
                continue
    except Exception:
        pass

    for el in new_el.iter():
        for attr_name in ATTRS:
            old_rid = el.get(attr_name)
            if not old_rid:
                continue
            new_rid = old_to_new.get(old_rid)
            if new_rid:
                el.set(attr_name, new_rid)


def _copy_background(src_slide, dst_slide) -> None:
    """Copy custom background (cSld.bg) and fix relationships."""
    try:
        src_cSld = src_slide._element.cSld
        dst_cSld = dst_slide._element.cSld
    except Exception:
        return
    src_bg = getattr(src_cSld, "bg", None)
    if src_bg is None:
        return
    try:
        dst_bg = getattr(dst_cSld, "bg", None)
        if dst_bg is not None:
            dst_cSld.remove(dst_bg)
    except Exception:
        pass
    from copy import deepcopy
    new_bg = deepcopy(src_bg)
    _fix_relationships_for_element(new_bg, src_slide, dst_slide)
    try:
        dst_cSld.insert(0, new_bg)
        try:
            dst_slide.follow_master_background = False
        except Exception:
            pass
    except Exception:
        try:
            dst_cSld.append(new_bg)
        except Exception:
            pass


def _delete_slide(prs: Presentation, index: int) -> None:
    """Delete slide by dropping relationship and removing from sldIdLst."""
    sldIdLst = prs.slides._sldIdLst
    sldId = sldIdLst[index]
    try:
        rId = sldId.rId  # type: ignore[attr-defined]
    except Exception:
        rId = sldId.get(qn("r:id"))
    if rId:
        try:
            prs.part.drop_rel(rId)
        except Exception:
            pass
    sldIdLst.remove(sldId)


def transform_title_slide(slide, third_width: int) -> None:
    """Reduce slide to a single cloned title shape preserving styles; adjust geometry only."""
    primary = pick_primary_title_shape(slide)
    if primary is None:
        for shp in slide.shapes:
            if getattr(shp, "has_text_frame", False):
                primary = shp
                break
    if primary is None:
        return

    # Remove all shapes on slide
    for shp in list(slide.shapes):
        _delete_shape(shp)

    # Clone original text shape element to preserve all formatting/themes
    from copy import deepcopy
    new_el = deepcopy(primary._element)
    try:
        slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
    except Exception:
        slide.shapes._spTree.append(new_el)

    # Rebind relationships if any
    try:
        _fix_relationships_for_element(new_el, slide, slide)
    except Exception:
        pass

    # Adjust geometry to fit third width and center vertically
    try:
        shp = slide.shapes[-1]
        shp.width = third_width
        shp.left = 0
        try:
            shp.height = slide.part.presentation.slide_height  # type: ignore[attr-defined]
            shp.top = 0
        except Exception:
            pass
        if getattr(shp, "has_text_frame", False):
            tf = shp.text_frame
            # Keep horizontal alignment as-is if set; otherwise center
            try:
                for p in tf.paragraphs:
                    p.alignment = p.alignment or PP_ALIGN.CENTER
            except Exception:
                for p in tf.paragraphs:
                    p.alignment = PP_ALIGN.CENTER
            try:
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            except Exception:
                pass
    except Exception:
        pass




def split_into_thirds_direct(
    input_pptx: str,
    output_pptx: str,
    *,
    title_min_font_pt: float = 36.0,
    title_min_width_ratio: float = 1.2,
    scenarios: Optional[List[str]] = None,
) -> None:
    """Direct mode preserving styles: duplicate each slide three times in-place then filter per third.

    Мы сохраняем тему, фон и стили, создавая новые слайды в рамках исходной презентации,
    а затем удаляем оригинальные. Экспорт в output_pptx уже с узкой шириной.
    """
    prs = Presentation(input_pptx)

    original_width = int(prs.slide_width)
    third_width = int(original_width // 3)

    # 1) Для каждого исходного слайда добавить 3 слайда-клона с сохранением тем/фона
    original_count = len(prs.slides)
    originals = [prs.slides[i] for i in range(original_count)]

    from copy import deepcopy
    clones_per_original = []
    for src_index, src in enumerate(originals):
        scenario_for_src = None
        if scenarios and src_index < len(scenarios):
            scenario_for_src = (scenarios[src_index] or "").strip().lower()
        clones = []
        num_clones = 1 if scenario_for_src == "title" else 3
        for _ in range(num_clones):
            new_slide = prs.slides.add_slide(src.slide_layout)
            _remove_placeholders(new_slide)
            # copy shapes xml
            for shape in src.shapes:
                try:
                    new_el = deepcopy(shape._element)
                except Exception:
                    continue
                _fix_relationships_for_element(new_el, src, new_slide)
                try:
                    new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
                except Exception:
                    try:
                        new_slide.shapes._spTree.append(new_el)
                    except Exception:
                        continue
            # copy background
            _copy_background(src, new_slide)
            clones.append(new_slide)
        clones_per_original.append((scenario_for_src, clones))

    # 2) Удалить оригинальные слайды (они первые original_count в последовательности)
    for idx in range(original_count - 1, -1, -1):
        _delete_slide(prs, idx)

    # 3) Для каждого трио отфильтровать и сдвинуть элементы под свою треть
    for scenario_for_src, clones in clones_per_original:
        for i, slide in enumerate(clones):
            if i == 0:
                left_bound, right_bound = 0, third_width
            elif i == 1:
                left_bound, right_bound = third_width, 2 * third_width
            else:
                left_bound, right_bound = 2 * third_width, original_width

            is_title = False
            if scenario_for_src in ("title", "split"):
                is_title = scenario_for_src == "title"
            else:
                is_title = is_title_slide(slide, third_width, min_font_pt=title_min_font_pt, min_width_ratio=title_min_width_ratio)

            if is_title:
                # Для титульного у нас теперь только один clone; просто привести его к узкой ширине
                transform_title_slide(slide, third_width)
                continue

            for shape in list(slide.shapes):
                try:
                    shape_left = int(shape.left)
                    shape_width = int(shape.width)
                except Exception:
                    _delete_shape(shape)
                    continue
                if shape_left > right_bound or (shape_left + shape_width) < left_bound:
                    _delete_shape(shape)
                    continue
                if i == 1:
                    shape.left -= third_width
                elif i == 2:
                    shape.left -= 2 * third_width

    # 4) Установить новую ширину слайдов и сохранить
    prs.slide_width = third_width
    prs.save(output_pptx)


