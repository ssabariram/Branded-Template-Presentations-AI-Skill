"""
<company-name> Element Animations — OOXML animation injection via raw lxml.

python-pptx has no public API for animations; we build the <p:timing> XML directly.

ECMA-376 schema constraint (critical):
    p:childTnLst is a CHILD of p:cTn, not a sibling.
    Every p:par / p:seq node follows: <p:par><p:cTn ...><p:childTnLst>...</p:childTnLst></p:cTn></p:par>

Public API:
    inject_animations(slide, animations_spec, slide_type)
    validate_animations(animations_spec, slide_num)
    ANIMATION_CATALOG
    RECOMMENDED_ANIMATIONS
"""

from lxml import etree

# ─── Namespaces ───────────────────────────────────────────────────────────────

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def _p(local):  return f"{{{NS_P}}}{local}"


# ─── Animation Effect Catalog ─────────────────────────────────────────────────

ANIMATION_CATALOG = {
    # ── Entrance ──────────────────────────────────────────────────────────────
    "appear": {
        "preset_class": "entr", "preset_id": 1, "preset_subtype": 0,
        "filter": "", "transition": "in", "anim_elem": "set",
        "dir_map": {},
        "description": "Element appears instantly with no animation.",
        "best_for": ["title", "body", "subtitle"],
    },
    "fade": {
        "preset_class": "entr", "preset_id": 10, "preset_subtype": 0,
        "filter": "fade", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {},
        "description": "Smooth cross-fade in. Universal, always safe.",
        "best_for": ["title", "body", "subtitle", "chart", "table"],
    },
    "fly_in": {
        "preset_class": "entr", "preset_id": 2, "preset_subtype": 8,
        "filter": "fly(fromBottom)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "from_bottom": (8,  "fly(fromBottom)"),
            "from_top":    (4,  "fly(fromTop)"),
            "from_left":   (2,  "fly(fromLeft)"),
            "from_right":  (1,  "fly(fromRight)"),
        },
        "description": "Element flies in from an edge.",
        "best_for": ["title", "body", "subtitle"],
    },
    "float_in": {
        "preset_class": "entr", "preset_id": 22, "preset_subtype": 8,
        "filter": "fly(fromBottom)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "from_bottom": (8, "fly(fromBottom)"),
            "from_top":    (4, "fly(fromTop)"),
        },
        "description": "Gentle float/drift in.",
        "best_for": ["body", "subtitle"],
    },
    "wipe": {
        "preset_class": "entr", "preset_id": 48, "preset_subtype": 8,
        "filter": "wipe(fromBottom)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "from_bottom": (8, "wipe(fromBottom)"),
            "from_top":    (4, "wipe(fromTop)"),
            "from_left":   (2, "wipe(fromLeft)"),
            "from_right":  (1, "wipe(fromRight)"),
        },
        "description": "Edge wipe reveal.",
        "best_for": ["body", "chart", "table"],
    },
    "zoom": {
        "preset_class": "entr", "preset_id": 9, "preset_subtype": 4,
        "filter": "zoom(in)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "in":              (4,  "zoom(in)"),
            "out":             (20, "zoom(out)"),
            "in_slide_center": (36, "zoom(in)"),
        },
        "description": "Zoom in or out entrance.",
        "best_for": ["title", "kpi", "chart"],
    },
    "split": {
        "preset_class": "entr", "preset_id": 54, "preset_subtype": 10,
        "filter": "barn(inVertical)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "horizontal_in":  (10, "barn(inHorizontal)"),
            "vertical_in":    (10, "barn(inVertical)"),
            "horizontal_out": (5,  "barn(outHorizontal)"),
            "vertical_out":   (5,  "barn(outVertical)"),
        },
        "description": "Barn-door split reveal.",
        "best_for": ["chart", "table", "body"],
    },
    "box": {
        "preset_class": "entr", "preset_id": 55, "preset_subtype": 4,
        "filter": "box(in)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "in":  (4, "box(in)"),
            "out": (1, "box(out)"),
        },
        "description": "Box iris in or out.",
        "best_for": ["chart", "title"],
    },
    "dissolve": {
        "preset_class": "entr", "preset_id": 3, "preset_subtype": 0,
        "filter": "dissolve", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {},
        "description": "Pixel-dissolve fade in.",
        "best_for": ["title", "body", "icon"],
    },
    "swivel": {
        "preset_class": "entr", "preset_id": 15, "preset_subtype": 3,
        "filter": "swivel(vertical)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "horizontal": (1, "swivel(horizontal)"),
            "vertical":   (3, "swivel(vertical)"),
        },
        "description": "Swivel rotate in.",
        "best_for": ["title", "chart"],
    },
    "bounce": {
        "preset_class": "entr", "preset_id": 63, "preset_subtype": 0,
        "filter": "bounce", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {},
        "description": "Bounce into place.",
        "best_for": ["kpi"],
    },
    "wheel": {
        "preset_class": "entr", "preset_id": 62, "preset_subtype": 1,
        "filter": "wheel(1)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "1_spoke": (1, "wheel(1)"),
            "2_spoke": (2, "wheel(2)"),
            "3_spoke": (3, "wheel(3)"),
            "4_spoke": (4, "wheel(4)"),
            "8_spoke": (8, "wheel(8)"),
        },
        "description": "Wheel wipe reveal.",
        "best_for": ["kpi", "chart"],
    },
    "blinds": {
        "preset_class": "entr", "preset_id": 52, "preset_subtype": 1,
        "filter": "blinds(horizontal)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "horizontal": (1, "blinds(horizontal)"),
            "vertical":   (2, "blinds(vertical)"),
        },
        "description": "Venetian-blinds reveal.",
        "best_for": ["table", "body"],
    },
    "checker": {
        "preset_class": "entr", "preset_id": 56, "preset_subtype": 1,
        "filter": "checker(across)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "across": (1, "checker(across)"),
            "down":   (2, "checker(downward)"),
        },
        "description": "Checkerboard reveal.",
        "best_for": ["kpi"],
    },
    "wedge": {
        "preset_class": "entr", "preset_id": 61, "preset_subtype": 0,
        "filter": "wedge", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {},
        "description": "Clock-sweep wedge reveal.",
        "best_for": ["kpi", "chart"],
    },
    "random_bars": {
        "preset_class": "entr", "preset_id": 53, "preset_subtype": 1,
        "filter": "randombar(horizontal)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "horizontal": (1, "randombar(horizontal)"),
            "vertical":   (2, "randombar(vertical)"),
        },
        "description": "Random bar strips reveal.",
        "best_for": ["body"],
    },
    "strips": {
        "preset_class": "entr", "preset_id": 60, "preset_subtype": 9,
        "filter": "strips(rightDown)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "right_down": (9,  "strips(rightDown)"),
            "right_up":   (10, "strips(rightUp)"),
            "left_down":  (5,  "strips(leftDown)"),
            "left_up":    (6,  "strips(leftUp)"),
        },
        "description": "Diagonal strip reveal.",
        "best_for": ["chart", "body"],
    },
    "plus": {
        "preset_class": "entr", "preset_id": 59, "preset_subtype": 4,
        "filter": "plus(out)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "out": (4, "plus(out)"),
            "in":  (1, "plus(in)"),
        },
        "description": "Plus/cross iris.",
        "best_for": ["kpi"],
    },
    "circle": {
        "preset_class": "entr", "preset_id": 57, "preset_subtype": 4,
        "filter": "circle(out)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "out": (4, "circle(out)"),
            "in":  (1, "circle(in)"),
        },
        "description": "Circular iris reveal.",
        "best_for": ["kpi", "chart"],
    },
    "diamond": {
        "preset_class": "entr", "preset_id": 58, "preset_subtype": 4,
        "filter": "diamond(out)", "transition": "in", "anim_elem": "animEffect",
        "dir_map": {
            "out": (4, "diamond(out)"),
            "in":  (1, "diamond(in)"),
        },
        "description": "Diamond iris reveal.",
        "best_for": ["kpi"],
    },
    # ── Exit ──────────────────────────────────────────────────────────────────
    "disappear": {
        "preset_class": "exit", "preset_id": 1, "preset_subtype": 0,
        "filter": "", "transition": "out", "anim_elem": "set",
        "dir_map": {},
        "description": "Element disappears instantly.",
        "best_for": ["title", "body"],
    },
    "fade_out": {
        "preset_class": "exit", "preset_id": 10, "preset_subtype": 0,
        "filter": "fade", "transition": "out", "anim_elem": "animEffect",
        "dir_map": {},
        "description": "Cross-fade out.",
        "best_for": ["title", "body", "chart"],
    },
    "fly_out": {
        "preset_class": "exit", "preset_id": 2, "preset_subtype": 8,
        "filter": "fly(fromBottom)", "transition": "out", "anim_elem": "animEffect",
        "dir_map": {
            "to_bottom": (8, "fly(fromBottom)"),
            "to_top":    (4, "fly(fromTop)"),
            "to_left":   (2, "fly(fromLeft)"),
            "to_right":  (1, "fly(fromRight)"),
        },
        "description": "Element flies off-screen.",
        "best_for": ["body", "chart"],
    },
    "wipe_out": {
        "preset_class": "exit", "preset_id": 48, "preset_subtype": 8,
        "filter": "wipe(fromBottom)", "transition": "out", "anim_elem": "animEffect",
        "dir_map": {
            "to_bottom": (8, "wipe(fromBottom)"),
            "to_top":    (4, "wipe(fromTop)"),
            "to_left":   (2, "wipe(fromLeft)"),
            "to_right":  (1, "wipe(fromRight)"),
        },
        "description": "Wipe out.",
        "best_for": ["body"],
    },
    "zoom_out": {
        "preset_class": "exit", "preset_id": 9, "preset_subtype": 4,
        "filter": "zoom(out)", "transition": "out", "anim_elem": "animEffect",
        "dir_map": {},
        "description": "Zoom out exit.",
        "best_for": ["title", "kpi"],
    },
    # ── Emphasis ──────────────────────────────────────────────────────────────
    "pulse": {
        "preset_class": "emph", "preset_id": 14, "preset_subtype": 0,
        "filter": "", "transition": "none", "anim_elem": "animScale",
        "dir_map": {},
        "description": "Quick scale pulse to draw attention.",
        "best_for": ["kpi", "title"],
    },
    "spin": {
        "preset_class": "emph", "preset_id": 3, "preset_subtype": 0,
        "filter": "", "transition": "none", "anim_elem": "animRot",
        "dir_map": {
            "clockwise":         (0, ""),
            "counter_clockwise": (1, ""),
        },
        "description": "Rotation spin for emphasis.",
        "best_for": ["kpi", "icon"],
    },
    "grow_shrink": {
        "preset_class": "emph", "preset_id": 12, "preset_subtype": 0,
        "filter": "", "transition": "none", "anim_elem": "animScale",
        "dir_map": {},
        "description": "Scale grow/shrink for emphasis.",
        "best_for": ["kpi", "title"],
    },
    "bold_flash": {
        "preset_class": "emph", "preset_id": 24, "preset_subtype": 0,
        "filter": "", "transition": "none", "anim_elem": "animEffect",
        "dir_map": {},
        "description": "Bold flash emphasis on text.",
        "best_for": ["title", "body"],
    },
}

# ─── Valid triggers ───────────────────────────────────────────────────────────

VALID_TRIGGERS = {"on_click", "after_prev", "with_prev", "on_load"}

# ─── Smart defaults by slide type ─────────────────────────────────────────────

RECOMMENDED_ANIMATIONS = {
    "COVER": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",   "duration_ms": 800},
        {"shape": "subtitle", "effect": "fade",   "trigger": "after_prev","delay_ms": 200, "duration_ms": 600},
    ],
    "COVER_ALT": [
        {"shape": "title",    "effect": "fly_in", "trigger": "on_load",  "dir": "from_bottom", "duration_ms": 700},
        {"shape": "subtitle", "effect": "fade",   "trigger": "after_prev","delay_ms": 100, "duration_ms": 600},
    ],
    "COVER_FULL": [
        {"shape": "title",    "effect": "zoom",   "trigger": "on_load",  "dir": "in", "duration_ms": 800},
        {"shape": "subtitle", "effect": "fade",   "trigger": "after_prev","delay_ms": 200, "duration_ms": 600},
    ],
    "CHAPTER": [
        {"shape": "title",    "effect": "fly_in", "trigger": "on_load",  "dir": "from_left", "duration_ms": 600},
        {"shape": "subtitle", "effect": "fade",   "trigger": "after_prev","delay_ms": 100, "duration_ms": 500},
    ],
    "SECTION_BLUE": [
        {"shape": "title",    "effect": "wipe",   "trigger": "on_load",  "dir": "from_left", "duration_ms": 500},
    ],
    "SECTION_GREY": [
        {"shape": "title",    "effect": "wipe",   "trigger": "on_load",  "dir": "from_left", "duration_ms": 500},
    ],
    "CONTENT": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
        {"shape": "body",     "effect": "wipe",   "trigger": "on_click", "dir": "from_left",
         "duration_ms": 400, "text_build": "by_bullet"},
    ],
    "CONTENT_SIDEBAR": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
        {"shape": "body",     "effect": "fly_in", "trigger": "on_click", "dir": "from_left",
         "duration_ms": 400, "text_build": "by_bullet"},
    ],
    "TWO_COLUMN": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
        {"shape": "body",     "effect": "wipe",   "trigger": "on_click", "dir": "from_left", "duration_ms": 400},
        {"shape": "right",    "effect": "wipe",   "trigger": "on_click", "dir": "from_right", "duration_ms": 400},
    ],
    "QUOTE": [
        {"shape": "title",    "effect": "dissolve","trigger": "on_load", "duration_ms": 900},
    ],
    "CLOSING": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 800},
    ],
    # For slide types where body is a free shape (not a placeholder), only animate title
    "CHART": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
    ],
    "TABLE": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
    ],
    "CONTENT_CHART": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
        {"shape": "body",     "effect": "fly_in", "trigger": "on_click", "dir": "from_left", "duration_ms": 400},
    ],
    "CONTENT_TABLE": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
        {"shape": "body",     "effect": "fly_in", "trigger": "on_click", "dir": "from_left", "duration_ms": 400},
    ],
    "KPI": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
    ],
    "TIMELINE": [
        {"shape": "title",    "effect": "fade",   "trigger": "on_load",  "duration_ms": 400},
    ],
}


# ─── Validation ───────────────────────────────────────────────────────────────

def validate_animations(animations_spec, slide_num):
    warnings = []
    if not isinstance(animations_spec, list):
        warnings.append(f"Slide {slide_num}: 'animations' must be a list")
        return warnings

    for idx, anim in enumerate(animations_spec):
        prefix = f"Slide {slide_num}, animation {idx+1}"

        effect = anim.get("effect", "")
        if effect not in ANIMATION_CATALOG:
            warnings.append(
                f"{prefix}: Unknown effect '{effect}'. "
                f"Valid: {', '.join(sorted(ANIMATION_CATALOG.keys()))}"
            )
            continue

        meta = ANIMATION_CATALOG[effect]

        trigger = anim.get("trigger", "on_click")
        if trigger not in VALID_TRIGGERS:
            warnings.append(f"{prefix}: Invalid trigger '{trigger}'. Use: {', '.join(VALID_TRIGGERS)}")

        direction = anim.get("dir")
        if direction is not None:
            valid_dirs = list(meta["dir_map"].keys())
            if valid_dirs and direction not in valid_dirs:
                warnings.append(
                    f"{prefix}: effect '{effect}' dir='{direction}' invalid. "
                    f"Valid: {', '.join(valid_dirs)}"
                )
            elif not valid_dirs:
                warnings.append(f"{prefix}: effect '{effect}' does not support 'dir'")

        dur = anim.get("duration_ms")
        if dur is not None and (not isinstance(dur, int) or not 100 <= dur <= 5000):
            warnings.append(f"{prefix}: duration_ms must be integer 100–5000 (got {dur})")

        delay = anim.get("delay_ms")
        if delay is not None and (not isinstance(delay, int) or delay < 0):
            warnings.append(f"{prefix}: delay_ms must be a non-negative integer (got {delay})")

        text_build = anim.get("text_build")
        if text_build is not None and text_build not in ("all_at_once", "by_bullet"):
            warnings.append(f"{prefix}: text_build must be 'all_at_once' or 'by_bullet'")

        if anim.get("shape") is None:
            warnings.append(f"{prefix}: missing required 'shape' field")

    return warnings


# ─── Shape Resolution ─────────────────────────────────────────────────────────

_SHAPE_ROLE_PH_IDX = {
    "title":    0,
    "body":     1,
    "subtitle": 1,
    "left":     1,
    "right":    2,
}


def _resolve_shape_id(slide, shape_spec):
    if isinstance(shape_spec, int):
        for s in slide.shapes:
            if s.shape_id == shape_spec:
                return s.shape_id, s
        return None, None

    ph_idx = _SHAPE_ROLE_PH_IDX.get(shape_spec)
    if ph_idx is not None:
        try:
            ph = slide.placeholders[ph_idx]
            return ph.shape_id, ph
        except (KeyError, IndexError):
            pass

    for s in slide.shapes:
        if shape_spec.lower() in s.name.lower():
            return s.shape_id, s

    return None, None


# ─── Node ID counter ──────────────────────────────────────────────────────────

class _IdCounter:
    def __init__(self): self._v = 0
    def next(self):
        self._v += 1
        return str(self._v)


# ─── XML Builders — modelled exactly on real PowerPoint output ────────────────
#
# Reference structure (from a real PowerPoint file with animations):
#
#   <p:timing>
#     <p:tnLst>
#       <p:par>
#         <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
#           <p:childTnLst>
#             <p:seq concurrent="1" nextAc="seek">
#               <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
#                 <p:childTnLst>
#                   <!-- one <p:par> per animation click group -->
#                   <p:par>
#                     <p:cTn id="3" fill="hold">
#                       <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
#                       <p:childTnLst>
#                         <p:par>
#                           <p:cTn id="4" fill="hold">
#                             <p:stCondLst><p:cond delay="0"/></p:stCondLst>
#                             <p:childTnLst>
#                               <!-- first shape in click group: clickEffect -->
#                               <p:par>
#                                 <p:cTn id="5" presetID="..." presetClass="entr"
#                                        presetSubtype="0" fill="hold" grpId="0"
#                                        nodeType="clickEffect">
#                                   <p:stCondLst><p:cond delay="0"/></p:stCondLst>
#                                   <p:childTnLst>
#                                     <p:set>...</p:set>
#                                     <p:animEffect ...>...</p:animEffect>
#                                   </p:childTnLst>
#                                 </p:cTn>
#                               </p:par>
#                               <!-- same-click subsequent shapes: withEffect -->
#                               <p:par>
#                                 <p:cTn id="8" ... nodeType="withEffect">...</p:cTn>
#                               </p:par>
#                             </p:childTnLst>
#                           </p:cTn>
#                         </p:par>
#                       </p:childTnLst>
#                     </p:cTn>
#                   </p:par>
#                 </p:childTnLst>
#               </p:cTn>
#               <p:prevCondLst>
#                 <p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>
#               </p:prevCondLst>
#               <p:nextCondLst>
#                 <p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>
#               </p:nextCondLst>
#             </p:seq>
#           </p:childTnLst>
#         </p:cTn>
#       </p:par>
#     </p:tnLst>
#     <p:bldLst>
#       <p:bldP spid="2" grpId="0"/>
#       <p:bldP spid="3" grpId="0" build="p"/>
#     </p:bldLst>
#   </p:timing>


def _build_timing_skeleton(slide_elem, ctr):
    """
    Build the outer timing skeleton directly on slide_elem.
    Returns (seq_elem, seq_cTn_childTnLst, bldLst).
    Animations are appended to seq_cTn_childTnLst; bldLst gets p:bldP entries.
    """
    timing   = etree.SubElement(slide_elem, _p("timing"))
    tnLst    = etree.SubElement(timing, _p("tnLst"))

    root_par = etree.SubElement(tnLst, _p("par"))
    root_cTn = etree.SubElement(root_par, _p("cTn"),
                                id=ctr.next(), dur="indefinite",
                                restart="never", nodeType="tmRoot")
    root_childTnLst = etree.SubElement(root_cTn, _p("childTnLst"))

    seq     = etree.SubElement(root_childTnLst, _p("seq"),
                               concurrent="1", nextAc="seek")
    seq_cTn = etree.SubElement(seq, _p("cTn"),
                               id=ctr.next(), dur="indefinite", nodeType="mainSeq")
    seq_children = etree.SubElement(seq_cTn, _p("childTnLst"))

    # prevCondLst / nextCondLst go on the p:seq element (not on cTn)
    prev_cond_lst = etree.SubElement(seq, _p("prevCondLst"))
    prev_cond     = etree.SubElement(prev_cond_lst, _p("cond"),
                                     evt="onPrev", delay="0")
    etree.SubElement(etree.SubElement(prev_cond, _p("tgtEl")), _p("sldTgt"))

    next_cond_lst = etree.SubElement(seq, _p("nextCondLst"))
    next_cond     = etree.SubElement(next_cond_lst, _p("cond"),
                                     evt="onNext", delay="0")
    etree.SubElement(etree.SubElement(next_cond, _p("tgtEl")), _p("sldTgt"))

    bldLst = etree.SubElement(timing, _p("bldLst"))

    return seq_children, bldLst


def _outer_delay_for_trigger(trigger, delay_ms):
    """Return the delay string for the outermost p:cond in a click group."""
    if trigger == "on_click":
        return "indefinite"
    return str(delay_ms or 0)


def _node_type_for_trigger(trigger):
    return "clickEffect" if trigger == "on_click" else "withEffect"


def _add_visibility_set(parent, ctr, shape_id, exit_mode=False, para_idx=None):
    """
    Prepend a <p:set> that makes the shape visible (or hidden for exit).
    PowerPoint uses <p:boolVal val="0"/> in <p:to> for style.visibility.
    """
    p_set = etree.SubElement(parent, _p("set"))
    cBhvr = etree.SubElement(p_set, _p("cBhvr"))
    set_cTn = etree.SubElement(cBhvr, _p("cTn"), id=ctr.next(), dur="1", fill="hold")
    set_stCond = etree.SubElement(set_cTn, _p("stCondLst"))
    etree.SubElement(set_stCond, _p("cond"), delay="0")
    tgtEl  = etree.SubElement(cBhvr, _p("tgtEl"))
    spTgt  = etree.SubElement(tgtEl, _p("spTgt"), spid=str(shape_id))
    if para_idx is not None:
        txEl = etree.SubElement(spTgt, _p("txEl"))
        etree.SubElement(txEl, _p("pRg"), st=str(para_idx), end=str(para_idx))
    attrNameLst = etree.SubElement(cBhvr, _p("attrNameLst"))
    etree.SubElement(attrNameLst, _p("attrName")).text = "style.visibility"
    to_val = etree.SubElement(p_set, _p("to"))
    etree.SubElement(to_val, _p("strVal"), val="hidden" if exit_mode else "visible")


def _add_anim_effect(parent, ctr, shape_id, dur_ms, transition, filter_str, para_idx=None):
    """Append a <p:animEffect> element."""
    anim_eff = etree.SubElement(parent, _p("animEffect"),
                                transition=transition, filter=filter_str)
    cBhvr = etree.SubElement(anim_eff, _p("cBhvr"))
    etree.SubElement(cBhvr, _p("cTn"), id=ctr.next(), dur=str(dur_ms))
    tgtEl = etree.SubElement(cBhvr, _p("tgtEl"))
    spTgt = etree.SubElement(tgtEl, _p("spTgt"), spid=str(shape_id))
    if para_idx is not None:
        txEl = etree.SubElement(spTgt, _p("txEl"))
        etree.SubElement(txEl, _p("pRg"), st=str(para_idx), end=str(para_idx))


def _add_animScale(parent, ctr, shape_id, dur_ms, grow=False):
    """<p:animScale> for pulse / grow_shrink emphasis."""
    scale_pct = "150000" if grow else "110000"
    anim_scale = etree.SubElement(parent, _p("animScale"))
    cBhvr = etree.SubElement(anim_scale, _p("cBhvr"))
    etree.SubElement(cBhvr, _p("cTn"), id=ctr.next(), dur=str(dur_ms),
                     fill="hold", autoRev="1")
    tgtEl = etree.SubElement(cBhvr, _p("tgtEl"))
    etree.SubElement(tgtEl, _p("spTgt"), spid=str(shape_id))
    etree.SubElement(anim_scale, _p("by"), x=scale_pct, y=scale_pct)


def _add_animRot(parent, ctr, shape_id, dur_ms):
    """<p:animRot> for spin emphasis."""
    anim_rot = etree.SubElement(parent, _p("animRot"))
    cBhvr = etree.SubElement(anim_rot, _p("cBhvr"))
    etree.SubElement(cBhvr, _p("cTn"), id=ctr.next(), dur=str(dur_ms), fill="hold")
    tgtEl = etree.SubElement(cBhvr, _p("tgtEl"))
    etree.SubElement(tgtEl, _p("spTgt"), spid=str(shape_id))
    etree.SubElement(anim_rot, _p("by"), val="21600000")


def _append_one_effect(click_group_children, ctr, shape_id, meta,
                        dur_ms, grp_id, node_type, para_idx=None):
    """
    Append one effect par (<p:par><p:cTn ...><p:childTnLst>...)) to
    click_group_children. Used for both single animations and by-bullet.
    """
    inner_par = etree.SubElement(click_group_children, _p("par"))
    # Attribute order must match PowerPoint: presetID presetClass presetSubtype fill grpId nodeType
    # Attribute order must match PowerPoint: presetID presetClass presetSubtype fill [grpId] nodeType
    inner_cTn_attrib = {
        "id": ctr.next(),
        "presetID": str(meta["preset_id"]),
        "presetClass": meta["preset_class"],
        "presetSubtype": str(meta["preset_subtype"]),
        "fill": "hold",
    }
    if grp_id is not None:
        inner_cTn_attrib["grpId"] = str(grp_id)
    inner_cTn_attrib["nodeType"] = node_type
    inner_cTn = etree.SubElement(inner_par, _p("cTn"), **inner_cTn_attrib)
    inner_stCond = etree.SubElement(inner_cTn, _p("stCondLst"))
    etree.SubElement(inner_stCond, _p("cond"), delay="0")
    inner_children = etree.SubElement(inner_cTn, _p("childTnLst"))

    anim_tag = meta["anim_elem"]
    exit_mode = meta["preset_class"] == "exit"

    if anim_tag == "animEffect":
        _add_visibility_set(inner_children, ctr, shape_id, exit_mode, para_idx)
        _add_anim_effect(inner_children, ctr, shape_id, dur_ms,
                         meta["transition"], meta["filter"], para_idx)
    elif anim_tag == "set":
        _add_visibility_set(inner_children, ctr, shape_id, exit_mode, para_idx)
    elif anim_tag == "animScale":
        _add_animScale(inner_children, ctr, shape_id, dur_ms,
                       grow=meta["preset_id"] == 12)
    elif anim_tag == "animRot":
        _add_animRot(inner_children, ctr, shape_id, dur_ms)


def _add_single_animation(seq_children, ctr, shape_id, meta,
                           trigger, delay_ms, dur_ms, grp_id=0, para_idx=None):
    """
    Append one click group (outer par → middle par → inner effect par)
    matching the structure PowerPoint generates.

    seq_children  ← p:cTn[mainSeq]/p:childTnLst
      p:par  (click group)
        p:cTn fill=hold
          p:stCondLst > p:cond delay="indefinite"  (or "0" for auto)
          p:childTnLst
            p:par  (middle layer)
              p:cTn fill=hold
                p:stCondLst > p:cond delay="0"
                p:childTnLst
                  p:par  (effect)
                    p:cTn presetID=... nodeType=clickEffect/withEffect
                      p:stCondLst > p:cond delay="0"
                      p:childTnLst
                        p:set ...
                        p:animEffect ...
    """
    outer_delay = _outer_delay_for_trigger(trigger, delay_ms)
    node_type   = _node_type_for_trigger(trigger)

    # Outer click-group par
    outer_par     = etree.SubElement(seq_children, _p("par"))
    outer_cTn     = etree.SubElement(outer_par, _p("cTn"), id=ctr.next(), fill="hold")
    outer_stCond  = etree.SubElement(outer_cTn, _p("stCondLst"))
    etree.SubElement(outer_stCond, _p("cond"), delay=outer_delay)
    outer_children = etree.SubElement(outer_cTn, _p("childTnLst"))

    # Middle par (always delay="0")
    mid_par      = etree.SubElement(outer_children, _p("par"))
    mid_cTn      = etree.SubElement(mid_par, _p("cTn"), id=ctr.next(), fill="hold")
    mid_stCond   = etree.SubElement(mid_cTn, _p("stCondLst"))
    etree.SubElement(mid_stCond, _p("cond"), delay="0")
    mid_children = etree.SubElement(mid_cTn, _p("childTnLst"))

    # Effect par
    _append_one_effect(mid_children, ctr, shape_id, meta,
                        dur_ms, grp_id, node_type, para_idx)


def _add_bybullet_sequence(seq_children, ctr, slide, shape_id, meta,
                            trigger, delay_ms, dur_ms):
    """
    One click group per non-empty paragraph. First uses the specified trigger;
    each subsequent one fires with the previous (afterPrev / withPrev behavior
    PowerPoint uses: each bullet is its own click group with delay=indefinite,
    except sub-bullets which follow with delay=0).
    """
    shape = None
    for s in slide.shapes:
        if s.shape_id == shape_id:
            shape = s
            break

    if shape is None or not hasattr(shape, "text_frame"):
        _add_single_animation(seq_children, ctr, shape_id, meta,
                               trigger, delay_ms, dur_ms)
        return

    paragraphs = shape.text_frame.paragraphs
    is_first = True
    for i, para in enumerate(paragraphs):
        if not para.text.strip():
            continue
        t   = trigger if is_first else "on_click"
        gid = 0 if is_first else None   # PowerPoint omits grpId on bullets after first
        d   = delay_ms if is_first else 0
        is_first = False
        _add_single_animation(seq_children, ctr, shape_id, meta,
                               t, d, dur_ms, grp_id=gid, para_idx=i)


# ─── Main Entry Point ─────────────────────────────────────────────────────────

def inject_animations(slide, animations_spec, slide_type="CONTENT"):
    """
    Write a <p:timing> block onto slide matching real PowerPoint output structure.

    - animations_spec=None  → use RECOMMENDED_ANIMATIONS[slide_type]
    - animations_spec=[]    → suppress all animations
    - animations_spec=[...] → use the provided list
    """
    if animations_spec is None:
        animations_spec = RECOMMENDED_ANIMATIONS.get(slide_type, [])
    if not animations_spec:
        return

    slide_elem = slide._element

    # Remove any existing <p:timing>
    for child in list(slide_elem):
        if child.tag == _p("timing"):
            slide_elem.remove(child)

    ctr = _IdCounter()
    seq_children, bldLst = _build_timing_skeleton(slide_elem, ctr)

    for anim in animations_spec:
        effect = anim.get("effect", "fade")
        if effect not in ANIMATION_CATALOG:
            print(f"  WARNING: Unknown animation effect '{effect}', skipping")
            continue

        meta = dict(ANIMATION_CATALOG[effect])

        shape_spec = anim.get("shape", "body")
        shape_id, shape_obj = _resolve_shape_id(slide, shape_spec)
        if shape_id is None:
            print(f"  WARNING: Shape '{shape_spec}' not found on slide, skipping animation")
            continue

        # Apply direction override
        direction = anim.get("dir")
        if direction and direction in meta["dir_map"]:
            subtype, filter_str = meta["dir_map"][direction]
            meta["preset_subtype"] = subtype
            meta["filter"] = filter_str

        dur_ms     = anim.get("duration_ms", 500)
        delay_ms   = anim.get("delay_ms", 0)
        trigger    = anim.get("trigger", "on_click")
        text_build = anim.get("text_build")

        if text_build == "by_bullet" and meta["anim_elem"] == "animEffect":
            _add_bybullet_sequence(seq_children, ctr, slide, shape_id, meta,
                                   trigger, delay_ms, dur_ms)
            # bldLst entry with build="p" for by-bullet
            etree.SubElement(bldLst, _p("bldP"),
                             spid=str(shape_id), grpId="0", build="p")
        else:
            _add_single_animation(seq_children, ctr, shape_id, meta,
                                  trigger, delay_ms, dur_ms)
            etree.SubElement(bldLst, _p("bldP"),
                             spid=str(shape_id), grpId="0")

