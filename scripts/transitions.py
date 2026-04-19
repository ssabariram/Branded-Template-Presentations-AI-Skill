"""
<company_name> Slide Transitions — OOXML transition injection via raw lxml.

python-pptx has no public API for transitions; we manipulate the XML directly.

Usage:
    from transitions import inject_transition, validate_transition, TRANSITION_CATALOG

The public surface:
    inject_transition(slide, transition_spec)  — writes <p:transition> onto the slide element
    validate_transition(transition_spec, slide_num) — returns list of warning strings
    TRANSITION_CATALOG — dict of every supported transition name → metadata
"""

from lxml import etree

# ─── Namespaces ───────────────────────────────────────────────────────────────

NS_P   = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NS_MC  = "http://schemas.openxmlformats.org/markup-compatibility/2006"

def _ptag(local):  return f"{{{NS_P}}}{local}"
def _p14tag(local): return f"{{{NS_P14}}}{local}"
def _mctag(local):  return f"{{{NS_MC}}}{local}"


# ─── Transition Catalog ───────────────────────────────────────────────────────
#
# Each entry describes one supported transition.
#
# Fields:
#   ns          "p" (ECMA-376 core) or "p14" (PowerPoint 2010+ extension)
#   xml_tag     local XML element name inside <p:transition>
#   attrs       fixed XML attributes on the transition child element (dict)
#   dir_values  allowed values for the optional "dir" parameter (empty = no dir)
#   description human-readable description
#   best_for    list of slide types this transition suits
#
# The "dir" param, when present, maps to the child element's "dir" attribute.
# Some transitions use "dir" on the *transition* element itself (cover, pull,
# push, wipe) — those are marked with dir_on_parent=True.

TRANSITION_CATALOG = {
    # ── Core (ECMA-376) ───────────────────────────────────────────────────────
    "fade": {
        "ns": "p", "xml_tag": "fade",
        "attrs": {},
        "dir_values": [],
        "description": "Smooth cross-fade to next slide.",
        "best_for": ["COVER", "COVER_ALT", "COVER_FULL", "CLOSING", "QUOTE", "CHAPTER"],
    },
    "cut": {
        "ns": "p", "xml_tag": "cut",
        "attrs": {},
        "dir_values": [],
        "description": "Instant cut with no animation (default PowerPoint behaviour).",
        "best_for": ["CONTENT", "CONTENT_SIDEBAR", "TWO_COLUMN", "TABLE", "CHART"],
    },
    "push": {
        "ns": "p", "xml_tag": "push",
        "attrs": {},
        "dir_values": ["l", "r", "u", "d"],
        "dir_on_parent": True,
        "description": "New slide pushes current slide off-screen.",
        "best_for": ["CHAPTER", "SECTION_BLUE", "SECTION_GREY", "TIMELINE"],
    },
    "cover": {
        "ns": "p", "xml_tag": "cover",
        "attrs": {},
        "dir_values": ["l", "r", "u", "d", "lu", "ru", "ld", "rd"],
        "dir_on_parent": True,
        "description": "New slide slides over the current slide.",
        "best_for": ["CHAPTER", "SECTION_BLUE", "SECTION_GREY"],
    },
    "pull": {
        "ns": "p", "xml_tag": "pull",
        "attrs": {},
        "dir_values": ["l", "r", "u", "d", "lu", "ru", "ld", "rd"],
        "dir_on_parent": True,
        "description": "Current slide pulls away to reveal next slide.",
        "best_for": ["CLOSING", "QUOTE"],
    },
    "wipe": {
        "ns": "p", "xml_tag": "wipe",
        "attrs": {},
        "dir_values": ["l", "r", "u", "d"],
        "dir_on_parent": True,
        "description": "A wipe that reveals the next slide from one edge.",
        "best_for": ["CONTENT", "CONTENT_SIDEBAR", "CONTENT_CHART", "CONTENT_TABLE"],
    },
    "dissolve": {
        "ns": "p", "xml_tag": "dissolve",
        "attrs": {},
        "dir_values": [],
        "description": "Pixel-dissolve between slides.",
        "best_for": ["COVER", "CLOSING", "QUOTE"],
    },
    "split": {
        "ns": "p", "xml_tag": "split",
        "attrs": {},
        "dir_values": ["horz", "vert"],  # mapped to orient attr
        "dir_attr": "orient",
        "description": "Slide splits open/closes from center.",
        "best_for": ["TWO_COLUMN", "CONTENT_CHART", "SECTION_BLUE", "SECTION_GREY"],
    },
    "zoom": {
        "ns": "p", "xml_tag": "zoom",
        "attrs": {},
        "dir_values": ["in", "out"],
        "dir_attr": "dir",
        "description": "Zooms in or out to transition.",
        "best_for": ["KPI", "QUOTE", "COVER_FULL"],
    },
    "wheel": {
        "ns": "p", "xml_tag": "wheel",
        "attrs": {"spokes": "4"},
        "dir_values": [],
        "description": "Rotating wheel wipe.",
        "best_for": ["KPI", "TIMELINE"],
    },
    "blinds": {
        "ns": "p", "xml_tag": "blinds",
        "attrs": {},
        "dir_values": ["horz", "vert"],
        "dir_attr": "dir",
        "description": "Venetian-blinds effect.",
        "best_for": ["TABLE", "CONTENT_TABLE"],
    },
    "checker": {
        "ns": "p", "xml_tag": "checker",
        "attrs": {},
        "dir_values": ["horz", "vert"],
        "dir_attr": "dir",
        "description": "Checkerboard pattern reveal.",
        "best_for": ["KPI"],
    },
    "circle": {
        "ns": "p", "xml_tag": "circle",
        "attrs": {},
        "dir_values": [],
        "description": "Circular iris wipe.",
        "best_for": ["QUOTE", "KPI"],
    },
    "diamond": {
        "ns": "p", "xml_tag": "diamond",
        "attrs": {},
        "dir_values": [],
        "description": "Diamond-shaped iris wipe.",
        "best_for": ["QUOTE"],
    },
    "plus": {
        "ns": "p", "xml_tag": "plus",
        "attrs": {},
        "dir_values": [],
        "description": "Plus / cross-shaped wipe.",
        "best_for": ["SECTION_BLUE", "SECTION_GREY"],
    },
    "wedge": {
        "ns": "p", "xml_tag": "wedge",
        "attrs": {},
        "dir_values": [],
        "description": "Clock-wipe wedge.",
        "best_for": ["TIMELINE", "CHAPTER"],
    },
    "comb": {
        "ns": "p", "xml_tag": "comb",
        "attrs": {},
        "dir_values": ["horz", "vert"],
        "dir_attr": "dir",
        "description": "Interlocking comb effect.",
        "best_for": ["TABLE", "CONTENT_TABLE"],
    },
    "randomBar": {
        "ns": "p", "xml_tag": "randomBar",
        "attrs": {},
        "dir_values": ["horz", "vert"],
        "dir_attr": "dir",
        "description": "Random horizontal or vertical bars.",
        "best_for": ["CONTENT", "CONTENT_SIDEBAR"],
    },
    "strips": {
        "ns": "p", "xml_tag": "strips",
        "attrs": {},
        "dir_values": ["lu", "ru", "ld", "rd"],
        "dir_attr": "dir",
        "description": "Diagonal strips.",
        "best_for": ["CHAPTER", "SECTION_BLUE"],
    },
    "newsflash": {
        "ns": "p", "xml_tag": "newsflash",
        "attrs": {},
        "dir_values": [],
        "description": "Spin-and-zoom newsflash effect.",
        "best_for": ["QUOTE", "KPI"],
    },
    "random": {
        "ns": "p", "xml_tag": "random",
        "attrs": {},
        "dir_values": [],
        "description": "PowerPoint picks a random transition. Avoid in professional decks.",
        "best_for": [],
    },

    # ── PowerPoint 2010+ (p14 namespace) ─────────────────────────────────────
    "conveyor": {
        "ns": "p14", "xml_tag": "conveyor",
        "attrs": {},
        "dir_values": ["l", "r"],
        "dir_attr": "dir",
        "description": "Conveyor belt slides in from the side.",
        "best_for": ["TIMELINE", "CHAPTER"],
    },
    "doors": {
        "ns": "p14", "xml_tag": "doors",
        "attrs": {},
        "dir_values": ["horz", "vert"],
        "dir_attr": "dir",
        "description": "Double-door open/close reveal.",
        "best_for": ["COVER", "COVER_ALT", "CHAPTER", "SECTION_BLUE"],
    },
    "ferris": {
        "ns": "p14", "xml_tag": "ferris",
        "attrs": {},
        "dir_values": ["l", "r"],
        "dir_attr": "dir",
        "description": "Ferris-wheel rotation of slides.",
        "best_for": ["KPI", "TWO_COLUMN"],
    },
    "flip": {
        "ns": "p14", "xml_tag": "flip",
        "attrs": {},
        "dir_values": ["l", "r"],
        "dir_attr": "dir",
        "description": "3-D card flip.",
        "best_for": ["TWO_COLUMN", "COVER_FULL"],
    },
    "flythrough": {
        "ns": "p14", "xml_tag": "flythrough",
        "attrs": {},
        "dir_values": ["in", "out"],
        "dir_attr": "dir",
        "description": "Camera flies through to next slide.",
        "best_for": ["COVER", "COVER_FULL", "CHAPTER"],
    },
    "gallery": {
        "ns": "p14", "xml_tag": "gallery",
        "attrs": {},
        "dir_values": ["l", "r"],
        "dir_attr": "dir",
        "description": "Gallery scroll across slides.",
        "best_for": ["CONTENT", "CONTENT_SIDEBAR", "CHART"],
    },
    "glitter": {
        "ns": "p14", "xml_tag": "glitter",
        "attrs": {},
        "dir_values": ["l", "r", "u", "d"],
        "dir_attr": "dir",
        "description": "Glitter/sparkle particles sweep across.",
        "best_for": ["COVER", "CLOSING", "QUOTE"],
    },
    "honeycomb": {
        "ns": "p14", "xml_tag": "honeycomb",
        "attrs": {},
        "dir_values": [],
        "description": "Hexagonal honeycomb reveal.",
        "best_for": ["KPI", "SECTION_BLUE", "SECTION_GREY"],
    },
    "pan": {
        "ns": "p14", "xml_tag": "pan",
        "attrs": {},
        "dir_values": ["l", "r", "u", "d"],
        "dir_attr": "dir",
        "description": "Camera pan to next slide.",
        "best_for": ["TIMELINE", "CONTENT"],
    },
    "prism": {
        "ns": "p14", "xml_tag": "prism",
        "attrs": {},
        "dir_values": ["l", "r", "u", "d"],
        "dir_attr": "dir",
        "description": "Prism refraction effect.",
        "best_for": ["CHAPTER", "SECTION_BLUE", "SECTION_GREY"],
    },
    "reveal": {
        "ns": "p14", "xml_tag": "reveal",
        "attrs": {},
        "dir_values": ["l", "r"],
        "dir_attr": "dir",
        "description": "Slide peels back to reveal the next.",
        "best_for": ["TWO_COLUMN", "CONTENT_CHART", "CONTENT_TABLE"],
    },
    "ripple": {
        "ns": "p14", "xml_tag": "ripple",
        "attrs": {},
        "dir_values": [],
        "description": "Water ripple from center.",
        "best_for": ["QUOTE", "CLOSING"],
    },
    "shred": {
        "ns": "p14", "xml_tag": "shred",
        "attrs": {},
        "dir_values": ["in", "out"],
        "dir_attr": "dir",
        "description": "Slide shreds into pieces.",
        "best_for": ["CLOSING", "QUOTE"],
    },
    "switch": {
        "ns": "p14", "xml_tag": "switch",
        "attrs": {},
        "dir_values": ["l", "r"],
        "dir_attr": "dir",
        "description": "3-D flip/switch between slides.",
        "best_for": ["TWO_COLUMN", "COVER_ALT"],
    },
    "vortex": {
        "ns": "p14", "xml_tag": "vortex",
        "attrs": {},
        "dir_values": ["l", "r", "u", "d"],
        "dir_attr": "dir",
        "description": "Vortex swirl.",
        "best_for": ["CLOSING", "CHAPTER"],
    },
    "warp": {
        "ns": "p14", "xml_tag": "warp",
        "attrs": {},
        "dir_values": ["in", "out"],
        "dir_attr": "dir",
        "description": "Warp/perspective zoom.",
        "best_for": ["COVER_FULL", "QUOTE"],
    },
    "window": {
        "ns": "p14", "xml_tag": "window",
        "attrs": {},
        "dir_values": ["horz", "vert"],
        "dir_attr": "dir",
        "description": "Window blinds / pane reveal.",
        "best_for": ["CONTENT", "TABLE", "CHART"],
    },
    "flash": {
        "ns": "p14", "xml_tag": "flash",
        "attrs": {},
        "dir_values": [],
        "description": "Quick flash to white then to next slide.",
        "best_for": ["QUOTE", "KPI"],
    },
    "wheelReverse": {
        "ns": "p14", "xml_tag": "wheelReverse",
        "attrs": {"spokes": "4"},
        "dir_values": [],
        "description": "Reverse-spinning wheel wipe.",
        "best_for": ["KPI", "TIMELINE"],
    },
}

# ─── Speed / Duration defaults ─────────────────────────────────────────────────

VALID_SPEEDS = {"slow", "med", "fast"}

# Default durations (ms) per speed — used when speed is provided without p14:dur
_SPEED_TO_DUR = {"slow": 1500, "med": 800, "fast": 400}

# Sensible default speed per slide role
_TYPE_DEFAULT_SPEED = {
    "COVER": "slow", "COVER_ALT": "slow", "COVER_FULL": "slow",
    "CLOSING": "slow",
    "CHAPTER": "med", "SECTION_BLUE": "med", "SECTION_GREY": "med",
    "QUOTE": "slow",
    "KPI": "med", "TIMELINE": "med",
    "CONTENT": "fast", "CONTENT_SIDEBAR": "fast",
    "TWO_COLUMN": "fast", "CHART": "fast", "TABLE": "fast",
    "CONTENT_CHART": "fast", "CONTENT_TABLE": "fast",
}


# ─── Validation ───────────────────────────────────────────────────────────────

def validate_transition(transition_spec, slide_num):
    """
    Validate a transition specification.  Returns a list of warning strings.

    transition_spec fields:
        type        (str, required)  — key from TRANSITION_CATALOG
        speed       (str, optional)  — "slow" | "med" | "fast"
        duration_ms (int, optional)  — override animation duration in milliseconds
        dir         (str, optional)  — direction qualifier; valid values depend on type
        advance_ms  (int, optional)  — auto-advance after N milliseconds (0 = click only)
    """
    warnings = []
    if not isinstance(transition_spec, dict):
        warnings.append(f"Slide {slide_num}: transition must be an object, got {type(transition_spec).__name__}")
        return warnings

    t = transition_spec.get("type", "")
    if t not in TRANSITION_CATALOG:
        warnings.append(
            f"Slide {slide_num}: Unknown transition type '{t}'. "
            f"Valid types: {', '.join(sorted(TRANSITION_CATALOG.keys()))}"
        )
        return warnings

    meta = TRANSITION_CATALOG[t]

    speed = transition_spec.get("speed")
    if speed is not None and speed not in VALID_SPEEDS:
        warnings.append(
            f"Slide {slide_num}: transition speed '{speed}' invalid. Use: slow | med | fast"
        )

    dur = transition_spec.get("duration_ms")
    if dur is not None and (not isinstance(dur, int) or dur < 100 or dur > 10000):
        warnings.append(
            f"Slide {slide_num}: transition duration_ms must be an integer 100–10000 (got {dur})"
        )

    adv = transition_spec.get("advance_ms")
    if adv is not None and (not isinstance(adv, int) or adv < 0):
        warnings.append(
            f"Slide {slide_num}: transition advance_ms must be a non-negative integer (got {adv})"
        )

    direction = transition_spec.get("dir")
    if direction is not None:
        valid_dirs = meta.get("dir_values", [])
        if not valid_dirs:
            warnings.append(
                f"Slide {slide_num}: transition '{t}' does not support 'dir' (got '{direction}')"
            )
        elif direction not in valid_dirs:
            warnings.append(
                f"Slide {slide_num}: transition '{t}' dir='{direction}' invalid. "
                f"Valid: {', '.join(valid_dirs)}"
            )

    return warnings


# ─── XML Injection ────────────────────────────────────────────────────────────

def inject_transition(slide, transition_spec, slide_type="CONTENT"):
    """
    Write a <p:transition> element onto a slide's XML element.

    Parameters
    ----------
    slide           : pptx Slide object
    transition_spec : dict — see validate_transition for field docs
    slide_type      : str  — used only to pick a sensible default speed

    The function:
    1. Removes any existing <p:transition> or <mc:AlternateContent> wrapper
       that contains one.
    2. Builds the new <p:transition> element (core ECMA-376 transitions) or
       wraps it in <mc:AlternateContent> (p14 extensions).
    3. Appends it to the slide element (after <p:cSld> and <p:clrMapOvr>).
    """
    if not transition_spec:
        return

    t = transition_spec.get("type", "")
    if t not in TRANSITION_CATALOG:
        print(f"  WARNING: Unknown transition type '{t}', skipping")
        return

    meta = TRANSITION_CATALOG[t]
    slide_elem = slide._element

    # Remove any existing transition (plain or wrapped in AlternateContent)
    _remove_existing_transition(slide_elem)

    # Resolve speed and duration
    speed = transition_spec.get("speed") or _TYPE_DEFAULT_SPEED.get(slide_type, "fast")
    duration_ms = transition_spec.get("duration_ms", _SPEED_TO_DUR[speed])
    advance_ms = transition_spec.get("advance_ms")
    direction = transition_spec.get("dir")

    if meta["ns"] == "p":
        # Build directly as SubElement — no standalone etree.Element() calls
        _build_core_transition(slide_elem, meta, speed, duration_ms, advance_ms, direction)
    else:
        # p14 transitions require <mc:AlternateContent> for back-compat
        _build_p14_transition(slide_elem, meta, speed, duration_ms, advance_ms, direction)


def _remove_existing_transition(slide_elem):
    """Remove <p:transition> and any <mc:AlternateContent> wrapping one."""
    # Plain <p:transition>
    for child in list(slide_elem):
        if child.tag == _ptag("transition"):
            slide_elem.remove(child)
        # <mc:AlternateContent> that wraps a transition
        elif child.tag == _mctag("AlternateContent"):
            # Peek inside to see if it contains a transition
            if child.find(f".//{_ptag('transition')}") is not None:
                slide_elem.remove(child)


def _add_transition_child(tr_elem, meta, direction):
    """Add the transition-type child element to an existing <p:transition>."""
    ns = NS_P if meta["ns"] == "p" else NS_P14
    child = etree.SubElement(tr_elem, f"{{{ns}}}{meta['xml_tag']}", **meta.get("attrs", {}))
    if direction and direction in meta.get("dir_values", []):
        dir_attr = meta.get("dir_attr", "dir")
        on_parent = meta.get("dir_on_parent", False)
        if on_parent:
            tr_elem.set("dir", direction)
        else:
            child.set(dir_attr, direction)


def _build_core_transition(parent, meta, speed, duration_ms, advance_ms, direction):
    """Append a plain <p:transition> directly to parent (no standalone element)."""
    attrs = {"spd": speed, "advClick": "1"}
    if advance_ms is not None and advance_ms > 0:
        attrs["advTm"] = str(advance_ms)
    tr = etree.SubElement(parent, _ptag("transition"), **attrs)
    _add_transition_child(tr, meta, direction)


def _build_p14_transition(parent, meta, speed, duration_ms, advance_ms, direction):
    """
    Append <mc:AlternateContent> wrapping a p14-extended transition directly to parent.
    The mc/p14 namespaces must be declared on the slide root for clean serialization.
    """
    slide_root = parent  # parent IS slide_elem for p14 transitions
    # Ensure mc and p14 namespaces are declared on the slide root element
    # by adding them to its nsmap via makeelement trick — lxml won't add xmlns
    # attrs via .set(), so we patch the nsmap by re-creating the element isn't
    # feasible. Instead we declare them inline on AlternateContent (accepted by PowerPoint).
    alt = etree.SubElement(parent, _mctag("AlternateContent"),
                           nsmap={"mc": NS_MC, "p14": NS_P14})

    # ── Choice (modern PowerPoint) ──
    choice = etree.SubElement(alt, _mctag("Choice"))
    choice.set("Requires", "p14")
    attrs = {"spd": speed, "advClick": "1", f"{{{NS_P14}}}dur": str(duration_ms)}
    if advance_ms is not None and advance_ms > 0:
        attrs["advTm"] = str(advance_ms)
    tr_choice = etree.SubElement(choice, _ptag("transition"), **attrs)
    _add_transition_child(tr_choice, meta, direction)

    # ── Fallback (legacy PowerPoint / LibreOffice) ──
    fallback = etree.SubElement(alt, _mctag("Fallback"))
    fb_attrs = {"spd": speed, "advClick": "1"}
    if advance_ms is not None and advance_ms > 0:
        fb_attrs["advTm"] = str(advance_ms)
    tr_fallback = etree.SubElement(fallback, _ptag("transition"), **fb_attrs)
    etree.SubElement(tr_fallback, _ptag("fade"))


# ─── Intelligent Default Recommendations ─────────────────────────────────────

# Maps slide type → recommended transition name.
# These are chosen so transitions feel purposeful:
#   - Cover/Closing: cinematic, polished (fade, glitter, ripple)
#   - Chapter/Section: motion that signals a structural shift (push, doors, prism)
#   - Content slides: subtle, non-distracting (wipe, cut, fade)
#   - Data/KPI: energetic but professional (zoom, flash, wheel)
#   - Quote: dramatic pause (dissolve, warp)
#   - Timeline: directional motion that mirrors progress (conveyor, pan)

RECOMMENDED_TRANSITION = {
    "COVER":          {"type": "fade",      "speed": "slow"},
    "COVER_ALT":      {"type": "fade",      "speed": "slow"},
    "COVER_FULL":     {"type": "flythrough","speed": "slow", "dir": "in"},
    "CHAPTER":        {"type": "push",      "speed": "med",  "dir": "l"},
    "SECTION_BLUE":   {"type": "doors",     "speed": "med",  "dir": "horz"},
    "SECTION_GREY":   {"type": "prism",     "speed": "med",  "dir": "l"},
    "CONTENT":        {"type": "wipe",      "speed": "fast", "dir": "l"},
    "CONTENT_SIDEBAR":{"type": "wipe",      "speed": "fast", "dir": "l"},
    "TWO_COLUMN":     {"type": "split",     "speed": "fast", "dir": "horz"},
    "QUOTE":          {"type": "dissolve",  "speed": "slow"},
    "CLOSING":        {"type": "fade",      "speed": "slow"},
    "CHART":          {"type": "wipe",      "speed": "fast", "dir": "l"},
    "TABLE":          {"type": "wipe",      "speed": "fast", "dir": "l"},
    "CONTENT_CHART":  {"type": "reveal",    "speed": "fast", "dir": "l"},
    "CONTENT_TABLE":  {"type": "reveal",    "speed": "fast", "dir": "l"},
    "KPI":            {"type": "zoom",      "speed": "med",  "dir": "in"},
    "TIMELINE":       {"type": "conveyor",  "speed": "med",  "dir": "l"},
}
