"""Built-in default S1000D BREX modules.

Ports the built-in default BREX handling from `s1kd-brexcheck`
(`default_brex_dmc`, `search_brex_fname_from_default_brex`, `load_brex` in
`s1kd-brexcheck.c`): the seven BREX data modules `s1kd-brexcheck` ships
(business rules exchange issues A/D/E/F/G/H of the S1000D specification
itself, plus the legacy `DMC-AE-A-...` BREX for S1000D <= 3.0) are bundled
under `acd/brex/` so a schema-appropriate default BREX is always available,
whether selected explicitly (`-B`/`--default-brex`) or used as a fallback
when a referenced BREX data module cannot be located.
"""

from os.path import join
from os.path import dirname
from os.path import abspath

from .s1000d import ref_dict_to_str


DEFAULT_BREX_DIR = join(dirname(abspath(__file__)), "brex")

# logical DMC (model+sysDiff+sys+subSys+subSubSys+assy+disassy+disassyVariant
# +info+infoVariant+itemLoc, i.e. ref_dict_to_str() with no issue/inWork) ->
# (bundled filename, issueNumber, inWork), ported from the hard-coded table
# in load_brex() / search_brex_fname_from_default_brex() (s1kd-brexcheck.c).
_DEFAULT_BREX = {
    "DMC-S1000D-H-04-10-0301-00A-022A-D": ("DMC-S1000D-H-04-10-0301-00A-022A-D_001-00_EN-US.XML", "001", "00"),
    "DMC-S1000D-G-04-10-0301-00A-022A-D": ("DMC-S1000D-G-04-10-0301-00A-022A-D_001-00_EN-US.XML", "001", "00"),
    "DMC-S1000D-F-04-10-0301-00A-022A-D": ("DMC-S1000D-F-04-10-0301-00A-022A-D_001-00_EN-US.XML", "001", "00"),
    "DMC-S1000D-E-04-10-0301-00A-022A-D": ("DMC-S1000D-E-04-10-0301-00A-022A-D_012-00_EN-US.XML", "012", "00"),
    "DMC-S1000D-D-04-10-0301-00A-022A-D": ("DMC-S1000D-D-04-10-0301-00A-022A-D_006-00_EN-US.XML", "006", "00"),
    "DMC-S1000D-A-04-10-0301-00A-022A-D": ("DMC-S1000D-A-04-10-0301-00A-022A-D_005-00_EN-US.XML", "005", "00"),
    "DMC-AE-A-04-10-0301-00A-022A-D": ("DMC-AE-A-04-10-0301-00A-022A-D_003-00.XML", "003", "00"),
}


def default_brex_dmc(schema: str) -> str:
    """Return the logical DMC of the built-in default BREX matching a schema.

    Port of `default_brex_dmc` (`s1kd-brexcheck.c:1553-1577`): the object's
    declared `xsi:noNamespaceSchemaLocation` (as returned by
    `xml_processing.get_schema_from_xml`) selects the BREX for the issue of
    the specification it targets, by substring match against the schema URL
    (so e.g. a `S1000D_4-0-1` or `S1000D_4-0-2` sub-issue still matches
    `S1000D_4-0`, exactly as in the C original). A missing or unrecognised
    schema (S1000D 6.0+, or no `xsi:noNamespaceSchemaLocation` at all)
    resolves to the latest bundled issue, and anything else (S1000D <= 3.0)
    resolves to the legacy `DMC-AE-A-...` BREX. Issue "A"
    (`DMC-S1000D-A-...`) is never selected here -- like the C original, it
    is only reachable via `find_default_brex_fallback`.

    Args:
        schema (str): schema URL as returned by `xml_processing.get_schema_from_xml`,
            or None

    Returns:
        str: logical DMC key into the built-in default BREX table, usable
            with `default_brex_path`
    """
    if not schema or "S1000D_6" in schema:
        return "DMC-S1000D-H-04-10-0301-00A-022A-D"
    if "S1000D_5-0" in schema:
        return "DMC-S1000D-G-04-10-0301-00A-022A-D"
    if "S1000D_4-2" in schema:
        return "DMC-S1000D-F-04-10-0301-00A-022A-D"
    if "S1000D_4-1" in schema:
        return "DMC-S1000D-E-04-10-0301-00A-022A-D"
    if "S1000D_4-0" in schema:
        return "DMC-S1000D-D-04-10-0301-00A-022A-D"
    return "DMC-AE-A-04-10-0301-00A-022A-D"


def default_brex_path(logical_dmc: str) -> str:
    """Resolve a built-in default BREX's logical DMC to its bundled file path.

    Args:
        logical_dmc (str): a logical DMC key of the built-in default BREX
            table, e.g. as returned by `default_brex_dmc` or
            `find_default_brex_fallback`

    Returns:
        str: absolute path to the bundled BREX XML file, or None if
            `logical_dmc` is not one of the built-in default BREX
    """
    entry = _DEFAULT_BREX.get(logical_dmc)
    if entry is None:
        return None
    return join(DEFAULT_BREX_DIR, entry[0])


def find_default_brex_fallback(ref: dict) -> str:
    """Check whether a BREX reference that could not be found on disk names
    one of the built-in default BREX, so it can be used as a fallback.

    Port of `search_brex_fname_from_default_brex` (`s1kd-brexcheck.c:369-380`):
    the reference matches if its DMC (ignoring issue/inWork/language) is one
    of the seven built-in default BREX, and either the reference does not
    specify an issue (accepted regardless of which issue is bundled, as when
    a data module names its BREX without an `issueInfo`/`issno`) or it
    specifies exactly the bundled issue and inWork.

    Args:
        ref (dict): BREX reference dict, as returned by `s1000d.get_brex_ref`

    Returns:
        str: the matched logical DMC (usable with `default_brex_path`), or
            None if the reference does not name a built-in default BREX
    """
    base_ref = dict(ref)
    base_ref["issueNumber"] = ""
    base_ref["inWork"] = ""
    try:
        base_dmc = ref_dict_to_str(base_ref)
    except KeyError:
        return None

    entry = _DEFAULT_BREX.get(base_dmc)
    if entry is None:
        return None

    _, issue_number, in_work = entry
    ref_issue = ref.get("issueNumber") or ""
    ref_in_work = ref.get("inWork") or ""
    if ref_issue and (ref_issue != issue_number or ref_in_work != in_work):
        return None
    return base_dmc
