"""Layered-BREX fixture for brex_checker_rework.md §4.5, walking the real
chain DM -> ATABREX 01A -> ATABREX 00A -> S1000D default (issue F) named in
the plan document, and asserting every layer is collected exactly once.

Uses the real evidence-base DM `DMC-CTTAE29N-A-00-00-00-01A-00KA-D_001-01_SX-US.XML`
and the real BREX files (see the §4.2 "Use brexDmRef for BREX resolution"
task note in brex_checker_rework.md, which independently verified this exact
chain). All three BREX layers -- both ATABREX modules and the S1000D 4.2
default BREX itself -- are physically present in the evidence folder (they
are the same three BREX the plan document's whole evidence base in §1 was
built from), so this chain resolves entirely from real, on-disk files with
no built-in-default fallback involved; `test_default_brex.py` already covers
the fallback path itself with our bundled `acd/brex/` copies.

A synthetic cyclic-reference fixture already exists:
test_brex_dm_ref_resolution.py::test_init_brex_list_cycle_guard_prevents_infinite_loop.

Skipped when the evidence corpus is not present on disk; override its
location with ACD_BREX_EVIDENCE_DIR.
"""

import os
from os.path import isfile, join

import pytest

from acd.brex_checker import BrexChecker

EVIDENCE_DIR = os.environ.get(
    "ACD_BREX_EVIDENCE_DIR",
    r"C:\Users\munte\Develop\TD\SITEC\Seventh Delivery\CMP 21-77-05",
)
REAL_DM = join(EVIDENCE_DIR, "DMC-CTTAE29N-A-00-00-00-01A-00KA-D_001-01_SX-US.XML")
ATABREX_01A = join(EVIDENCE_DIR, "DMC-ATABREX-F-00-00-00-01A-022A-D_004-00_EN-US.XML")
ATABREX_00A = join(EVIDENCE_DIR, "DMC-ATABREX-F-00-00-00-00A-022A-D_004-00_EN-US.XML")
S1000D_DEFAULT = join(EVIDENCE_DIR, "DMC-S1000D-F-04-10-0301-00A-022A-D_001-00_EN-US.XML")

pytestmark = [
    pytest.mark.evidence,
    pytest.mark.skipif(
        not (isfile(REAL_DM) and isfile(ATABREX_01A) and isfile(ATABREX_00A) and isfile(S1000D_DEFAULT)),
        reason="real CMP 21-77-05 evidence folder not available on this machine "
               "(set ACD_BREX_EVIDENCE_DIR to point at it)",
    ),
]


@pytest.fixture
def checker():
    checker = BrexChecker()
    checker.set_xml(REAL_DM)
    checker.set_brex_path(EVIDENCE_DIR)
    return checker


def test_chain_resolves_dm_through_both_atabrex_layers_to_the_s1000d_default(checker):
    checker._init_brex_list()

    assert checker._brex_list[0] == [ATABREX_01A, ATABREX_00A, S1000D_DEFAULT]


def test_every_layer_is_collected_exactly_once(checker):
    checker._init_brex_list()

    chain = checker._brex_list[0]
    assert len(chain) == len(set(chain))
    assert len(chain) == 3


def test_no_built_in_brex_fallback_is_needed_for_this_real_chain(checker):
    # All three layers are physically present in the evidence folder, so
    # resolution succeeds without ever falling back to our bundled
    # acd/brex/ copies -- test_default_brex.py covers that fallback path
    # directly, with our own bundled files.
    checker._init_brex_list()

    assert checker._brex_fallbacks == []


def test_validate_checks_the_real_dm_against_all_three_real_layers_without_raising(checker):
    result = checker.validate()

    assert set(result.keys()) >= {ATABREX_01A, ATABREX_00A, S1000D_DEFAULT}
    for brex_path in (ATABREX_01A, ATABREX_00A, S1000D_DEFAULT):
        assert set(result[brex_path].keys()) >= {"0", "1", "2", "xpathError"}
