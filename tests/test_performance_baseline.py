"""Performance baseline for brex_checker_rework.md §4.5, protecting the P3
caching work (§4.4: parse each BREX once per run rather than once per
document, and compile each objectPath once) over the real evidence folder.

Marked `evidence` (deselect with `pytest -m "not evidence"` for a fast local
loop): a real, un-mocked run over this real corpus (67 real CSDB objects
checked against a layered chain drawn from the same 919-rule, 3-BREX
evidence base as the rest of the plan document) takes on the order of a
minute or more even with caching working correctly -- it is measuring real
performance, not a synthetic fixture, so it cannot be made instant. Skipped
entirely when the evidence corpus is not present on disk; override its
location with ACD_BREX_EVIDENCE_DIR.
"""

import os
import time
from os.path import isdir, isfile, join
from unittest.mock import patch

import pytest

from acd.brex_checker import BrexChecker

EVIDENCE_DIR = os.environ.get(
    "ACD_BREX_EVIDENCE_DIR",
    r"C:\Users\munte\Develop\TD\SITEC\Seventh Delivery\CMP 21-77-05",
)

pytestmark = [
    pytest.mark.evidence,
    pytest.mark.skipif(
        not isdir(EVIDENCE_DIR),
        reason="real CMP 21-77-05 evidence folder not available on this machine "
               "(set ACD_BREX_EVIDENCE_DIR to point at it)",
    ),
]


def _xml_files() -> list:
    return [
        name for name in os.listdir(EVIDENCE_DIR)
        if name.lower().endswith(".xml") and isfile(join(EVIDENCE_DIR, name))
    ]


def test_each_brex_schema_pair_is_parsed_at_most_once_per_run():
    # The real regression guard for §4.4's caching work. Without it,
    # _get_object_rule_nodes (the expensive BREX parse + rule-node
    # selection step -- extracting up to 919 rules from a 300+KB file) runs
    # once per (document, brex-in-its-resolved-chain) pair: on this real
    # folder, with chains 1-3 layers deep, well over one hundred calls.
    # With caching (_get_content_rules), it runs at most once per distinct
    # (brex, schema) pair actually encountered across the whole run --
    # necessarily fewer than the number of documents checked, since many
    # documents share both the same BREX chain and the same schema
    # (observed 29 distinct pairs across 67 real documents on the reference
    # evidence folder).
    total_documents = len(_xml_files())
    assert total_documents > 0

    checker = BrexChecker()
    checker.set_xml_dir(EVIDENCE_DIR)
    with patch.object(checker, '_get_object_rule_nodes',
                       wraps=checker._get_object_rule_nodes) as spy:
        checker.validate()

    assert 0 < spy.call_count < total_documents


def test_full_directory_run_completes_within_a_generous_ceiling():
    # A loose smoke ceiling, not a tight benchmark -- dev-machine speed
    # varies and this is real XPath evaluation work, not something to
    # micro-optimise for in a test. Its purpose is to catch a gross
    # regression (e.g. the caching added in §4.4 silently stops working,
    # or a rule/document loop becomes accidentally quadratic), not to
    # enforce a specific number. Reference observation on the dev machine
    # this plan was authored on: ~90 seconds for the full folder.
    checker = BrexChecker()
    checker.set_xml_dir(EVIDENCE_DIR)

    started = time.perf_counter()
    results = checker.validate()
    elapsed = time.perf_counter() - started

    assert len(results) == len(_xml_files())
    assert elapsed < 600, (
        f"directory-mode validate() over the evidence folder took {elapsed:.1f}s "
        "(generous ceiling 600s) -- possible caching regression, see §4.4"
    )
