# BREX Checker — Rework: evidence, taxonomy and defects

The tasks of this rework (N1–N80, all done) are in `tasklist.json` in the repository root; their `group` is the §4 heading they came from. This file keeps the evidence, the analysis and the order they were worked in.

Target file: `acd/brex_checker.py`
Reference implementation: `s1kd-brexcheck` (`D:\Downloads\s1kd-tools-master\tools\s1kd-brexcheck\s1kd-brexcheck.c`, value/range/pattern helpers in `tools/common/s1kd_tools.c`)

## 1. Evidence base

Everything below was derived by parsing three real, complex BREX data modules and replaying our checker's own logic against them and against the 63 non-BREX CSDB objects in the same folder (`C:\Users\munte\Develop\TD\SITEC\Seventh Delivery\CMP 21-77-05`).

| BREX | `structureObjectRule` | `contextRules` | `objectValue` | Other rule containers |
|---|---|---|---|---|
| `DMC-S1000D-F-04-10-0301-00A-022A-D_001-00_EN-US.XML` (S1000D 4.2 default BREX) | 260 | 6 | 1347 | `nonContextRules` (1), `brDecisionRef` (261) |
| `DMC-ATABREX-F-00-00-00-00A-022A-D_004-00_EN-US.XML` (ATA civil aviation BREX) | 192 | 25 | 990 | `brDecisionRef` (1) |
| `DMC-ATABREX-F-00-00-00-01A-022A-D_004-00_EN-US.XML` (ATA CMP BREX) | 467 | 29 | 125 | `snsRules` (749 `snsCode`) |
| **Total** | **919** | **60** | **2462** | |

Rule population breakdown:

- `allowedObjectFlag`: `0` = 657, `1` = 46, `2` = 216
- `objectValue/@valueForm`: `single` = 2349, `pattern` = 40, `range` = 73
- `objectValue/@valueTailoring`: `lexical` = 1440, `restrictable` = 829, absent = 193
- `contextRules/@rulesContext`: 26 distinct target schemas plus unqualified. 476 rules sit in unqualified groups (apply to every object type), 443 rules are schema-qualified.
- `brSeverityLevel` / `defaultBrSeverityLevel`: **0 occurrences** in these three, but part of the schema.
- SNS depth used: `snsSystem` (78) → `snsSubSystem` (391) → `snsSubSubSystem` (280). No `snsAssy`.
- No `notationRule` in these three.

XPath engine comparison over all 919 `objectPath` expressions:

- `elementpath` (XPath 2.0, what we use): **919 / 919 compiled and evaluated**.
- `libxml2` XPath 1.0 (what a default `s1kd-brexcheck` build uses): 913 / 919 — 6 fail on `matches()`, `tokenize()` and `lower-case()`.

**Our XPath engine choice is better than the s1kd default.** Almost every gap below is in how we *select*, *interpret* and *report* rules, not in how we evaluate XPath.

### 1.1 Headline reproduction

A minimal BREX with 10 deliberately-violated rules (5 x flag 0, 2 x flag 1, 3 x flag 2) run through the current `BrexChecker.validate()` returns:

```
"Summary": "3 Errors"
```

**3 of 10 violations reported.** Root causes, in order of impact: the error-key collision (§3.1), boolean-valued flag-0 rules being dropped (§3.2), and partial-match pattern semantics (§3.4).

---

## 2. Rule taxonomy

### 2.A Structural categories — the kinds of rule a BREX can carry

| # | Category | BREX element | In the 3 BREX | `brex_checker.py` | `s1kd-brexcheck` |
|---|---|---|---|---|---|
| A1 | Context rules (structure/object rules) | `contextRules//structureObjectRule` | 919 | Yes (fragile selection) | Yes |
| A2 | SNS rules | `snsRules/snsSystem/…` | 749 codes | **No** | Yes (`-S`, `-St`, `-Su`) |
| A3 | Notation rules | `notationRuleList/notationRule` | 0 (schema-legal) | **No** | Yes (`-n`) |
| A4 | Non-context rules (human-readable BRs) | `nonContextRules/nonContextRule` | 1 | **No** | **No** |
| A5 | Business-rule decision refs | `brDecisionRef/@brDecisionIdentNumber` | 262 | **No** | Yes (reported per error) |
| A6 | Severity levels | `@brSeverityLevel`, `brex/@defaultBrSeverityLevel`, `.brseveritylevels` | 0 | **No** | Yes (`-w`) |
| A7 | Rule context targeting | `contextRules/@rulesContext` | 60 groups | Yes (broken extractor, §3.3/§3.6) | Yes |
| A8 | Layered BREX | `brexDmRef` chain | 1 chain | Partial (§3.7) | Yes (`-l`, `-B`, `-I`, `-r`) |

### 2.B Semantic categories — `objectPath/@allowedObjectFlag`

| # | Flag | Meaning | Count | Our status |
|---|---|---|---|---|
| B1 | `0` | Object **must not** be present / condition must not hold | 657 | Partial — node results OK, boolean results dropped |
| B2 | `1` | Object **must** be present / condition must hold | 46 | Mostly OK, falsy-non-boolean edge case |
| B3 | `2` | Object present, but its **value is constrained** | 216 | Partial — several value-semantics defects |

### 2.C Value-constraint categories — `objectValue/@valueForm`

| # | Form | Semantics (S1000D / libxml2) | Count | Our status |
|---|---|---|---|---|
| C1 | `single` (or absent) | Exact string equality | 2349 | OK for attribute/text results; broken for element results |
| C2 | `pattern` | **XSD regular expression, whole-value match** | 40 | Broken — partial match plus wrong regex dialect |
| C3 | `range` | S1000D range `a~c` **and set `a\|b\|c`**, numeric-then-lexicographic | 73 | Broken — integer-only enumeration, no set support |
| C4 | `valueTailoring` (`lexical` / `restrictable`) | Tailoring policy metadata | 2269 | Ignored (also ignored by s1kd) |
| C5 | S1000D <= 3.0 legacy `@val1`/`@val2`, `objval`, `objpath`, `objuse`, `@objappl` | Old spelling | 0 | **No** (s1kd: yes) |

### 2.D Reporting / metadata categories

| # | Category | Our status | s1kd |
|---|---|---|---|
| D1 | Violating node line number | Partial, heuristic text search | Yes (`xmlGetLineNo`) |
| D2 | Canonical XPath of the violating node | **No** | Yes (`xpath_of`) |
| D3 | Copy of the violating XML subtree | **No** | Yes (`-8` for deep copy) |
| D4 | `objectUse` message | Yes (`.text` only) | Yes (full content) |
| D5 | Allowed-value list echoed into the report | Yes | Yes |
| D6 | Report formats | Yes (`to_xml_report` `-x`-compatible, `to_json_report`, `to_excel_report`, `to_html_report`) | Yes (`-x`) |
| D7 | Per-run summary / statistics | Yes (`run_summary()`, plus `<summary>` in `to_xml_report`) | Yes (`-T`) |
| D8 | Invalid-`objectPath` diagnostics | **No** — aborts the run | Yes (`<xpathError>`, continues) |

### 2.E Subject-matter categories — what the 919 rules are actually about

Derived by classifying every `objectPath` expression:

| Category | Rules |
|---|---|
| Identification and status (`dmIdent`, `dmCode`, `issueInfo`, `dmAddress`, `dmStatus`) | 146 |
| References (`dmRef`, `pmRef`, `externalPubRef`, `internalRef`, `infoEntityRef`) | 105 |
| Controlled-value attributes (`@*Type`, `@*Category`, `@*Class`, `@*Flag`) | 102 |
| Applicability and product attributes (`applic`, `assert`, `evaluate`, `productAttribute`) | 94 |
| Illustrated parts data (`catalogSeqNumber`, `itemSeqNumber`, `partSegment`) | 52 |
| Tables, figures and graphics (`table`, `tgroup`, `figure`, `graphic`, `hotspot`) | 43 |
| Description and common text (`description`, `levelledPara`, `para`, lists, `footnote`) | 43 |
| Units of measure / quantities / data values | 29 |
| Procedural content (`proceduralStep`, `preliminaryRqmts`, `reqCondGroup`, …) | 25 |
| Security, data restrictions and copyright | 18 |
| Fault / schedule / process / crew / service bulletin | 16 |
| Warnings, cautions and notes | 14 |
| Change and update control (`reasonForUpdate`, `@changeType`, `@changeMark`) | 11 |
| Front matter and publication module | 9 |
| Quality assurance and responsible partner | 9 |
| SNS / numbering codes inside `objectPath` | 8 |
| Data management list / DDN / comment | 5 |
| Remaining element-specific constraints (`sbTopic`, `supportEquipDescr`, `taskDuration`, `circuitBreakerRef`, `unitOfIssueQualificationSegment`, …) | 190 |

**All 19 subject-matter categories are handled by the same generic XPath engine.** They need no per-category code — they need the engine, the value semantics and the reporting to be correct. That is what the task list targets.

---

## 3. Defects found in `brex_checker.py`

Each defect below is reproduced, not inferred.

### 3.1 Error keys collide — only the last violation of each flag survives (CRITICAL)

`_check_rules` does `error_0, error_1, error_2 = 1, 1, 1` **once**, then passes those ints by value into `_check_object_flag_0/1/2` for **every rule**. The `error_0 += 1` inside each helper mutates a local copy. So each rule starts numbering at `1` again and `brex_violations[brex]['0'] |= {1: {...}}` overwrites the previous rule's entry.

Reproduced: 5 flag-0 violations → 1 reported; 2 flag-1 → 1 reported; 3 flag-2 → 1 reported.

### 3.2 Boolean-valued `objectPath` results are silently discarded (flag 0)

`_check_object_flag_0` guards with `if type(selector.select(root)) is not bool:`. Any rule whose XPath returns a boolean (`A and B`, `count(...) > 0`, `//x/@y = "z" and ...`) is skipped entirely.

s1kd handles this explicitly (`is_invalid`): empty node-set → `invalid = obj->boolval`.

**24 of the 657 flag-0 rules in these three BREX return a boolean** — 3.7% of the largest rule class, never evaluated.

### 3.3 `get_schema_from_xml` is attribute-order dependent (HIGH)

`xml_processing.get_schema_from_xml` uses `(xsi:noNamespaceSchemaLocation=")(.*?)(">)`, which requires `xsi:noNamespaceSchemaLocation` to be the **last** attribute on the root element. When it is not, the capture runs on to the next `">` in the document.

Verified against the delivery: `DMC-S1000D-F-04-10-0301-00A-022A-D_001-00_EN-US.XML` yields `…brex.xsd" xmlns:dc="http://www.purl.org/dc/elements/1.1/" xmlns:rdf=…` instead of the schema URI.

Consequence: `value['contextRules'] == schema` never matches, so **only the unqualified rule groups run** — 443 of 919 rules (48%) are silently skipped for any object whose root element does not put that attribute last. Many authoring tools emit it first.

### 3.4 `pattern` values are matched partially and in the wrong regex dialect (HIGH)

- S1000D / `xmlRegexpExec` semantics: the **whole value** must match. We use `regex.search`, so a substring match passes. Verified false negative: value `XX12XX` passes pattern `[0-9]{2}`.
- XSD character-class subtraction is used verbatim in these BREX: `00[A-Z-[IO]]{1,3}|00[0]{1}` and `^[-A-Z0-9-[O]]{1,15}$`. Python's `regex` module (even with `V1`) does not interpret `[A-Z-[IO]]` as subtraction. Verified: `00IOI` is accepted although XSD excludes `I` and `O`.
- Other XSD-only constructs to translate: `\i`, `\c`, `\I`, `\C`, `\p{IsBasicLatin}`-style block escapes, and the implicit `^(?:…)$` anchoring.

### 3.5 `range` handling is integer-only and has no set support (HIGH)

`_show_rules` expands `range` with `findall(r"([a-z]*)(\d+)", …)` then counts integers. Verified:

| `valueAllowed` | Our result | Correct |
|---|---|---|
| `accpnl51~accpnl99` | `accpnl51 … accpnl99` | OK |
| `a~c` | **IndexError (crash)** | `a`, `b`, `c` |
| `A\|B\|C` | **IndexError (crash)** | set `A`, `B`, `C` |
| `01\|02` | `['1','2']` | set `01`, `02` |
| `aa01~aa09` | `['aa1' … 'aa9']` | `aa01 … aa09` |
| `0001~0099` | `['1' … '99']` | `0001 … 0099` |

s1kd's `is_in_set` / `is_in_range` splits on `|`, then on `~`, and compares **numerically when both ends and the value parse as numbers, lexicographically otherwise** — no enumeration at all.

The 73 ranges in these three BREX happen to be of the safe `xx51~xx99` shape, so this is currently latent — but it crashes or mis-validates on any BREX using alphabetic ranges, sets, or zero padding.

### 3.6 Rule selection is structurally over-specific

`_get_object_rule_nodes` uses `findall('//structureObjectRuleGroup/structureObjectRule/objectPath')` and `_show_rules` walks back up with `getparent().getparent().getparent()` to reach `contextRules`.

- Requires exactly `contextRules / structureObjectRuleGroup / structureObjectRule`. Rules placed directly under `contextRules`, or in nested groups, are missed (and the parent walk would read the wrong element's `@rulesContext`).
- lxml emits `FutureWarning: This search incorrectly ignores the root element` for the leading `//`.
- s1kd uses the descendant axis and filters at selection time: `//contextRules[not(@rulesContext) or @rulesContext=$schema]//structureObjectRule`.

On these three BREX both approaches return the same 919 rules, so this is a robustness fix.

### 3.7 BREX resolution does not use `brexDmRef`

`s1000d.get_brex_ref` returns **the first `dmRef` anywhere in the document whose `infoCode` is `022`**; `brexDmRef` is not referenced anywhere in the codebase. Any `dmRef` to a BREX DM in the content will hijack resolution. s1kd uses `//brexDmRef|//brexref`.

Also: `_init_brex_list` terminates the layering walk with `if ref_dict_to_str(get_brex_ref(xml)) in xml` (a substring test against the *file path*), has no visited-set cycle guard, and `ref_dict_to_str(get_brex_ref(...))` raises `TypeError` when `get_brex_ref` returns `None`.

### 3.8 Value checking is only applied to flag `2`

s1kd applies `objectValue` checking to **any** rule carrying `objectValue` children (a flag-`1` rule that matched can still fail on value). We only check flag `2`. No practical difference in these three BREX (all 46 flag-1 rules are value-less) but it is a semantic gap.

### 3.9 Element-valued flag-2 results are compared against strings

In `_check_object_flag_2`, `element` may be an lxml `_Element` (elementpath returns elements for element-selecting paths, strings for attribute/text paths). `element not in value["values_allowed"]` is then always true, `search(regex, element)` raises `TypeError`, and the fallback

```python
regex2 = search(r"(@)([a-zA-Z]+)(^[a-zA-Z])", value['xpath'], V1)
... element.attrib[regex2.group(2)] ...
```

can never match (the `^` anchor sits mid-pattern), so `regex2` is `None` and `regex2.group(2)` raises an uncaught `AttributeError`.

Latent in these three BREX: all 18 element-returning flag-2 rules happen to have no `objectValue` children and are skipped before reaching this code.

### 3.10 DDN objects are excluded from every check

`_check_rules` wraps all three flag branches in `if schema != "http://www.s1000d.org/S1000D_4-2/xml_schema_flat/ddn.xsd"`. The three BREX carry **22 rules with `rulesContext="…/ddn.xsd"`** that can therefore never fire. The exclusion is also hard-coded to 4.2 and would not match a 4.1 or 5.0 DDN.

### 3.11 Directory mode defects

- `validate()` filters with `".xml" in _.lower()` — matches `foo.xml.bak`; it is not an extension test.
- `"-022a-" not in _.lower()` means **BREX DMs are never validated**, including against themselves. s1kd validates BREX objects like any other.
- `result` is only ever the **last** file's result; the function returns it and the per-file results are discarded unless `debug=True`.
- If the directory contains no matching file, `result` is unbound, raising `NameError`.
- The `debug` JSON is assembled by writing `{`, appending `dump(...)` per file, then `}` — producing invalid JSON for more than one file (no separating commas).
- `self._brex_list` / `self._brex_dir_path` are reset per file, discarding any `override_brex_list()`.

### 3.12 Robustness / hygiene

- No `try/except` around `elementpath.Selector(...)`; a single malformed `objectPath` aborts the whole run. s1kd records `<xpathError>` and continues.
- Namespaces are hard-coded to `rdf` + `xsi`. s1kd registers **every namespace in scope at the `objectPath` node** (`dc`, `xlink`, … are also declared in these BREX).
- No EXSLT `date:` / `math:` / `set:` / `str:` support (s1kd registers all four). Unused in these three BREX but common in tailored ones.
- `objectUse` is read with `.text`, dropping content after any child element, and `[0]` raises `IndexError` when `objectUse` is absent.
- `_check_rules` re-reads the XML into `self.xml_content` while the flag helpers use `self._xml_content` — two attributes for one thing.
- `_show_rules` is re-run (re-parsing every BREX, rebuilding the whole rule list) for **every** data module; `elementpath.Selector` is recompiled per (rule, document) pair.
- `_check_object_flag_1` calls `selector.select(root)` twice per rule.
- The Saxon path exists only in `_check_object_flag_0`; `_check_object_flag_1` / `_2` ignore `self._saxon`.
- Line numbers for attribute results are found by scanning raw text for the attribute name, reporting the **last** matching line rather than the violating node's.
- `regex_builder` builds a regex from unescaped attribute values; a value containing regex metacharacters produces a wrong or invalid pattern.

---

## 4. Task list

Moved to `tasklist.json`. Each task keeps its full text in `details`; `group` names the subsection below.

### 4.1 P0 — Correctness bugs that make current output wrong

### 4.2 P1 — Rule categories we do not check but `s1kd-brexcheck` does

### 4.3 P2 — Rule categories neither tool checks

### 4.4 P3 — Improvements to checks we already perform

### 4.5 P4 — Test suite

### 4.6 P0/P1 — Rework found by auditing §4.1–§4.5 against the evidence base

Every task in this group was reproduced against `acd/brex_checker.py` as it stands after §4.1–§4.5, either on the real `CMP 21-77-05` corpus or with a minimal fixture. The full suite (347 unit tests + 12 evidence tests) passes as-is, so none of these is caught today.

---

## 5. Suggested order

1. §4.1 in full — without it every downstream number is wrong (3 of 10 violations reported today).
2. §4.5 fixtures for the P0 items, so the fixes stay fixed.
3. SNS rules, `brDecisionRef` and severity levels from §4.2 — the three categories a customer will notice missing first.
4. The structured-violation refactor in §4.4, then the XML report and the differential harness.
5. §4.3 as differentiators once parity with `s1kd-brexcheck` is established and measured.
6. §4.6 — §4.1–§4.5 are complete and green; these are the defects that audit surfaced. Take the `Decimal` crash and the line-number heuristics first (both affect output that is wrong today on real data), then the SNS/notation gating, then the remaining three.
