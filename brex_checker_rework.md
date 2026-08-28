# BREX Checker — Rework Plan

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
| D6 | XML report format | **No** (ad-hoc JSON) | Yes (`-x`) |
| D7 | Per-run summary / statistics | Counts a collapsed dict, so wrong | Yes (`-T`) |
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

### 4.1 P0 — Correctness bugs that make current output wrong

- [x] **Fix the error-key collision.** Replace the by-value `error_0/1/2` counters with a per-BREX, per-flag accumulating collection (a `list` of violation records, or a counter held on `self`). Ref §3.1.
- [x] **Handle boolean XPath results for flag `0`.** Mirror s1kd `is_invalid`: empty node-set → violation iff the boolean result is true; non-empty node-set → violation. Ref §3.2.
- [x] **Harden flag `1` result interpretation.** Distinguish `bool` / node-set / string / number results explicitly instead of relying on Python truthiness, and evaluate the selector once. Ref §3.2, §3.12.
- [x] **Replace `get_schema_from_xml` with a parsed lookup** of `/*/@xsi:noNamespaceSchemaLocation` (lxml, namespace-aware) instead of the order-dependent regex. Ref §3.3.
- [x] **Add a regression test** asserting the schema is read correctly when `xsi:noNamespaceSchemaLocation` is the first, middle and last attribute.
- [x] **Make `pattern` matching whole-value.** Use `regex.fullmatch`, or wrap the pattern as `^(?:…)$`. Ref §3.4.
- [x] **Translate XSD regex syntax to Python before compiling.** At minimum character-class subtraction `[A-Z-[IO]]` → `[A-Z--[IO]]` (regex `V1` set operations), plus `\i`, `\c`, `\I`, `\C` and XSD block/category escapes. Both ATABREX modules use subtraction today. Ref §3.4.
- [x] **Add a test** covering `00[A-Z-[IO]]{1,3}|00[0]{1}` accepting `00ABC` and rejecting `00IOI`, and `[0-9]{2}` rejecting `XX12XX`.
- [x] **Reimplement `range` as `is_in_set` / `is_in_range`** (port `s1kd_tools.c:378-434`): split on `|`, then on `~`, compare numerically when both bounds and the value are numeric, otherwise lexicographically. Stop enumerating values. Ref §3.5.
- [x] **Add a test matrix for range/set values**: `a~c`, `A|B|C`, `01|02`, `aa01~aa09`, `0001~0099`, `20~100` (numeric vs lexicographic), `accpnl51~accpnl99`.
- [x] **Normalise flag-2 result items to their string value** before comparison (element → text content; attribute → value), removing the broken `TypeError` fallback and the unreachable `(@)([a-zA-Z]+)(^[a-zA-Z])` regex. Ref §3.9.
- [x] **Wrap `elementpath.Selector` construction and evaluation in `try/except`**, record the failure as an `xpathError` entry against that rule, and continue with the remaining rules. Ref §3.12.
- [x] **Remove the hard-coded DDN exclusion** and let `rulesContext` matching decide; the three BREX contain 22 DDN rules that currently never run. Ref §3.10.
- [x] **Fix directory mode**: return a mapping of `{filename: result}`, handle the empty-directory case, emit valid JSON in `debug` mode, and stop wiping an explicit `override_brex_list()`. Ref §3.11.
- [x] **Fix `_append_summary`** to count actual violations (it currently measures the collapsed dict). Ref §3.1, §3.11.
- [x] **Resolve `self.xml_content` vs `self._xml_content`** to a single attribute. Ref §3.12.

### 4.2 P1 — Rule categories we do not check but `s1kd-brexcheck` does

- [x] **Implement SNS rule checking (category A2).** Port `check_brex_sns_rules` (`s1kd-brexcheck.c:1057-1144`): walk `systemCode` → `subSystemCode` → `subSubSystemCode` → `assyCode` down `snsSystem` / `snsSubSystem` / `snsSubSubSystem` / `snsAssy`, stopping at the first failing level. Needed for `DMC-ATABREX-F-00-00-00-01A-022A-D_004-00_EN-US.XML` (749 SNS codes). Implemented as `BrexChecker._check_sns_rules` / `_sns_should_check` / `_get_sns_rules_group` in `acd/brex_checker.py`; results land in `result["sns"]` and feed `_append_summary`. Verified against the real ATA CMP BREX: parsed SNS counts (78 `snsSystem` / 391 `snsSubSystem` / 280 `snsSubSubSystem` / 0 `snsAssy`) match the evidence base exactly, and end-to-end runs against the 62 CMP 21-77-05 data modules produce plausible pass/fail results.
- [x] **Implement the three SNS modes** — normal (optional levels default to `0` / `00` / `0000`), `strict` (no shorthand), `unstrict` (any code valid when the level is omitted); port `should_check` (`s1kd-brexcheck.c:1038`). Expose as a parameter on `validate()`. Implemented as a `sns_mode` parameter (`"normal"` / `"strict"` / `"unstrict"`, validated against `SNS_MODES`) threaded through `validate()` -> `_check_rules` -> `_check_sns_rules` -> `_sns_should_check` in `acd/brex_checker.py`.
- [x] **Combine `snsRules` from all BREX in the layer chain** into one rule set before checking, as s1kd's `check_brex_sns` does. Done in `_get_sns_rules_group`.
- [x] **Restrict SNS checking to `dmodule` roots** (s1kd skips PM / DML / DDN / comment). Done in `_check_sns_rules` (checks `dmod_root.tag == "dmodule"`).
- [x] **Implement notation rule checking (category A3).** Port `check_brex_notation_rules` / `check_entity`: read `NOTATION` declarations from the internal DTD subset, accept a notation if some `notationRule` names it with `@allowedNotationFlag != "0"`, and report the `objectUse` of the first inclusion rule otherwise. Requires parsing with DTD loading enabled. Implemented as `BrexChecker._check_notation_rules` / `_check_entity_notation` / `_get_notation_rules_group` in `acd/brex_checker.py`; results land in `result["notations"]` and feed `_append_summary`. Follows the C original in walking `ENTITY` declarations of the internal DTD subset (unparsed/`NDATA` entities specifically) rather than `<!NOTATION>` declarations directly, since that is what `check_entity`'s `entity->content` actually reads; the internal DTD subset is available from lxml's default parser with no special DTD-loading options needed (it is inline in the document, unlike an external subset).
- [x] **Capture and report `brDecisionRef/@brDecisionIdentNumber` (category A5)** on every violation — 262 rules in these three BREX carry one, and it is the identifier customers quote. Implemented in `acd/brex_checker.py`: `_show_rules` reads the `brDecisionRef` sibling of `objectPath`/`objectUse` (confirmed against the real ATABREX evidence file, e.g. `SOR-230` at line 773 of `DMC-ATABREX-F-00-00-00-00A-022A-D_004-00_EN-US.XML`) and carries its `brDecisionIdentNumber` (or `None` when absent) into every rule dict; `_check_object_flag_0/1/2` add it to each violation record under `BrDecisionIdentNumber`, and the three `xpathError` diagnostics carry it too so a malformed rule can still be traced back to its decision number. Matches the C original, which only wires `brDecisionRef` into `check_brex_rules` (structure-object-rule checks) — SNS and notation-rule violations have no associated `brDecisionRef` in the schema and are unaffected. Covered by `tests/test_br_decision_ref.py`.
- [x] **Implement business-rule severity levels (category A6).** Read `structureObjectRule/@brSeverityLevel`, fall back to `brex/@defaultBrSeverityLevel`, resolve against a `.brseveritylevels` file (`<brSeverityLevels><brSeverityLevel value="brsl01" fail="yes">Error</brSeverityLevel>…`), and let `fail="no"` levels report as warnings that do not fail the run. Implemented in `acd/brex_checker.py`: `_show_rules` reads `structureObjectRule/@brSeverityLevel` (falling back to the `brex` root's `@defaultBrSeverityLevel` when absent) into each rule dict's `brSeverityLevel` key; `set_severity_levels_path` / `_get_severity_levels` load and cache a `.brseveritylevels` file into a `{value: {fail, type}}` table; `_is_severity_failure` ports `is_failure` (`s1kd-brexcheck.c:569-605`) — a violation fails unless a severity-levels file is set, defines that exact value, and marks it `fail="no"`. `_check_object_flag_0/1/2` attach `BrSeverityLevel` and `Fail` to every violation record, and `_append_summary` excludes `Fail: False` entries from the error count, reporting them instead as `"N Errors, M Warnings"` (unchanged `"N Errors"` format when there are no warnings, so existing callers/tests are unaffected). No explicit path defaults to the old always-fail behaviour, matching the C original when no `-w` flag is given. Covered by `tests/test_severity_levels.py`. Parent-directory auto-discovery of `.brseveritylevels` is the next task item, not part of this one.
- [x] **Search parent directories for `.brseveritylevels`** by default, with an override parameter. Implemented in `acd/brex_checker.py`: `_find_severity_levels_file` walks from the checked XML's directory upward through its parents (adapting `find_config`, `s1kd_tools.c:30-56`, to search from the file's directory instead of the process's cwd) looking for a file named `.brseveritylevels`, returning the first match or `None` at the filesystem root. `_get_severity_levels` now falls back to this search whenever `set_severity_levels_path` has not been called explicitly. Two override paths: `set_severity_levels_path` (explicit file, unchanged, takes precedence and skips the search) and the new `set_severity_levels_search(enabled: bool)` (turns the default search off entirely, e.g. to ignore a discoverable file). Covered by new cases in `tests/test_severity_levels.py`: discovery in the XML's own directory, discovery in an ancestor directory, explicit-path precedence over an auto-discovered file, and search disabled ignoring a discoverable file; the two pre-existing "no severity levels" tests now call `set_severity_levels_search(False)` so they no longer depend on the ambient filesystem lacking a stray `.brseveritylevels` above the pytest temp dir.
- [x] **Apply `objectValue` checking to every rule that has `objectValue` children**, not only flag `2` (s1kd `is_invalid` → `check_objects_values`). Ref §3.8. Implemented in `acd/brex_checker.py`: extracted the per-element value-matching loop out of `_check_object_flag_2` into a shared `_check_object_values(value, elements)` helper (port of `check_objects_values`, `s1kd-brexcheck.c:275-304`), then wired it into `_check_object_flag_1` — when a flag-`1` rule's node-set is non-empty (the presence check passed) and the rule carries `values_allowed`/`regex_allowed`/`ranges_allowed`, its matched nodes are now also value-checked, with any failures appended to the `'2'` violation bucket. Flag `0` is intentionally left unchanged: per the C reference, `check_objects_values` is only reached when the node-set is non-empty, and a non-violating flag-`0` rule always has an *empty* node-set (a non-empty one is already a presence violation), so there is nothing to value-check there.
- [x] **Use `brexDmRef` for BREX resolution.** Add a dedicated lookup for `//brexDmRef|//brexref` instead of "first `dmRef` with `infoCode` 022". Ref §3.7. Implemented in `acd/s1000d.py`: `get_brex_ref` now parses the object and runs the XPath `//brexDmRef|//brexref` (matching `s1kd-brexcheck`'s `find_brex_fname_from_doc`) to get the one node that actually carries the applicable-BREX reference, then reads its nested `.//dmCode|.//avee` (plus `.//issueInfo` for the 4.0+ form) instead of scanning every `dmRef`/`refdm` in the document for `infoCode`/`incode` `022`. Verified against the real ATA CMP evidence base (`DMC-CTTAE29N-A-00-00-00-01A-00KA-D...` → `DMC-ATABREX-...-01A-...` → `DMC-ATABREX-...-00A-...` → `DMC-S1000D-F-04-10-0301-00A-022A-D`, the last self-referencing to terminate the chain) — identical, correct resolution before and after the change. Covered by `tests/test_brex_dm_ref_resolution.py`, including a case with a decoy `dmRef` (`infoCode="022"`, pointing at a nonexistent DM) placed before the real `brexDmRef` in document order: confirmed by inlining the old scan-based logic that it would have followed the decoy and raised `NoBrexDefined`, while the new lookup resolves and checks the real BREX. The legacy S1000D <= 3.0 `brexref`/`avee` form is also covered. The pre-existing `_init_brex_list` cycle-guard/`None`-handling gap (next task item) is unchanged.
- [x] **Add a visited-set cycle guard** to the layered-BREX walk in `_init_brex_list`, and handle `get_brex_ref` returning `None` without a `TypeError`. Ref §3.7.
- [x] **Ship the built-in default BREX modules** (`s1kd-tools/tools/s1kd-brexcheck/brex/`, issues A/D/E/F/G/H plus `DMC-AE-A-…`) and add a "check against default BREX only" mode equivalent to `s1kd-brexcheck -B`, plus the fallback when the referenced BREX cannot be found. Implemented as a new `acd/default_brex.py` module plus the 7 bundled BREX XML files under `acd/brex/` (shipped as package data via `pyproject.toml`), ported from `s1kd-brexcheck`'s `brex/`, `default_brex_dmc`, `search_brex_fname_from_default_brex` and `load_brex` (`s1kd-brexcheck.c`): `default_brex_dmc(schema)` selects the schema-appropriate built-in BREX by the same substring rules as the C original (so e.g. `S1000D_4-0-1`/`S1000D_4-0-2` sub-issues still match `S1000D_4-0`), `default_brex_path(logical_dmc)` resolves one of the 7 logical DMCs to its bundled file path, and `find_default_brex_fallback(ref)` matches a BREX reference against the 7 built-ins by base DMC plus issue/inWork (accepting any issue when the reference does not specify one). `BrexChecker.use_default_brex(enabled=True)` in `acd/brex_checker.py` is the `-B`/`--default-brex` equivalent: `_init_brex_list` short-circuits to the single schema-selected built-in BREX, ignoring any `brexDmRef`/`brexref` and skipping layering entirely, matching the C original's `-B` behaviour. The fallback is wired into `_init_brex_list`'s existing `brexDmRef` walk (collapsed from two dead-duplicate branches into one while making this change, since branch two's guard was provably unreachable): when `find_document_by_reference` fails to locate a referenced BREX, `find_default_brex_fallback` is tried before giving up, at every layer of the chain (not just the top-level reference), exactly as `find_brex_fname_from_doc` is used uniformly in the C original. Per explicit product decision, the fallback is **not** silent like the C tool (which only logs when the fallback also fails) — every substitution is recorded in the result dict's new always-present `brexFallback` list (`Reference`, `UsedBuiltinBrex`, `BuiltinBrexPath` per entry), which `_append_summary` excludes from the error/warning count as purely informational. `override_brex_list`/`use_default_brex` reset `_brex_fallbacks` so stale markers never leak across runs or directory-mode files. Verified end-to-end against the real bundled files (not just fixtures): schema selection for all 7 built-ins resolves to an existing file, `use_default_brex()` checks a legacy S1000D 3.0 object against the real bundled `DMC-AE-A-...` BREX and ignores an unrelated `brexDmRef`, and a `brexDmRef` naming `DMC-AE-A-...` that can't be found on disk falls back to the real bundled copy (which self-references to terminate the layering walk, confirming the existing self-reference guard still applies to a fallback-resolved file) and is flagged in `brexFallback`; a reference to neither an on-disk nor a built-in BREX still raises `NoBrexDefined` unchanged. Covered by `tests/test_default_brex.py`.
- [x] **Add multiple BREX search paths and recursive search** (`-I`, `-r` equivalents) alongside the current single `set_brex_path`. Implemented in `acd/brex_checker.py`: `add_brex_search_path(path)` appends a directory to a new `_brex_search_paths` list (the `-I`/`--include` equivalent, repeatable), `clear_brex_search_paths()` resets it, and `set_brex_recursive_search(enabled)` toggles a new `_brex_recursive_search` flag (the `-r`/`--recursive` equivalent) that applies uniformly to the primary path and every added search path, mirroring the C original's `find_csdb_object(..., recursive_search)` calls for `search_dir` and each `spaths[i]` (`s1kd-brexcheck.c:538-547`). `_init_brex_list`'s resolution loop now tries `self._brex_dir_path[0]` first and, only if that misses, walks `_brex_search_paths` in the order added, stopping at the first match, before falling back to the built-in default BREX as before. `acd/s1000d.py`'s `find_document_by_reference` gained a `recursive: bool = True` parameter (default preserves the pre-existing always-recursive `os.walk` behaviour; `False` restricts the search to files directly inside the given directory, via `os.listdir`). Covered by `tests/test_brex_search_paths.py`: search-path fallback and ordering, rejecting a non-directory path, clearing search paths, recursive-on-by-default vs. recursive-off for both the primary path and an added search path, and direct unit tests of `find_document_by_reference`'s `recursive` parameter.
- [x] **Support S1000D <= 3.0 legacy BREX spellings (category C5)**: `objpath` / `objuse` / `objval`, `@objappl`, `@val1` / `@val2`, `contextrules/@context`. s1kd accepts all of these; we accept none. Implemented in `acd/brex_checker.py`: `_get_object_rule_nodes` now also collects `//structrules/objrule/objpath` alongside the modern `//structureObjectRuleGroup/structureObjectRule/objectPath` -- the real S1000D <= 3.0 default BREX nests rules exactly three levels below `contextrules` (`objpath` -> `objrule` -> `structrules` -> `contextrules`), same depth as the modern schema below `contextRules`, so the existing fixed-depth `getparent().getparent().getparent()` walk in `_show_rules` needed no change, just an `@rulesContext|@context` fallback lookup. `_show_rules` reads `objectUse|objuse`, `@allowedObjectFlag|@objappl`, and `objectValue|objval` with `@valueForm|@valtype` plus `@valueAllowed|@val1[~@val2]` (port of `get_value_allowed`, `s1kd-brexcheck.c:194-216`, which concatenates `val1`/`val2` into the same `first~last` string our `is_in_set` already parses). `_check_rules`'s dispatch loop now also routes a rule whose flag is entirely absent into the flag-2 value-check path whenever it carries `objval` children -- the shape of the great majority of rules in the real S1000D 3.0 default BREX (e.g. `//@accpnltype`, no `@objappl` at all) -- matching s1kd's `is_invalid` (`s1kd-brexcheck.c:307-331`), which falls through to `check_objects_values` whenever `allowedObjectFlag` is `NULL`. Verified end-to-end against the real bundled `acd/brex/DMC-AE-A-04-10-0301-00A-022A-D_003-00.XML` (the actual S1000D <= 3.0 default BREX shipped with s1kd-brexcheck), not just hand-built fixtures. Covered by `tests/test_legacy_brex_spellings.py`: flag 0/1/2 violations under legacy spelling, a flag-less value-only rule, `range` and `pattern` valtypes, a rule with no `objval` children never being a violation, `contextrules/@context` schema-qualification correctly skipping a non-matching object, a no-false-positives case with valid values, and the real bundled default BREX checked end-to-end against both a violating and a conforming object.
- [x] **Select rules with the descendant axis and filter at selection time**, matching `//contextRules[not(@rulesContext) or @rulesContext=$schema]//structureObjectRule`, and stop walking parents by fixed depth. Also removes the lxml `FutureWarning`. Ref §3.6. Implemented in `acd/brex_checker.py`: `_get_object_rule_nodes` now takes an optional `schema` and selects with `Element.xpath()` (real XPath via libxml2) instead of `findall()` (the restricted ElementPath dialect, whose leading `//` on a parsed tree is what raised the "This search incorrectly ignores the root element" `FutureWarning`), using the descendant axis `//contextRules[not(@rulesContext) or @rulesContext=$schema]//structureObjectRule/objectPath` with `$schema` bound as an XPath variable so only unqualified or schema-matching rules are selected at all, regardless of how deeply `structureObjectRule` is nested under `contextRules`. The S1000D <= 3.0 spelling gets the equivalent treatment: `//contextrules[not(@context) or @context=$schema]//objrule/objpath`. `_show_rules` threads `schema` through from its caller (`_check_rules`, which already had it from `get_schema_from_xml`) and, since the enclosing `contextRules`/`contextrules` element is no longer guaranteed to sit exactly three `getparent()` calls up, finds it with `next(x.iterancestors('contextRules', 'contextrules'), None)` instead — correct at any nesting depth, and a `None` result (no such ancestor) degrades to an empty `contextRules` rather than raising. Verified: full suite passes with `-W error::FutureWarning`, confirming the warning is gone; the existing `_check_object_flag_0/1/2` schema-equality checks stay in place as a harmless no-op now that filtering already happened at selection.
- [x] **Register every namespace in scope at the `objectPath` node** rather than a hard-coded `rdf` + `xsi` dictionary. Ref §3.12. Implemented in `acd/brex_checker.py`: `_show_rules` now captures `x.nsmap` (lxml's full in-scope namespace map, walking every ancestor up to the BREX root, not just a fixed depth) for each `objectPath`/`objpath` node, remaps the lxml default-namespace key `None` to `''` as elementpath expects, and stores it per-rule as `value['namespaces']` layered over `NS_DICT` as a base (so the well-known `rdf`/`xsi` prefixes stay resolvable even when a BREX declares neither). `_check_object_flag_0/1/2` now construct their `elementpath.Selector` with `value.get('namespaces', NS_DICT)` instead of the module-level constant, and the Saxon path in `_check_object_flag_0` declares each rule's namespaces (skipping the empty/default prefix, which `PyXdmNode.declare_namespace` does not accept) instead of iterating `NS_DICT`. `NS_DICT` itself is unchanged and still exported from `acd/__init__.py` for backward compatibility. Covered by `tests/test_objectpath_namespaces.py`: a prefix declared only on the BREX root resolves at a deeply nested `objectPath`, a prefix declared only on one `structureObjectRule` resolves for that rule and does not leak into a sibling rule's namespace map, `rdf`/`xsi` still resolve with no local declaration, and a direct check of `_show_rules`'s per-rule `namespaces` dict.
- [x] **Add `remove-deleted` support** (`-^`): drop elements with `@changeType="delete"` before checking. Implemented in `acd/brex_checker.py`: `_remove_deleted_elements` ports `rem_delete_nodes`/`rem_delete_elems` (`s1kd_tools.c:1054-1088`) -- recursively unlinks any element whose `@changeType` (or legacy `@change`) attribute is `"delete"`, without descending into a removed element's children (they go with it). `_check_rules` gained a `remove_deleted: bool = False` parameter; when set, it is applied to the parsed document right after parsing and before the schema-based SNS/notation/content-rule checks, matching the C original's ordering (`rem_delete_elems` runs once, immediately after parse, ahead of every subsequent check). Since nodes are removed in place on the already-parsed tree rather than by re-parsing, surviving elements keep their original `.sourceline`, so line numbers in violation records are unaffected. The Saxon path (`_check_object_flag_0`) previously always re-read `self._xml_path` from disk, which would have bypassed the in-memory removal; it now accepts an `xml_text` parameter and, when `remove_deleted` is active, `_check_rules` serialises the modified tree once and passes it through so Saxon evaluates the same reduced document as the `elementpath` path. `validate()` gained the matching `remove_deleted` parameter and threads it to `_check_rules` in both single-file and directory mode. Covered by `tests/test_remove_deleted.py`: a deleted element still flagged by default (opt-in only), the same element dropped from a flag-0 "must not be present" check once enabled, a non-deleted sibling still correctly flagged with the option on, the legacy `@change="delete"` spelling, a deleted parent taking its children with it, a flag-1 "must be present" rule that starts passing (element present) and starts failing (element removed) as the option is toggled, and the option threaded end-to-end through `validate()`.
- [x] **Add parser options**: XInclude processing, entity resolution, XML catalogs, and the `ignore-empty` behaviour for empty / non-XML inputs. Implemented in `acd/brex_checker.py`: `_build_xml_parser`/`_finish_parse`/`_parse_xml_file`/`_parse_xml_text` build a single lxml `XMLParser` from the checker's parser-option state and apply XInclude processing afterwards, and now back every internal parse of the checked object and every BREX file (`_check_rules`, `_get_object_rule_nodes`, `_get_sns_rules_group`, `_get_notation_rules_group`) instead of ad hoc `etree.parse(...)` calls -- mirroring `read_xml_doc` (`s1kd_tools.c:532-543`), which applies `DEFAULT_PARSE_OPTS` uniformly to every CSDB object it reads. `set_xinclude(enabled)` is the `--xinclude` equivalent (`ElementTree.xinclude()`, mirroring `xmlXIncludeProcessFlags`). `set_resolve_entities(enabled)` controls lxml's `resolve_entities` parser flag (`--noent`); `set_entity_resolution(load_external_dtd, allow_network)` is the `--dtdload`/`--net` pair, off by default so parsing a checked object cannot trigger unexpected file/network access on its own -- verified that an external-DTD-declared entity raises `XMLSyntaxError` by default and resolves once `load_external_dtd=True` is set (entities declared directly in the internal DTD subset, even via a `SYSTEM` identifier, were already resolved regardless, since that is a document-supplied declaration rather than a fetched external subset). `set_xml_catalog(path)` is the `--xml-catalog <file>` equivalent; since lxml has no catalog-loading binding, it appends to the `XML_CATALOG_FILES` environment variable, which libxml2 reads on its first catalog lookup (documented caveat: only takes effect if set before that first lookup in the process). `set_ignore_empty(enabled)` is `-e`/`--ignore-empty`: `_is_valid_xml_file` attempts a real parse and `validate()` uses it to skip an empty/non-XML object before `_init_brex_list()`/`_check_rules()` ever run -- left out of the results mapping entirely in directory mode (matching `s1kd-brexcheck`'s `continue`), reported as `{"Skipped": True, "Summary": "..."}` for a single object. Covered by `tests/test_parser_options.py`: XInclude on/off, external-DTD entity resolution on/off, the network-access-requires-dtd-loading guard, XML catalog registration/dedup/missing-file rejection, and ignore-empty for both single-file and directory mode with and without the flag set.
- [x] **Emit the node's canonical XPath and a copy of the violating subtree** in each violation record (categories D2, D3), shallow by default with an opt-in deep copy. Implemented in `acd/brex_checker.py`: `_select_with_nodes` re-implements the value-extraction step `elementpath`'s own `Selector.select` performs (`XPathToken.get_results`), but from the lower-level, un-formatted `root_token.select()`, so it can return the raw XPath node backing each node-set item alongside the same formatted result `.select()` would give -- the formatted API alone discards an attribute/text result's parent element entirely, which is what `dump_nodes_xml` (`s1kd-brexcheck.c:334-360`) needs to report a node's canonical XPath and copy it. `_node_xpath_and_copy` reads the node's `extended_path` (elementpath's own canonical-path computation, e.g. `/dml[1]/forbiddenElement[2]` or `/dml[1]/someElement[1]/@constrainedAttr`) and, for an attribute/text result, copies its owning element instead of the bare value (`if (node->type == XML_ATTRIBUTE_NODE) node = node->parent;` in the C original) -- shallow by default (tag and attributes only, matching `xmlCopyNode(node, 2)`), or the full recursive subtree when a new `deep_copy_nodes` parameter (the `-8`/`--deep-copy-nodes` equivalent) is enabled, threaded through `validate()` -> `_check_rules()` -> `_check_object_flag_0/1/2` / `_check_object_values`. Every content-rule violation record (flags 0, 1, 2) now carries `NodeXpath` and `Object` (an XML string, `None` when no node backs the violation, e.g. a flag-1 "required but missing" violation or a boolean-valued flag-0 rule). The Saxon path (`_check_object_flag_0`'s `self._saxon` branch) is left emitting `None` for both fields, consistent with its existing acknowledged incompleteness (§3.12, "Make the Saxon path complete or remove it"). Covered by `tests/test_node_xpath_and_copy.py`: canonical XPath and shallow copy for a repeated element violation (positional indices `[1]`/`[2]`), deep copy including descendants and tail text, an attribute violation resolving to its owning element, `None`/`None` for a missing-element flag-1 violation and a boolean flag-0 rule, and `deep_copy_nodes` threaded end-to-end through `validate()`.
- [x] **Report the violating node's real line number** from the parsed tree instead of scanning raw text for the attribute name. Ref §3.12, category D1. Implemented in `acd/brex_checker.py`: extracted the element-resolution half of `_node_xpath_and_copy` (walking a raw `_select_with_nodes` XPath node to its owning lxml element -- itself for an element result, the parent for an attribute/text result) into a shared `_resolve_owning_element(node)` helper, and added `_node_line_number(node)` on top of it that reads the owning element's `.sourceline` (lxml's binding to libxml2's `xmlGetLineNo`, matching what `s1kd-brexcheck` reports). `_check_object_flag_0`'s element/attribute branch and `_check_object_values` (shared by flags `1`/`2`) now call `_node_line_number(nodes[idx])` instead of the old `element.sourceline` attempt with a text-scan-for-the-attribute-name fallback on `AttributeError` -- that fallback matched the *first* line in the file containing the attribute name, so it silently misreported the line for every violation past the first repeated attribute/element in a document, and only handled attributes, never element or text-node results without a direct `.sourceline`. Fixed to `"x"` only when a violation has no backing node at all (a boolean-valued flag-0 rule, or a flag-1 "required but missing" result). The Saxon path (`_check_object_flag_0`'s `self._saxon` branch) still text-scans via `regex_builder`, matching its existing acknowledged incompleteness ("Make the Saxon path complete or remove it", §3.12/§4.4). Covered by `tests/test_violation_line_numbers.py`: multiple `forbiddenElement` violations resolve to their own distinct real lines rather than all collapsing to the first occurrence, repeated-attribute violations resolve to each owning element's real line, an element-text value violation resolves correctly, and `Line` is a real `int` rather than a string.
- [x] **Add an XML report output** compatible with the `s1kd-brexcheck -x` shape (`brexCheck/document/brex/error/{brDecisionRef,objectPath,objectUse,object}`) so existing tooling can consume our output (category D6). Implemented in `acd/brex_checker.py`: `to_xml_report(result)` converts a `validate()` result (single-object or directory-mode) into a `<brexCheck>` document tree matching the C tool's shape (`document[@path]/{sns,notations,brex[@path]/{error,xpathError}}`), ported piecewise from the relevant `s1kd-brexcheck.c` fragments -- `_append_error_node` from the `<error>` construction in `check_brex_rules` (`s1kd-brexcheck.c:900-938`, including its `brSeverityLevel`/`fail` attribute logic), `_append_sns_notation_nodes` from `check_brex_sns_rules` (`s1kd-brexcheck.c:1077-1134`) and `check_brex_notation_rules` (`s1kd-brexcheck.c:1213-1224`). One difference from the C original: our violation records are already one-per-matched-node (see `_check_object_flag_0`/`_check_object_values`), so each becomes its own `<error>` with at most one `<object>` child, rather than one `<error>` per rule holding several `<object>` children. `objectValue` echoing (allowed values) is not emitted, since our violation records store parsed value lists rather than the original `<objectValue>` nodes -- out of scope for this task, which only asked for the `brDecisionRef`/`objectPath`/`objectUse`/`object` shape. Covered by `tests/test_xml_report.py`: root/document/brex shape, flag-0 error with `brDecisionRef`/`objectPath`/`objectUse`/`object` (including line/xpath attributes and the embedded node copy), flag-1 missing-element error with no `object` child, `brSeverityLevel`/`fail` attribute combinations (`fail="no"`, no-severity `fail="yes"`, failing-severity omits `fail`), `xpathError` reporting, SNS `noErrors`/`error` nodes, a notation `error` node, directory-mode conversion, and the `ignore-empty` skipped-result case.
- [ ] **Add a run summary** with documents passed/failed and violations by severity (category D7).

### 4.3 P2 — Rule categories neither tool checks

- [ ] **Surface `nonContextRules/nonContextRule` (category A4)** in the report as informational entries — text plus `brDecisionRef` — so authors see the human-readable business rules that apply. They are not machine-checkable, but silently dropping them hides a whole rule class.
- [ ] **Report `objectValue/@valueTailoring` (category C4)** with each value violation, distinguishing `lexical` (value list may be extended) from `restrictable` (value list may only be narrowed), so downstream tailoring review knows which allowed-value sets a project may legally change.
- [ ] **Add a BREX self-consistency lint pass** that runs over a BREX before it is used:
  - [ ] every `objectPath` compiles;
  - [ ] every `structureObjectRule` has an `objectUse`;
  - [ ] `allowedObjectFlag="2"` rules actually carry `objectValue` children (18 in these three BREX do not, and are silently no-ops);
  - [ ] `valueForm="pattern"` values are valid XSD regexes;
  - [ ] `valueForm="range"` values parse as a range or set;
  - [ ] `@brSeverityLevel` values exist in the severity-level file;
  - [ ] duplicate `structureObjectRule/@id` and duplicate `brDecisionIdentNumber`.
- [ ] **Detect rules made unreachable by `rulesContext`** — e.g. a `rulesContext` naming a schema no object in the CSDB uses, or a typo'd schema URI. Report as a BREX warning.
- [ ] **Report per-rule hit statistics** (rules evaluated / matched / violated) so a project can see which BREX rules never fire against its data set.
- [ ] **Add cross-BREX conflict detection for layered BREX** — the same `objectPath` given contradictory `allowedObjectFlag`s, or `restrictable` value sets widened by a lower layer.
- [ ] **Validate SNS rule tables themselves** (duplicate `snsCode` at a level, missing `snsTitle`, codes outside the pattern declared elsewhere in the BREX).

### 4.4 P3 — Improvements to checks we already perform

- [ ] **Read `objectUse` as full text content** (`''.join(node.itertext())`) instead of `.text`, and tolerate a missing `objectUse` instead of raising `IndexError`. Ref §3.12, category D4.
- [ ] **Parse each BREX once per run**, not once per data module, and cache the extracted rule list. Ref §3.12.
- [ ] **Compile each `objectPath` once** into a reusable selector and evaluate it against each document, instead of recompiling per (rule, document). Ref §3.12.
- [ ] **Escape attribute values in `regex_builder`** (`regex.escape`), or drop the raw-text search entirely once line numbers come from the parsed tree. Ref §3.12.
- [ ] **Make the Saxon path complete or remove it** — it is only wired into `_check_object_flag_0`, so `saxon=True` currently gives an inconsistent mix of engines. Ref §3.12.
- [ ] **Select the XPath version the way s1kd does**: XPath 1.0 for BREX declaring S1000D <= 3.0, XPath 2.0 for 4.0+, with an explicit override. Ours is always XPath 2.0 — correct for these BREX, wrong for a 3.0 BREX that relies on XPath 1.0 `=`-on-node-set semantics.
- [ ] **Replace the `".xml" in name` / `"-022a-" not in name` directory filter** with a real extension test, and check BREX data modules like any other object. Ref §3.11.
- [ ] **Return structured violation objects** (dataclass: brex, rule id, `brDecisionIdentNumber`, flag, severity, `objectPath`, `objectUse`, allowed values, node xpath, line, node snippet) instead of nested dicts keyed by integers, and derive the JSON and XML reports from them.
- [ ] **Deduplicate identical violations** raised by the same rule appearing in multiple layered BREX.
- [ ] **Add a progress callback** rather than the hard `tqdm` dependency inside the library.
- [ ] **Raise a clear error when the object has no `xsi:noNamespaceSchemaLocation`** (DTD-based objects) instead of comparing `rulesContext` against `None`.

### 4.5 P4 — Test suite

- [ ] **Add the 10-violation fixture** from §1.1 as a regression test asserting exactly 10 violations with the correct flags and messages.
- [ ] **Add fixtures for each `allowedObjectFlag`** covering node-set, boolean, string and numeric XPath results.
- [ ] **Add a `valueForm` fixture set** covering `single`, `pattern` (including XSD subtraction and whole-value anchoring) and `range` / set forms.
- [ ] **Add an SNS fixture** exercising normal, strict and unstrict modes against `DMC-ATABREX-F-00-00-00-01A-022A-D_004-00_EN-US.XML`.
- [ ] **Add a notation-rule fixture** with a DM whose internal DTD subset declares an excluded notation.
- [ ] **Add a layered-BREX fixture** (`DM → ATABREX 01A → ATABREX 00A → S1000D default`) asserting all layers are collected exactly once, plus a cyclic-reference fixture.
- [ ] **Add a `rulesContext` fixture** with `xsi:noNamespaceSchemaLocation` in first, middle and last attribute position, asserting the same rule set is selected each time.
- [ ] **Add a malformed-`objectPath` fixture** asserting the run completes and reports an `xpathError` for that rule only.
- [ ] **Add a differential harness** that runs both `brex_checker.py` and `s1kd-brexcheck -clnST` over the whole `CMP 21-77-05` folder and diffs the violation sets, so parity is measurable and intentional divergences are recorded.
- [ ] **Add a performance baseline** over the 63-object folder x 919 rules to protect the caching work in §4.4.

---

## 5. Suggested order

1. §4.1 in full — without it every downstream number is wrong (3 of 10 violations reported today).
2. §4.5 fixtures for the P0 items, so the fixes stay fixed.
3. SNS rules, `brDecisionRef` and severity levels from §4.2 — the three categories a customer will notice missing first.
4. The structured-violation refactor in §4.4, then the XML report and the differential harness.
5. §4.3 as differentiators once parity with `s1kd-brexcheck` is established and measured.
