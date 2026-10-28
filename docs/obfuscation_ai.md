# AI-suggested obfuscation targets

`acd.obfuscation.obfuscate_document` blacks out or replaces the targets it is given: a dict of
strings and compiled patterns, each mapped to `"blackout"` or `"placeholder"`. `acd.obfuscation_ai`
has a language model draft that dict. Any OpenAI-compatible chat completions API works: a local
Ollama server, OpenAI, LM Studio, vLLM or llama.cpp's server.

## Process

1. **Configure the provider.** You need a base URL and a model, plus an API key if the provider
   asks for one. Each setting is read from the first of these that has it:
   1. the arguments of `ai_provider(...)` or the CLI flags `--base-url --model --api-key`;
   2. the environment variables `ACD_AI_BASE_URL`, `ACD_AI_MODEL`, `ACD_AI_API_KEY`, `ACD_AI_TIMEOUT`;
   3. a JSON file, given as `config=` / `--config`, or through `ACD_AI_CONFIG`;
   4. the defaults: the Ollama server `http://10.3.1.174:11434`, model `gemma4:26b-iq4xs-text-135k`, no key.

   ```json
   {
     "base_url": "https://api.openai.com/v1",
     "model": "gpt-5-mini",
     "api_key_env": "OPENAI_API_KEY",
     "timeout": 600,
     "options": {"max_tokens": 2000}
   }
   ```

   Use `api_key_env` to name the environment variable that holds the key, so that the key stays
   out of the file. Everything under `options` is sent with every request. Answers are capped at
   1500 tokens by default (`max_tokens`), so a model that starts repeating itself cannot hold up a
   run.

   A base URL without a path is treated as an Ollama server. `/v1` is added to it, and two
   options are set by default (`AI_OLLAMA_OPTIONS`), which `options` can override:
   - `reasoning_effort: "none"` turns thinking off. Without it, qwen3.5:4b thought for more than
     5 minutes over a one-line question.
   - `response_format: null` turns JSON mode off. With JSON mode, qwen3.5:4b wrote 2 characters
     in 3 minutes; without it, it wrote the whole answer in 2 minutes. The answer is parsed
     either way: fences, `<think>` blocks and text around the JSON are dropped, and a model that
     answers without JSON is asked once more.

   For other servers the same switches are available as `--no-think` and `--no-json-mode`.

2. **Suggest.** `suggest_obfuscation_targets(files, provider)` reads every document:
   - PDF pages, annotations, form fields and metadata;
   - text files;
   - scans and image files, which are read with OCR when `ocr=True`.

   The text is cut into chunks of 2500 characters, which fits Ollama's default context of 4096
   tokens. A model with a larger context takes larger chunks (`chunk_chars=` / `--chunk-chars`):
   the default gemma4 model has 135k tokens and did 15 PDFs in 96 s with 30000 characters,
   141 s with 2500, and found the same strings. Each chunk overlaps the previous one by a line. A chunk that repeats (the same form in
   many files) is sent only once. The model lists the strings it finds in each of these categories
   (`AI_CATEGORIES`, or your own):

   | category | obfuscated as | becomes a pattern |
   |---|---|---|
   | person (names, initials with surname, signatures) | blackout | no |
   | email, phone, address | blackout | no |
   | company | placeholder | no |
   | identifier (part, drawing, document, serial, order numbers, CAGE codes) | placeholder | yes |
   | project (project, programme and aircraft or helicopter type numbers: EC 135, H145, a root number 1644) | placeholder | yes |

   The prompt also tells the model that values after labels such as Project, Projekt, Program,
   Platform, Root number or Aircraft are codes, even when they look like a model name or a short
   number. A project number of digits only ("1644") becomes a pattern that does not match inside
   longer numbers ("21644").

   Dates are not left to the model: built-in patterns find them (`AI_PATTERN_CATEGORIES`), and
   they are blacked out. The patterns match day, month and year with dots or slashes, also spaced
   out ("20. 04. 2022") or with a two-digit year ("24.11.99"); ISO dates; day-month-year with
   hyphens and a four-digit year; and dates with English or German month names. They do not match
   tolerances ("12.27"), part numbers ("1644-0102") or version numbers. `pattern_categories={}`
   turns this off.

   `instructions=` adds text to the prompt (for example "the customer's CAGE code is D9893"), and
   `keep=` lists strings that are never obfuscated.

3. **Check.** Nothing the model says is taken on trust:
   - A string is kept only if it occurs in its chunk. Whitespace and case may differ, and the
     text's own spelling is used. Made-up strings are rejected.
   - Dates the model lists, quantities, page numbers and sizes ("1200x800") are rejected (the
     date patterns take care of dates). So are field labels ("DCN-Nr", "lfd. Nr"), strings with fewer than 3 letters
     or digits, and ordinary words ("Page", "Approved"). Values of a drawing are rejected as
     "technical value": surface roughness ("Rz 1", "Ra 0.8"), diameters with or without a fit
     ("Ø7H8") and metric threads ("M10x1.5").
   - A value the model copied together with its label ("DCN Number : LLI_DCN-114242") keeps
     only the value. A string cut off at an unbalanced bracket keeps the part before it.
   - A string must also look like its category: a person has no digits, an identifier and a
     project have at least one digit, an e-mail address has an `@`, a phone number has at least 6 digits and no
     words, and an address has a word in it.
   - A string written in other case elsewhere ("HEIKE WOCHNER") is obfuscated in that spelling
     too, because `obfuscate_document` matches literal strings case-sensitively.
   - The PDF producer and creator fields name the software that made the file, so they are not
     sent to the model.
   - Two or more identifiers of the same shape become a regular expression. For example,
     `1641B0100-01` and `1641B0230-02` become `\d{4}B\d{4}-\d{2}`, which also catches the numbers
     the model missed. A shape needs at least 4 digits, and something besides digits: a shape
     of 8 digits alone would match any 8-digit number, such as two item numbers that OCR read as
     one (`0220 1850` on a drawing as `02201850`), or the grid labels of a sheet (`910111213`).
     Such strings stay targets on their own. The price: a number of that shape the model missed
     is no longer caught. On the test folder that was one, the "ID-Nr. (M-Item)" 13935344 of
     E2-2019070B122; add such numbers to the targets when reviewing them.
   - With `model_patterns=True`, the model proposes patterns as well. A proposed pattern is kept
     only if it compiles, matches its example, does not match the empty string, covers at most
     20% of the text, and matches nothing but identifiers (or strings already found).

4. **Review.** `save_suggestions(suggestions, "targets.json")` writes the suggestions as JSON:
   - every string, with its category, type, number of occurrences and source files;
   - every pattern, with what it matches;
   - everything that was rejected, and why;
   - `failures`: chunks the model gave no usable answer for (a timeout, or no JSON twice). Nobody
     has checked that text, so review it yourself. If the very first request cannot reach the
     provider, or the provider answers with an HTTP error, the run stops instead.

   Delete what may stay, change a type, and add what is missing. The model's suggestions are a
   first draft for a person to check, not a final list.

5. **Obfuscate.** `load_targets("targets.json")` returns the targets dict that `obfuscate_document`
   and `obfuscate_text` take. Check the output: OCR and the model can both miss strings.

```
python -m acd.obfuscation_ai suggest a.pdf b.xml -o targets.json --chunk-chars 30000
python -m acd.obfuscation_ai apply targets.json a.pdf b.xml --out-dir obfuscated
```

`apply` gives a string the same placeholder in every file. Its other options:
- `--logos <images or folders>` removes logos (see below) after the text is obfuscated. Then
  OCR reads the same pages that `suggest` read, so a shared `--ocr-cache` saves all of it. The
  logos are searched for in the original, because black boxes and placeholders in the
  obfuscated copy can hide a logo or pass for one. Their areas are then blacked out in the copy;
- `--summary summary.json` writes what was replaced in each file, and which logos were removed;
- `--ocr-cache <folder>` (also for `suggest`) keeps what OCR read, by the image's content. OCR
  takes about a minute per scanned A4 page on a laptop CPU, and longer for a drawing, so
  without the cache every rerun reads all the scans again. The cache files are pickles: use a
  folder nobody else writes to;
- `--signatures` blacks out signatures (see below);
- `--stamps` blacks out stamps (see below), those on scans and in images only with `--ocr`;
- `--report <folder>` writes `report.html`, which shows every page of every document side by
  side, original and obfuscated, with what changed boxed in red. It flags pages where nothing
  changed. The report shows the originals, so keep it where the originals may be.
  `python -m acd.obfuscation_report <originals> <obfuscated> <report folder> --summary summary.json`
  writes it on its own.

### A whole folder, with placeholders unique to it

`obfuscate_folder(folder, output_dir, targets)` (or `python -m acd.obfuscation_ai folder`)
obfuscates every PDF under a top folder and its subfolders into a copy of the folder. The
folder gets a random code of 5 characters, such as `E3B1Q` (from `0-9A-Z` without `I` and `O`,
so about 45 million codes). Every string gets one placeholder, the same in every file:
`PH-E3B1Q-000000001`, `PH-E3B1Q-000000002`, and so on. When text that mixes several projects'
documents is deobfuscated, a placeholder can only mean one thing.

The key goes next to the output folder (`<output folder>.key.json`), not into it, because the
output is what gets shared. It holds:
- the code and the placeholder format;
- every placeholder's original;
- what was done to each file: strings, logos, signature areas, or the error.

It is written after every file. A run that stops is resumed by starting it again: files already
done are skipped, and their placeholders are kept. `deobfuscate_text(text, key)` (or
`python -m acd.obfuscation_ai deobfuscate --key <key> [files]`) puts the originals back for
that folder's placeholders. Placeholders of other folders stay as they are, and so do blacked-out
strings.

```
python -m acd.obfuscation_ai suggest <folder>/*.pdf -o targets.json --ocr --ocr-cache cache
python -m acd.obfuscation_ai folder <folder> --out-dir <copy> --targets targets.json \
    --logos logos --ocr-cache cache --report report
python -m acd.obfuscation_ai deobfuscate --key <copy>.key.json answer.md
```

By default `folder` reads scans with OCR (turn it off with `--no-ocr`) and blacks out
signatures and stamps (turn them off with `--no-signatures` and `--no-stamps`). `--max-files N` stops after N files, and the
next run goes on from there. Use it to split a long run into steps.

A placeholder longer than the string it replaces first runs on into free space. If it still
does not fit, it shrinks, down to 3 pt; below that it would be unreadable, so it overflows
instead. In images the limit is 8 pixels. An 18-character folder placeholder fits in the
place of a 10-character part number at about two thirds of the size.

### Signatures and handwriting

`acd.signature_obfuscation` cannot read a signature, so it does not try to. It finds the place
a printed label names, and blacks that out:
- **Labels** are lines of at most 5 words containing "Signature(s)", "Unterschrift(en)",
  "Signed (by)" or "Visum". "Approval", "approved by" and "Genehmigung" count only with a ":",
  "&", "date", "Datum", "by" or "durch" in the line ("PURCHASER APPROVAL & DATE"). A bare
  "Approval" is a signature's type, not a label. "Signature type" and "Electronic signature" are
  values too.
- **Cells** are found from the table's ruling lines: runs of ink on a render of the page, plus
  the PDF's vector lines, because a signature image with a white background hides the line it was
  pasted over. Copies of scanned forms often show their cell borders light gray, so thin runs of
  faint ink (gray below 215 of 255) count too; faint areas 6 pixels thick or more are shading,
  not lines. Lines are drawn on 4 pixels past their ends, because on a scan a cell's border
  can stop just short of the line it meets, and the cell would run into its neighbour. Two
  label lines directly under one another in one cell are one label ("Unterschrift" over
  "SIGNATURE"). Of a label OCR read, only its words count, from the start of the line to the
  last label word, without the detector's margin: OCR's line can take in a signature written
  beside or just under it.
- **Which cells:** if the label fills its cell, it is a header. A column header names the cells
  below it of the same width. A row header (one with a cell of its width above it and none to its
  left: "Unterschrift" between "Datum" and "Name") names the cells to its right of the same
  height. If the label does not fill its cell, the cell is what it names: all of it except the
  label. A cell with no ink in it is left alone. Where the next cell is looked for, a line that
  is not the border (a stamp's frame lying on it) is stepped over.
- **What is removed:** everything in those areas, under a black box. That is images, vector pen
  strokes and text. An image or pen stroke that lies at least 30% in a signature area is removed
  whole, because a signature runs over its cell's lines. Pages without text are read with OCR to
  find the labels.

On a whole page, the OCR detector can take a printed label, the signature written across it and
the text beside it for one box, which then reads as nothing. `line_ocr` detects such a box again
on its own, where it splits into lines that read: "SUPPLIERS SIGNATURES & DATE" on the scanned
Eurocopter form is found that way.

Tested on 27.10.2026:
- **The LSA record:** the 4 signature images under its "SIGNATURE" header are found.
- **DCNs:** the "Electronic Signature" column and the dated "Process Approval Date" cell are
  found. The "Approval" values are left alone.
- **The scanned Eurocopter form:** the "PURCHASER APPROVAL & DATE" cells are found.

Tested on 28.10.2026, after the fixes above:
- **The LSA record:** the "Accepted by Customer" stamp in the fifth row is found too (its frame
  lay on the cell's border, which stopped the search down the column).
- **The scanned Eurocopter form M1644-01:** the "SUPPLIERS SIGNATURES & DATE" cell on page 1
  is found, without running into the "DECLARATION OF DESIGN" cell beside it; the purchaser's
  handwriting in the line of its label is covered. On pages 12 and 13 the signatures above the
  "Signature" lines are found (14 to 18 areas per page instead of 2 and 0: their faint borders
  did not close), the ones under "APPROVED BY CUSTOMER" too. The "Project" signature on page 13
  is not: it is written across its "Signature" label, which OCR does not read.
- **The approval table of MO-1644-00 (page 2):** the 3 signatures right of "Unterschrift /
  SIGNATURE" are found; before, the "Name" cell below was blacked out instead.

Still not found: handwriting next to labels that name no signature ("Design: ....", "Commerce:"
in the remarks of M1644-01), and handwritten notes in unlabelled cells. A signature stroke
running straight for 25 points or more counts as a ruling line and can cut its cell in two; the
part beyond it stays. I also tried taking all ink that OCR could not read in any table cell,
which would find unlabelled handwriting. It blacked out a third of a drawing's views, so it is
not used.

### Stamps

`acd.stamp_obfuscation` blacks out stamps, the same way as signatures:
- **Stamp annotations** (PDF /Stamp): all of them, deleted and their area blacked out.
- **Stamp images:** an image of at most 3% of the page in which OCR reads a stamp word
  (accepted, approved, released, rejected, verified, inspected, received, certified,
  freigegeben, geprüft, genehmigt, abgelehnt, eingegangen, kontrolliert) in a line of at most 6
  words. The whole image goes.
- **Framed stamps on scans** and in large images: a stamp word in a frame of its own of at most
  3% of the page. A table cell is no stamp's frame: "genehmigt / APPROVED" heads the columns of a
  drawing's approval table. So a frame whose border runs on beyond its corners (past its own
  line width of 4 points, for 6 points more) is left alone.

A stamp without a frame on a scan is not found, nor one whose words OCR cannot read. Tested on
28.10.2026: the "Accepted by Customer" stamp image in the LSA record is found; the approval
tables of MO-1644-00 and M1644-01 (with "APPROVED", "ACCEPTED", "genehmigt" in their headers and
"APPROVED BY CUSTOMER" in a form) give no stamp. The corpus of 944 PDFs holds 5 stamp
annotations (an AIRBUS logo and "*) This document has been signed according ..." notes).

### OCR of item numbers

RapidOCR reads the item numbers of neighbouring balloons on a drawing as one run of digits
(`(0220)(1850)` as `(02201850)`). `line_ocr` puts the space back where one gap between digits is
at least 1.3 times as wide as the next smaller ones and splitting there leaves numbers of 3
digits or more. Only runs standing alone are split (between spaces, brackets or the line's
ends), not digits in a code (`E2-2019390`). A misreading of a number known to be sensitive is
not looked for inside a longer number either. The OCR cache keys on the OCR version
(`line_ocr.OCR_VERSION`), so pages read before these changes are read again.

### Logos

`acd.logo_obfuscation` finds logos by how similar they look to reference images, and removes
them. A logo can be an image in a header, vector outlines in a drawing's title block, or
pixels on a scan; all three are found. Give one reference image per variant of the logo, for
example outlined, solid, and with a slogan, each cropped close to the logo.

1. **Search.** Every page is rendered at up to 100 dpi, with at most 2400 pixels on its longer
   side. Each reference is matched at every size from 8 to 72 pt high, in steps of 10%, at its own
   width and at 72% of it. Drawings print the same wordmark with taller letters. The matching is
   normalized cross-correlation on the page's ink (dark or coloured against light), blurred by a
   12th of the logo's height, so line weight, colour and scan noise matter little. Blank paper
   is never a match.
2. **Check.** The page is rendered once more, sharp, with at most 9000 pixels on its longer
   side. Each candidate is cut from it with the logo 60 pixels high, and compared at sizes of
   80–125% and widths of 88–113% of its estimate, with a light blur. It is a logo only if all of
   these hold:
   - the similarity is at least 0.6 (`LOGO_SCORE`);
   - its letters are at least 8 pt high (for a logo with a slogan under it, like 3.png, the whole
     logo is then at least 11 pt high);
   - it inks about as much of its box as the reference does;
   - no single connected shape covers more than 40% of its box (it is letters, not a filled bar);
   - it has few mid-tones, and it sits on paper.

   A line of text, a black box or a photo fails one of these.

   It must also not read as other words. OCR reads each reference once ("LIEBHERR",
   "AIRBUS HELICOPTERS"). A match is dropped if text that is like none of those words covers at
   least half of it (the PDF's words, or OCR lines with `ocr=True`).

   A logo made of several lines of text, such as a name over a slogan, needs more. Hatching and
   two lines of bold type look like it. Its match is dropped if any other word is in it, or if
   OCR read the area and found nothing like the logo's words. "Read the area" means a page
   without text, or an image on a page with text. In the test folder this removed 34 false
   matches of the AIRBUS HELICOPTERS reference on section hatching and bold headings.
3. **Remove.** In a PDF the logo is redacted: image pixels under it are blanked, vector paths
   inside it and text over it are deleted, and a black box takes its place. In an image file,
   the box is painted over it.

Results on 16 test PDFs of CMM 67-35-81 (27.10.2026), 87 pages, with 4 references (outlined,
solid, blue with "Aerospace.", small solid):
- **Removed:** 54 logos. 33 header logos of the inspection sheets, the LSA record and the MO
  procedure (one per page), 11 blue DCN logos (one per page), the outlined wordmark in 8 drawing
  title blocks, and 2 wordmarks on nameplate drawings. True logos scored 0.61 to 0.99.
- **Not found:** the outlined wordmark in 2 more title blocks (S4-9504760W02, Z3-9504760). The
  text pass had already removed it there: OCR read "LIEBHERR", and it became a placeholder.
- **Wrongly matched:** 2 small areas of a scanned Eurocopter form, at 0.61 and 0.67 (form lines,
  and the bold word "EQUIPMENT"). They were blacked out.
- **Rejected before the per-side paper and shape checks:** text lines, the edges of photos, and
  black redaction boxes had all matched at up to 0.86.

Rotated logos are only found with `rotations=(0, 90, 270)`, which takes 3 times as long. A
search takes 10–20 s for an A4 page and 1–2 minutes for an A0 sheet.

The document text is sent to the provider, so use a local model for confidential documents.

## Models tested

Tested against the Ollama server `10.3.1.174`. Two sources were used:
- a synthetic inspection sheet with 11 sensitive strings: 2 people, a company, an address in two
  parts, an e-mail address, a phone number and 4 part numbers, 2 of which share a shape;
- two real CMM 67-35-81 PDFs, a DCN and a packaging data sheet in German. A reviewer would
  obfuscate 14 strings in them: 5 people and 9 identifiers, among them the CAGE code D9893.

| model | synthetic: found | real: found | real: false positives | time for the 2 PDFs |
|---|---|---|---|---|
| gemma4:26b-iq4xs-text-135k (default) | 11/11 | 14/14 | 0 | 9 s |
| gemma4:26b-iq4xs-28k | 11/11 | 14/14 | 0 | 9 s |
| qwen3.5:9b | 11/11 | 14/14 | 0 | 79 s |
| qwen3.5:4b | 11/11 | 13/14 (missed 079388050) | 3 (LLI-FN, MHA-LONGITUDINAL, and "Tobias" split from "Nägele, Tobias") | 39 s |
| qwen2.5-coder:3b | 9/11 (missed the address) | 8/14 | 14 (German nouns: Palette, Holzrahmen, Verpackungstyp …) | 61 s |

The gemma4 runs (27.10.2026) had the GPU to themselves; the qwen runs (26.10.2026) shared it,
with about 3.5 tokens/s for qwen2.5-coder:3b. Both gemma4 models are the same 26B-A4B mixture of
experts in IQ4_XS; the 28k one also takes images.

On 15 more PDFs of the same CMM, picked at random, the default model listed 140 strings and 15
patterns in 141 s, with no failed chunk: 31 people, 101 identifiers, 7 companies and the company
address. A reviewer would delete a few of them: `BAAN` (the ERP system), a marker pen
("Faber Castel Multimark 1523 permanent"), the aircraft type `EC 135`, and public standard
numbers (`AMS 2641`, `M39029/58-360`, `MS3198-8P`).

Without the checks in step 3, the runs would have let through:
- a made-up `LILI_DCN-114242`;
- dimensions filed as addresses;
- dates;
- "Crystal Reports" and "PDFsharp" as companies.

The qwen3.5 and gemma4 models need thinking off, and qwen3.5 also needs JSON mode off; both are
the default for an Ollama URL. qwen2.5-coder:3b takes ordinary German words for names, so do
not use it. Whatever the model, review the JSON before you use it.

### Vision instead of OCR

`suggest_obfuscation_targets` reads scans and images with OCR, not with a vision model. Asking
gemma4:26b-iq4xs-28k directly with the image does not beat that yet (27.10.2026):
- The server's GPU runs out of memory on images larger than 1024 px, so pages are sent at 1024 px.
- The two real PDFs, sent as page images, gave 12 of 14 strings, against 14 of 14 from their text.
  The model misread `LLI_DCN-1144242` and "Nägeler, Tobias".
- On ICN illustrations (2025x2476 px), it found nothing. `D9893` in the margin is too small at
  1024 px.
- `obfuscate_document` needs OCR boxes to black out a string in an image anyway.
