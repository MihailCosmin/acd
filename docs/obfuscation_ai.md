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

   `instructions=` adds text to the prompt (for example "the customer's CAGE code is D9893"), and
   `keep=` lists strings that are never obfuscated.

3. **Check.** Nothing the model says is taken on trust:
   - A string is kept only if it occurs in its chunk. Whitespace and case may differ, and the
     text's own spelling is used. Made-up strings are rejected.
   - Dates (also spaced out, "20. 04. 2022"), quantities, page numbers and sizes ("1200x800")
     are rejected. So are field labels ("DCN-Nr", "lfd. Nr"), strings with fewer than 3 letters
     or digits, and ordinary words ("Page", "Approved").
   - A value the model copied together with its label ("DCN Number : LLI_DCN-114242") keeps
     only the value. A string cut off at an unbalanced bracket keeps the part before it.
   - A string must also look like its category: a person has no digits, an identifier has at
     least one digit, an e-mail address has an `@`, a phone number has at least 6 digits and no
     words, and an address has a word in it.
   - A string written in other case elsewhere ("HEIKE WOCHNER") is obfuscated in that spelling
     too, because `obfuscate_document` matches literal strings case-sensitively.
   - The PDF producer and creator fields name the software that made the file, so they are not
     sent to the model.
   - Two or more identifiers of the same shape become a regular expression. For example,
     `1641B0100-01` and `1641B0230-02` become `\d{4}B\d{4}-\d{2}`, which also catches the numbers
     the model missed. A shape needs at least 4 digits.
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
- `--report <folder>` writes `report.html`, which shows every page of every document side by
  side, original and obfuscated, with what changed boxed in red. It flags pages where nothing
  changed. The report shows the originals, so keep it where the originals may be.
  `python -m acd.obfuscation_report <originals> <obfuscated> <report folder> --summary summary.json`
  writes it on its own.

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
