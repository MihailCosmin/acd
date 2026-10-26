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
   4. the defaults: the Ollama server `http://10.3.1.174:11434`, model `qwen3.5:9b`, no key.

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
   tokens. Each chunk overlaps the previous one by a line. A chunk that repeats (the same form in
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
python -m acd.obfuscation_ai suggest a.pdf b.xml -o targets.json --model qwen3.5:4b
python -m acd.obfuscation_ai apply targets.json a.pdf b.xml --out-dir obfuscated
```

`apply` gives a string the same placeholder in every file.

The document text is sent to the provider, so use a local model for confidential documents.

## Small models tested

Tested on 26.10.2026 against the Ollama server `10.3.1.174`. The server was shared, and it gave
qwen2.5-coder:3b about 3.5 tokens/s. Two sources were used:
- a synthetic inspection sheet with 11 sensitive strings: 2 people, a company, an address in two
  parts, an e-mail address, a phone number and 4 part numbers, 2 of which share a shape;
- two real CMM 67-35-81 PDFs, a DCN and a packaging data sheet in German. A reviewer would
  obfuscate 14 strings in them: 5 people and 9 identifiers, among them the CAGE code D9893.

| model | synthetic: found | real: found | real: false positives | time for the 2 PDFs |
|---|---|---|---|---|
| qwen3.5:9b | 11/11 | 14/14 | 0 | 79 s |
| qwen3.5:4b | 11/11 | 13/14 (missed 079388050) | 3 (LLI-FN, MHA-LONGITUDINAL, and "Tobias" split from "Nägele, Tobias") | 39 s |
| qwen2.5-coder:3b | 9/11 (missed the address) | 8/14 | 14 (German nouns: Palette, Holzrahmen, Verpackungstyp …) | 61 s |

Without the checks in step 3, the same runs would have let through:
- a made-up `LILI_DCN-114242`;
- dimensions filed as addresses;
- dates;
- "Crystal Reports" and "PDFsharp" as companies.

Both qwen3.5 models need thinking and JSON mode off, which is the default for an Ollama URL.
qwen2.5-coder:3b takes ordinary German words for names, so do not use it.

The default is qwen3.5:9b; qwen3.5:4b is about twice as fast and nearly as good. Whatever the
model, review the JSON before you use it.
