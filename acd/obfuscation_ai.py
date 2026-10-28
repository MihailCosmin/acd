r"""Let an AI model suggest what `obfuscation` should obfuscate.

`obfuscate_document` needs the targets -- the strings and patterns to black
out or replace. This module has a language model read the documents and
suggest them, through any OpenAI-compatible chat completions API: OpenAI,
a local Ollama, LM Studio, vLLM, llama.cpp's server and the like.

The process:

1. **Configure the provider** -- base URL, model, API key if it needs one
   (`ai_provider`): as arguments, as the environment variables
   `ACD_AI_BASE_URL`, `ACD_AI_MODEL`, `ACD_AI_API_KEY`, or as a JSON file
   (`ACD_AI_CONFIG` or the `config` argument):

       {"base_url": "https://api.openai.com/v1", "model": "gpt-5-mini",
        "api_key_env": "OPENAI_API_KEY", "options": {"max_tokens": 2000}}

   `options` go into every request as they are. A base URL without a path
   is an Ollama server: `/v1` is added, and thinking and JSON mode are
   off unless the options say else (`AI_OLLAMA_OPTIONS`).

2. **Suggest** (`suggest_obfuscation_targets`): the text of every document
   is cut into chunks small enough for a small model's context, and the
   model lists the strings of each chunk that belong to the categories
   (`AI_CATEGORIES`: people, e-mail addresses, phone numbers, postal
   addresses, companies, identifiers -- or your own). Nothing a model says is
   taken on trust: a string it lists must occur in the chunk (whitespace
   aside) or it is rejected, as are dates, quantities, short and ordinary
   words and the strings to `keep`. Identifiers found more than once in the
   same shape -- `1641B0100-01`, `1641B0230-02` -- become a regular
   expression (`\d{4}B\d{4}-\d{2}`) that also catches the ones the model
   missed. With `model_patterns=True` the model proposes patterns too; one
   is kept only if it compiles, matches its example, and matches nothing but
   identifiers in the text.

3. **Review** the suggestions: `save_suggestions` writes them as JSON, with
   the category, count and type of every string, what each pattern matches,
   and what was rejected and why. Delete what must stay, change a type, add
   what is missing -- the AI is a helper, not the judge.

4. **Obfuscate** with the reviewed file: `load_targets(path)` returns the
   targets dict for `obfuscate_document` / `obfuscate_text`.

From the command line:

    python -m acd.obfuscation_ai suggest *.pdf -o targets.json --chunk-chars 30000
    python -m acd.obfuscation_ai apply targets.json *.pdf --out-dir obfuscated

The text goes to the provider: use a local model for confidential documents.
"""

from argparse import ArgumentParser

from collections import namedtuple

from hashlib import sha1

from json import JSONDecodeError
from json import dump as json_dump
from json import load as json_load
from json import loads as json_loads

from os import environ
from os import fspath
from os import listdir
from os import makedirs
from os import remove
from os import sep
from os import walk
from os.path import abspath
from os.path import basename
from os.path import dirname
from os.path import exists
from os.path import isdir
from os.path import join
from os.path import normcase
from os.path import relpath
from os.path import splitext

from sys import stderr

from typing import Callable
from typing import Optional
from urllib.parse import urlsplit

from regex import IGNORECASE
from regex import compile as regex_compile
from regex import escape as regex_escape
from regex import error as RegexError

from requests import RequestException
from requests import post


AI_DEFAULT_BASE_URL = "http://10.3.1.174:11434"
# gemma4 26B-A4B (IQ4_XS, 135k context, text only): found every string in the
# test documents, 2 PDFs in 9 s; qwen3.5:9b found as many in 79 s.
AI_DEFAULT_MODEL = "gemma4:26b-iq4xs-text-135k"

# Options for an Ollama server: thinking off (qwen3.5:4b thought for more than
# 5 minutes over one line), and JSON mode off -- with it qwen3.5:4b wrote 2
# characters in 3 minutes, without it the whole answer in 2.
AI_OLLAMA_OPTIONS = {"reasoning_effort": "none", "response_format": None}

# Small models answer slowly on a busy server: a chunk can take minutes.
AI_DEFAULT_TIMEOUT = 600

# Characters of text per request. With the prompt and the answer this fits
# the 4096-token context Ollama gives a model by default.
AI_CHUNK_CHARS = 2500

# Most tokens an answer may have: a small model can repeat itself until the
# request times out (qwen3.5:4b did in JSON mode). Unless `options` say else.
AI_MAX_TOKENS = 1500

# category -> (what it is, for the model; obfuscation type; whether strings of
# the same shape become a pattern)
AI_CATEGORIES = {
    "person": ("names of people, also initials with a surname (J. Miller), and signatures", "blackout", False),
    "email": ("e-mail addresses", "blackout", False),
    "phone": ("telephone and fax numbers", "blackout", False),
    "address": ("postal addresses: street with number, postcode with town", "blackout", False),
    "company": ("names of companies, customers, suppliers and organisations", "placeholder", False),
    "identifier": (
        "part numbers, drawing numbers, document numbers, serial numbers, order numbers, CAGE codes, "
        "customer and project codes",
        "placeholder",
        True,
    ),
    "project": (
        "project, programme and aircraft or helicopter type numbers and names (EC 135, H145, BK117), and the "
        "value after a label such as Project, Projekt, Program, Programm, Platform, Root number or Aircraft",
        "placeholder",
        True,
    ),
}

# Dates are not left to the model: these patterns find them, blacked out.
# Day, month, year with dots or slashes (also spaced out by PDF text: "20. 04.
# 2022"), ISO dates, day-month-year with hyphens and a four-digit year, and
# dates with month names in English or German.
_DAY = r"(?:0?[1-9]|[12]\d|3[01])"
_MONTH = r"(?:0?[1-9]|1[0-2])"
_MONTH_NAME = r"(?:jan|feb|m[aä]r|apr|ma[iy]|jun|jul|aug|sep|o[ck]t|nov|de[cz])[a-zä]*\.?"
AI_DATE_PATTERNS = (
    rf"(?<![\w.]){_DAY}\s?[./]\s?{_MONTH}\s?[./]\s?(?:19|20)?\d{{2}}(?![\w.])",
    r"(?<![\w-])(?:19|20)\d{2}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12]\d|3[01])(?![\w-])",
    rf"(?<![\w-]){_DAY}-{_MONTH}-(?:19|20)\d{{2}}(?![\w-])",
    rf"(?i)(?<!\w){_DAY}\.?[\s-]?{_MONTH_NAME}[\s-]?(?:19|20)?\d{{2}}(?!\w)",
    rf"(?i)(?<!\w){_MONTH_NAME}\s{_DAY},?\s(?:19|20)\d{{2}}(?!\w)",
)

# category -> (obfuscation type, patterns) found by these patterns alone.
AI_PATTERN_CATEGORIES = {"date": ("blackout", AI_DATE_PATTERNS)}

AIProvider = namedtuple("AIProvider", "base_url model api_key timeout options")

_SYSTEM_PROMPT = (
    "You find confidential strings in technical documents, so that they can be removed before the "
    "document is shared. You answer with JSON only."
)

_USER_PROMPT = """Categories:
{categories}
{instructions}
List every string in the TEXT below that belongs to one of the categories. Copy each string exactly as \
it is written in the text. Only list proper names and codes: never ordinary words or technical terms of any language, headings, field labels (Part No., DCN-Nr), units, sizes, dates or quantities. \
Codes include project and programme numbers and aircraft or helicopter types (EC 135, H135), even when they look \
like a model name or a short number: list the value after labels such as Project, Projekt, Program, Platform, \
Root number or Aircraft.
Answer as {{"findings": [{{"text": "...", "category": "..."}}]{patterns_key}}}. If there is nothing, \
answer {{"findings": []{patterns_empty}}}.

TEXT:
<<<
{text}
>>>"""

_PATTERNS_PROMPT = (
    'Also give, in "patterns", a regular expression (Python syntax) for every numbering scheme of '
    "identifiers you see, each with one example from the text: "
    '{"regex": "...", "example": "...", "category": "identifier"}.\n'
)

# Words a model may take for a name or a code; never targets on their own.
_ORDINARY = {
    "page", "date", "name", "signature", "sign", "signed", "approved", "checked", "drawn", "rev", "revision",
    "part", "part number", "number", "no", "qty", "quantity", "title", "sheet", "scale", "issue", "item", "status",
    "remarks", "description", "document", "drawing", "customer", "supplier", "company", "address", "phone", "fax",
    "email", "e-mail", "none", "n/a", "na", "yes", "ok", "nok", "pass", "fail", "material", "inspection",
}

# A string a target must at least have: this many letters or digits.
_MIN_SIGNIFICANT = 3

# A shape pattern needs this many digits in it, or it would match ordinary
# short codes ("A4", "M10").
_MIN_PATTERN_DIGITS = 4

# A pattern the model proposes is too broad when it covers more of the text.
_MAX_PATTERN_SHARE = 0.2

# Guard against patterns that backtrack forever (seconds per search).
_PATTERN_TIMEOUT = 2.0

_DATE = regex_compile(r"\d{1,4}[./-]\d{1,2}[./-]\d{1,4}|\d{1,2}\.?[A-Za-z]{3,9}\.?\d{2,4}")  # spaces removed

# Sizes ("1200x800", "744 x 160 mm") and field labels ("DCN-Nr", "lfd. Nr.",
# "Part No.") that small models take for addresses and identifiers.
_DIMENSION = regex_compile(r"\d+(?:[.,]\d+)?(?:\s*[x×*]\s*\d+(?:[.,]\d+)?)+\s*(?:mm|cm|m|in|px)?", IGNORECASE)
_FIELD_LABEL = regex_compile(r".{0,30}?\b(?:nr|no|num|number|nummer)\.?", IGNORECASE)

# Values of a drawing that models take for codes: surface roughness ("Rz 1",
# "Ra 0.8"), diameters with or without a fit ("Ø7H8") and metric threads
# ("M10x1.5").
_TECHNICAL = regex_compile(
    r"R(?:[aqtz]|max|z\s*max)\s*\d{1,3}(?:[.,]\d+)?(?:\s*[µu]m)?"
    r"|[Ø⌀]\s*\d{1,4}(?:[.,]\d+)?(?:\s*[A-Za-z]{1,2}\d{1,2})?"
    r"|M\d{1,2}(?:[.,]\d+)?(?:\s*[x×]\s*\d{1,3}(?:[.,]\d+)?)?",
    IGNORECASE,
)

# What a string of a built-in category must look like, whatever the model says.
_CATEGORY_CHECKS = {
    "person": lambda string: not any(char.isdigit() for char in string) and sum(char.isalpha() for char in string) >= 4,
    "email": lambda string: "@" in string,
    "phone": lambda string: sum(char.isdigit() for char in string) >= 6
    and not any(char.isalpha() for char in string.lower().replace("ext", "").replace("tel", "").replace("fax", "")),
    "identifier": lambda string: any(char.isdigit() for char in string),
    "project": lambda string: any(char.isdigit() for char in string),
    "address": lambda string: regex_compile(r"[^\W\d_]{3}").search(string) is not None,
}

# A field label before the value a model copied along with it.
_LABEL = regex_compile(r"[A-Za-z][A-Za-z. /-]{1,30}\s*:\s*(\S.*)")

# Document information entries that are not the document's own text: dates,
# and the software that made the file ("Crystal Reports", "PDFsharp").
_PDF_INFO_SKIPPED = ("format", "creationDate", "modDate", "trapped", "encryption", "producer", "creator")

_IMAGE_EXTENSIONS = None


def ai_provider(
    base_url: Optional[str] = None,
    model: Optional[str] = None,
    api_key: Optional[str] = None,
    config: Optional[str] = None,
    timeout: Optional[float] = None,
    options: Optional[dict] = None,
) -> AIProvider:
    r"""The AI provider to use: an OpenAI-compatible chat completions API.

    Every setting comes from the argument if given, else from the environment
    (`ACD_AI_BASE_URL`, `ACD_AI_MODEL`, `ACD_AI_API_KEY`, `ACD_AI_TIMEOUT`),
    else from the JSON config file, else from the defaults (the local Ollama
    server, `AI_DEFAULT_MODEL`, no key).

    Args:
        base_url (str, optional): e.g. "https://api.openai.com/v1", or an
            Ollama server "http://10.3.1.174:11434" (a URL without a path
            gets "/v1", and `AI_OLLAMA_OPTIONS`)
        model (str, optional): the model name, as the provider knows it
        api_key (str, optional): sent as a bearer token; local servers need
            none
        config (str, optional): JSON file with any of the keys base_url,
            model, api_key, api_key_env (the name of the environment variable
            that holds the key -- keeps it out of the file), timeout and
            options. Defaults to `ACD_AI_CONFIG`, if set.
        timeout (float, optional): seconds to wait for one answer
        options (dict, optional): extra fields for every request, e.g.
            {"reasoning_effort": "none"} to turn thinking off, or
            {"max_tokens": 1000}; {"response_format": None} turns JSON mode
            off. Merged over those of the config file.

    Returns:
        AIProvider: base_url (ending in the API version, without a slash),
            model, api_key (or None), timeout, options
    """
    config = config or environ.get("ACD_AI_CONFIG")
    settings = {}
    if config:
        with open(fspath(config), encoding="utf-8") as file:
            settings = json_load(file)
        if not isinstance(settings, dict):
            raise ValueError(f"{config}: the AI config must be a JSON object")
    if api_key is None and settings.get("api_key_env"):
        api_key = environ.get(settings["api_key_env"])

    def setting(value, variable, key, default):
        if value is not None:
            return value
        if environ.get(variable):
            return environ[variable]
        return settings.get(key, default)

    base_url = setting(base_url, "ACD_AI_BASE_URL", "base_url", AI_DEFAULT_BASE_URL).strip().rstrip("/")
    ollama = urlsplit(base_url).path in ("", "/")
    if ollama:
        base_url += "/v1"
    merged = dict(AI_OLLAMA_OPTIONS) if ollama else {}
    merged.update(settings.get("options") or {})
    merged.update(options or {})
    return AIProvider(
        base_url=base_url,
        model=setting(model, "ACD_AI_MODEL", "model", AI_DEFAULT_MODEL),
        api_key=setting(api_key, "ACD_AI_API_KEY", "api_key", None) or None,
        timeout=float(setting(timeout, "ACD_AI_TIMEOUT", "timeout", AI_DEFAULT_TIMEOUT)),
        options=merged,
    )


def ai_chat_json(provider: AIProvider, system: str, user: str) -> object:
    """Ask the model and return its answer parsed as JSON.

    JSON mode is asked for, unless the provider's options set
    "response_format" to None; a provider that refuses it (HTTP 400) is asked
    again without. (Ollama's JSON mode can make a model crawl: qwen3.5:4b
    wrote 2 characters in 3 minutes with it, its whole answer in 2 without.) Reasoning a model writes into its answer (<think>...),
    Markdown code fences and text around the JSON are dropped. An answer that
    still is no JSON is asked for once more.

    Raises:
        RuntimeError: the provider answered with an error, or twice without
            JSON
    """
    messages = [{"role": "system", "content": system}, {"role": "user", "content": user}]
    body = {"model": provider.model, "messages": messages, "temperature": 0, "stream": False}
    body.update(provider.options)
    body.setdefault("response_format", {"type": "json_object"})
    if body["response_format"] is None:  # options {"response_format": None}: JSON mode off
        del body["response_format"]
    body.setdefault("max_tokens", AI_MAX_TOKENS)
    for attempt in range(2):
        content = _ai_complete(provider, body)
        try:
            return _ai_parse_json(content)
        except ValueError:
            if attempt:
                raise RuntimeError(f"{provider.model} did not answer with JSON: {content[:300]!r}")
            body = dict(body, messages=messages + [
                {"role": "assistant", "content": content},
                {"role": "user", "content": "That was not valid JSON. Answer with the JSON object only."},
            ])


def _ai_complete(provider: AIProvider, body: dict) -> str:
    headers = {"Content-Type": "application/json"}
    if provider.api_key:
        headers["Authorization"] = f"Bearer {provider.api_key}"
    url = provider.base_url + "/chat/completions"
    response = post(url, json=body, headers=headers, timeout=provider.timeout)
    if response.status_code == 400 and "response_format" in body:
        body.pop("response_format")  # and from the retry for JSON too
        response = post(url, json=body, headers=headers, timeout=provider.timeout)
    if response.status_code != 200:
        raise RuntimeError(f"{url} answered {response.status_code}: {response.text[:500]}")
    try:
        return response.json()["choices"][0]["message"]["content"] or ""
    except (ValueError, KeyError, IndexError, TypeError) as error:
        raise RuntimeError(f"{url} gave no chat completion: {response.text[:500]}") from error


def _ai_parse_json(content: str) -> object:
    """The JSON value in a model's answer, whatever is around it."""
    text = content
    if "</think>" in text:
        text = text.rsplit("</think>", 1)[1]
    text = text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1] if "\n" in text else ""
        text = text.rsplit("```", 1)[0]
    try:
        return json_loads(text)
    except JSONDecodeError:
        pass
    starts = [index for index in (text.find("{"), text.find("[")) if index >= 0]
    if starts:
        start = min(starts)
        end = text.rfind("}" if text[start] == "{" else "]")
        if end > start:
            try:
                return json_loads(text[start:end + 1])
            except JSONDecodeError:
                pass
    raise ValueError("no JSON in the answer")


def obfuscation_source_text(filepath: str, ocr: bool = False) -> str:
    """The text of a document, as the model gets to read it.

    PDF: the text of every page, then the document information (title,
    author ...). With `ocr`, pages without a text layer are read with OCR
    (`line_ocr`). Image files: OCR if `ocr`, else nothing. Anything else is
    read as a text file.
    """
    filepath = fspath(filepath)
    extension = splitext(filepath)[1].lower()
    if extension == ".pdf":
        return _ai_pdf_text(filepath, ocr)
    if extension in _ai_image_extensions():
        if not ocr:
            return ""
        from PIL import Image

        with Image.open(filepath) as image:
            return _ai_ocr_text(image.convert("L"))
    with open(filepath, "rb") as file:
        data = file.read()
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig")
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16")
    for encoding in ("utf-8", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1")


def _ai_pdf_text(filepath: str, ocr: bool) -> str:
    from pymupdf import csGRAY
    from pymupdf import open as pdf_open

    pages = []
    with pdf_open(filepath) as doc:
        for page in doc:
            text = page.get_text()
            for annot in page.annots() or ():
                text += "\n" + (annot.info.get("content") or "")
            for widget in page.widgets() or ():
                text += "\n" + str(widget.field_value or "")
            if ocr and not text.strip():
                from numpy import frombuffer
                from numpy import uint8

                pixmap = page.get_pixmap(dpi=200, colorspace=csGRAY)
                gray = frombuffer(pixmap.samples, dtype=uint8).reshape(pixmap.height, pixmap.stride)[:, :pixmap.width]
                text = _ai_ocr_text(gray)
            pages.append(text)
        pages.append("\n".join(str(value) for key, value in (doc.metadata or {}).items()
                               if value and key not in _PDF_INFO_SKIPPED))
    return "\n".join(pages)


def _ai_ocr_text(gray) -> str:
    from numpy import asarray

    from .line_ocr import ocr_lines

    return "\n".join(line.text for line in ocr_lines(asarray(gray)))


def _ai_image_extensions() -> set:
    global _IMAGE_EXTENSIONS
    if _IMAGE_EXTENSIONS is None:
        from PIL import Image

        _IMAGE_EXTENSIONS = {extension.lower() for extension in Image.registered_extensions()} - {".pdf"}
    return _IMAGE_EXTENSIONS


def suggest_obfuscation_targets(
    sources: list,
    provider: Optional[AIProvider] = None,
    categories: Optional[dict] = None,
    instructions: str = "",
    keep: tuple = (),
    chunk_chars: int = AI_CHUNK_CHARS,
    patterns: bool = True,
    model_patterns: bool = False,
    ocr: bool = False,
    progress: Optional[Callable[[str], None]] = None,
    pattern_categories: Optional[dict] = None,
) -> dict:
    r"""Have the AI model suggest the obfuscation targets of documents.

        suggestions = suggest_obfuscation_targets(["BOM.pdf", "DMC.xml"],
            instructions="The customer is Liebherr-Aerospace; its CAGE code is D9893.")
        save_suggestions(suggestions, "targets.json")  # review, then:
        targets = load_targets("targets.json")
        obfuscate_document("BOM.pdf", targets)

    Args:
        sources (list): file paths (PDF, text, images with `ocr`), or texts
            given as ("name", "text") tuples
        provider (AIProvider, optional): defaults to `ai_provider()`
        categories (dict, optional): category -> (description for the model,
            "blackout" or "placeholder", whether same-shaped strings become a
            pattern). Defaults to `AI_CATEGORIES`.
        instructions (str, optional): told to the model in addition, e.g.
            names it must find, or what is not confidential
        keep (tuple, optional): strings never to obfuscate (case ignored)
        chunk_chars (int, optional): characters of text per request
        patterns (bool, optional): turn identifiers of one shape into a
            regular expression. Defaults to True.
        model_patterns (bool, optional): have the model propose regular
            expressions too (small models write poor ones; each is checked).
            Defaults to False.
        ocr (bool, optional): read scanned pages and images with OCR
        progress (callable, optional): called with a line of text per chunk
        pattern_categories (dict, optional): category -> (obfuscation type,
            regular expressions) found without the model. Defaults to
            `AI_PATTERN_CATEGORIES` (dates, blacked out); {} for none.

    Returns:
        dict: the suggestions, JSON-serialisable -- "provider", "sources"
            (name, characters, chunks), "strings" ([{text, category, type,
            count, sources, spellings}] -- spellings: the same string in
            other case, which is obfuscated too), "patterns" ([{regex, category, type, source,
            examples, matches}]), "rejected" ([{text, category, reason}])
            and "failures" ([{source, chunk, error}]: chunks the model gave
            no usable answer for -- review those parts of the text yourself)

    Raises:
        requests.ConnectionError, RuntimeError: the first request already
            fails to reach the provider, or it answers with an HTTP error
    """
    provider = provider or ai_provider()
    categories = categories or AI_CATEGORIES
    kept = {_ai_key(string) for string in keep}
    texts = []
    for source in sources:
        if isinstance(source, tuple):
            name, text = source
        else:
            name, text = basename(fspath(source)), obfuscation_source_text(source, ocr)
        texts.append((name, text))

    found = {}  # key -> {text, category, type, sources}
    rejected = {}
    proposed = []  # (regex, example, category)
    answers = {}  # chunk digest -> findings: forms repeat, ask once
    report_sources = []
    failures = []  # chunks the model gave no answer for: nobody looked at them
    for name, text in texts:
        chunks = _ai_chunks(text, chunk_chars)
        report_sources.append({"name": name, "characters": len(text), "chunks": len(chunks)})
        for number, chunk in enumerate(chunks, 1):
            digest = sha1(chunk.encode("utf-8")).hexdigest()
            if digest not in answers:
                if progress:
                    progress(f"{name}: chunk {number}/{len(chunks)} ({len(chunk)} characters)")
                try:
                    answers[digest] = _ai_ask_chunk(provider, chunk, categories, instructions, model_patterns)
                except (RuntimeError, RequestException) as error:
                    if not answers and not failures and _ai_unreachable(error):
                        raise  # the provider is down or wrong: no use going on
                    failures.append({"source": name, "chunk": number, "error": str(error)[:500]})
                    if progress:
                        progress(f"{name}: chunk {number} failed: {str(error)[:200]}")
                    continue
            findings, chunk_patterns = answers[digest]
            proposed.extend(chunk_patterns)
            for raw, category in findings:
                category = category if category in categories else _ai_guess_category(category, categories, raw)
                string, reason = _ai_check_finding(raw, category, chunk, kept)
                if reason:
                    rejected.setdefault((raw, reason), {"text": raw, "category": category, "reason": reason})
                    continue
                entry = found.setdefault(_ai_key(string), {
                    "text": string, "category": category, "type": categories[category][1], "sources": [],
                })
                if categories[category][1] == "blackout":
                    entry["type"] = "blackout"  # blackout wins, as in obfuscate_document
                if name not in entry["sources"]:
                    entry["sources"].append(name)

    whole = "\n".join(text for _, text in texts)
    strings = sorted(found.values(), key=lambda entry: (entry["category"], entry["text"].lower()))
    for entry in strings:
        matches = _ai_occurrences(entry["text"], whole)
        entry["count"] = len(matches)
        # obfuscate_document matches a string case-sensitively: every spelling is a target
        entry["spellings"] = sorted(set(matches) - {entry["text"]})
    suggested_patterns = []
    # A project number that is only digits ("1644") becomes a pattern: as a
    # string it would also match inside longer numbers
    for entry in [entry for entry in strings if entry["category"] == "project" and entry["text"].isdigit()]:
        strings.remove(entry)
        regex = rf"(?<![0-9]){entry['text']}(?![0-9])"
        suggested_patterns.append({
            "regex": regex, "category": "project", "type": entry["type"], "source": "number",
            "examples": [entry["text"]], "matches": sorted({m.group() for m in regex_compile(regex).finditer(whole)}),
        })
    for category, (kind, regexes) in (AI_PATTERN_CATEGORIES if pattern_categories is None else pattern_categories).items():
        for regex in regexes:
            matches = sorted({" ".join(m.group().split()) for m in regex_compile(regex).finditer(whole)})
            if matches:
                suggested_patterns.append({
                    "regex": regex, "category": category, "type": kind, "source": "built-in",
                    "examples": matches[:3], "matches": matches,
                })
    if patterns:
        suggested_patterns += _ai_shape_patterns(strings, categories, whole)
    for regex, example, category in proposed:
        category = category if category in categories else "identifier" if "identifier" in categories else None
        if category is None:
            continue
        pattern, reason = _ai_check_pattern(regex, example, whole, found)
        if reason:
            rejected.setdefault((regex, reason), {"text": regex, "category": category, "reason": reason})
        elif all(existing["regex"] != regex for existing in suggested_patterns):
            suggested_patterns.append({
                "regex": regex, "category": category, "type": categories[category][1], "source": "model",
                "examples": [example] if example else [], "matches": pattern,
            })
    return {
        "provider": {"base_url": provider.base_url, "model": provider.model},
        "sources": report_sources,
        "strings": strings,
        "patterns": suggested_patterns,
        "rejected": list(rejected.values()),
        "failures": failures,
    }


def _ai_unreachable(error: Exception) -> bool:
    """Whether an error says the provider cannot be used at all (as opposed
    to one answer that took too long or was no JSON)."""
    from requests import ConnectionError as RequestsConnectionError
    from requests import ReadTimeout

    if isinstance(error, ReadTimeout):
        return False
    return isinstance(error, RequestsConnectionError) or " answered " in str(error)


def _ai_chunks(text: str, size: int) -> list:
    """The text cut at line ends into chunks of at most `size` characters
    (a longer line is cut where it must), each starting with the last line of
    the chunk before -- so a name broken over two lines is seen whole once."""
    lines = []
    for line in text.splitlines():
        line = " ".join(line.split())
        while len(line) > size:
            lines.append(line[:size])
            line = line[size:]
        if line:
            lines.append(line)
    chunks = []
    current = []
    length = 0
    for line in lines:
        if current and length + len(line) + 1 > size:
            chunks.append("\n".join(current))
            current = current[-1:] if len(current[-1]) < size // 2 else []
            length = sum(len(kept) + 1 for kept in current)
        current.append(line)
        length += len(line) + 1
    if current:
        chunks.append("\n".join(current))
    return [chunk for chunk in chunks if any(char.isalnum() for char in chunk)]


def _ai_ask_chunk(provider: AIProvider, chunk: str, categories: dict, instructions: str, model_patterns: bool):
    """([(text, category)], [(regex, example, category)]) the model finds in a chunk."""
    prompt = _USER_PROMPT.format(
        categories="\n".join(f"- {name}: {description}" for name, (description, *_) in categories.items()),
        instructions=(instructions.strip() + "\n") if instructions.strip() else "",
        patterns_key=', "patterns": [...]' if model_patterns else "",
        patterns_empty=', "patterns": []' if model_patterns else "",
        text=chunk,
    )
    if model_patterns:
        prompt = prompt.replace("\nTEXT:", "\n" + _PATTERNS_PROMPT + "\nTEXT:", 1)
    answer = ai_chat_json(provider, _SYSTEM_PROMPT, prompt)
    findings = []
    proposed = []
    items = answer.get("findings", answer.get("items", [])) if isinstance(answer, dict) else answer
    for item in items if isinstance(items, list) else ():
        if isinstance(item, str):
            findings.append((item, ""))
        elif isinstance(item, dict) and isinstance(item.get("text"), str):
            findings.append((item["text"], str(item.get("category") or "").strip().lower()))
    for item in (answer.get("patterns") or ()) if isinstance(answer, dict) else ():
        if isinstance(item, dict) and isinstance(item.get("regex"), str):
            proposed.append((item["regex"], str(item.get("example") or ""), str(item.get("category") or "").lower()))
    return findings, proposed


def _ai_guess_category(category: str, categories: dict, text: str) -> str:
    """A category the model made up (or left out), mapped to the nearest one
    there is."""
    if not category:
        guess = "email" if "@" in text else "identifier" if any(char.isdigit() for char in text) else "company"
        return guess if guess in categories else next(iter(categories))
    for name in categories:
        if name in category or category in name:
            return name
    words = {"name": "person", "signature": "person", "mail": "email", "tel": "phone", "fax": "phone",
             "street": "address", "organi": "company", "customer": "company", "supplier": "company",
             "number": "identifier", "code": "identifier", "serial": "identifier"}
    for word, name in words.items():
        if word in category and name in categories:
            return name
    return "identifier" if "identifier" in categories else next(iter(categories))


def _ai_key(string: str) -> str:
    return " ".join(string.split()).casefold()


def _ai_literal(string: str):
    """A pattern matching `string` whatever the whitespace in it, and in any
    case -- as the model may have changed both."""
    return regex_compile(r"\s*".join(regex_escape(char) for char in "".join(string.split())), IGNORECASE)


def _ai_occurrences(string: str, text: str) -> list:
    """The string's occurrences in the text, as written there: whole words only.

    "LIDA" does not occur in "Validation", nor "LLI" in "will": as a
    spelling to obfuscate, "lida" and "lli" would cut into every such word.
    """
    pattern = regex_compile(r"(?<![^\W_])" + _ai_literal(string).pattern + r"(?![^\W_])", IGNORECASE)
    return [" ".join(match.group().split()) for match in pattern.finditer(text)]


def _ai_check_finding(raw: str, category: str, chunk: str, kept: set) -> tuple:
    """(the string as the text spells it, None), or (None, why it is rejected)."""
    string = " ".join(raw.split()).strip(" .,;:()[]\"'")
    for opening, closing in ("()", "[]"):
        if string.count(opening) != string.count(closing):  # "Jasin (LLI-FN": cut where the model did
            string = string.split(opening if string.count(opening) > string.count(closing) else closing)[0].strip()
    label = _LABEL.match(string)
    if label:  # "DCN Number : LLI_DCN-114242": the value, not the field
        string = label.group(1)
    if sum(char.isalnum() for char in string) < _MIN_SIGNIFICANT:
        return None, "too short"
    match = _ai_literal(string).search(chunk)
    if not match:
        return None, "not in the text"
    string = " ".join(match.group().split())
    key = _ai_key(string)
    if key in kept:
        return None, "to keep"
    if key in _ORDINARY:
        return None, "ordinary word"
    if _DATE.fullmatch("".join(string.split())):
        return None, "date"
    if not any(char.isalpha() for char in string) and sum(char.isdigit() for char in string) < 5 \
            and category != "project":  # a project's root number is short: 1644
        return None, "number"  # quantities, page numbers, years
    if _DIMENSION.fullmatch(string):
        return None, "dimension"
    if _TECHNICAL.fullmatch(string):
        return None, "technical value"
    if _FIELD_LABEL.fullmatch(string):
        return None, "field label"
    check = _CATEGORY_CHECKS.get(category)
    if check and not check(string):
        return None, f"not a{'n' if category[0] in 'aeiou' else ''} {category}"
    return string, None


def _ai_shape(string: str) -> str:
    r"""A pattern for strings shaped like `string`: every run of digits becomes
    \d{n}, letters and punctuation stay, whitespace may vary."""
    parts = []
    index = 0
    while index < len(string):
        char = string[index]
        if char.isdigit():
            end = index
            while end < len(string) and string[end].isdigit():
                end += 1
            parts.append(r"\d" if end - index == 1 else rf"\d{{{end - index}}}")
            index = end
        elif char.isspace():
            while index < len(string) and string[index].isspace():
                index += 1
            parts.append(r"\s*")
        else:
            parts.append(regex_escape(char))
            index += 1
    return r"(?<![A-Za-z0-9])" + "".join(parts) + r"(?![A-Za-z0-9])"


def _ai_shape_patterns(strings: list, categories: dict, whole: str) -> list:
    """A pattern for every shape at least two strings of a pattern category
    share, with every string of the text it matches."""
    shapes = {}
    for entry in strings:
        # A shape of digits alone matches every number that long: two item
        # numbers OCR read as one ("0220 1850" -> 02201850), a sheet's grid
        # labels (910111213). Such strings stay targets by themselves.
        if categories.get(entry["category"], ("", "", False))[2] and not entry["text"].replace(" ", "").isdigit():
            if sum(char.isdigit() for char in entry["text"]) >= _MIN_PATTERN_DIGITS:
                shapes.setdefault((_ai_shape(entry["text"]), entry["category"]), []).append(entry)
    found = []
    for (shape, category), entries in shapes.items():
        if len(entries) < 2:
            continue
        matches = sorted({" ".join(match.group().split()) for match in regex_compile(shape).finditer(whole)})
        found.append({
            "regex": shape, "category": category, "type": entries[0]["type"], "source": "shape",
            "examples": [entry["text"] for entry in entries], "matches": matches,
        })
    return found


def _ai_check_pattern(regex: str, example: str, whole: str, found: dict) -> tuple:
    """(the strings a model's pattern matches, None), or (None, why it is rejected)."""
    try:
        pattern = regex_compile(regex)
    except (RegexError, ValueError, TypeError):
        return None, "invalid regular expression"
    try:
        if pattern.fullmatch("", timeout=_PATTERN_TIMEOUT):
            return None, "matches nothing at all"
        if example and not pattern.fullmatch(example, timeout=_PATTERN_TIMEOUT):
            return None, "does not match its example"
        matches = [match.group() for match in pattern.finditer(whole, timeout=_PATTERN_TIMEOUT) if match.group().strip()]
    except TimeoutError:
        return None, "too slow"
    if not matches:
        return None, "matches nothing in the text"
    if sum(map(len, matches)) > _MAX_PATTERN_SHARE * max(len(whole), 1):
        return None, "too broad"
    for match in matches:
        if sum(char.isalnum() for char in match) < _MIN_SIGNIFICANT:
            return None, f"matches {match!r}"
        if not any(char.isdigit() for char in match) and _ai_key(match) not in found:
            return None, f"matches ordinary text ({match!r})"
    return sorted({" ".join(match.split()) for match in matches}), None


def save_suggestions(suggestions: dict, path: str) -> None:
    """Write the suggestions as JSON, to review and edit before use."""
    with open(fspath(path), "w", encoding="utf-8", newline="\n") as file:
        json_dump(suggestions, file, ensure_ascii=False, indent=2)
        file.write("\n")


def load_targets(suggestions) -> dict:
    """The targets dict for `obfuscate_document`, from suggestions (the dict,
    or the path of the JSON file `save_suggestions` wrote): every string as a
    str key, every pattern compiled -- with its type."""
    if not isinstance(suggestions, dict):
        with open(fspath(suggestions), encoding="utf-8") as file:
            suggestions = json_load(file)
    targets = {}
    for entry in suggestions.get("strings", ()):
        for string in [entry["text"]] + list(entry.get("spellings", ())):
            targets[string] = entry.get("type", "placeholder")
    for entry in suggestions.get("patterns", ()):
        targets[regex_compile(entry["regex"])] = entry.get("type", "placeholder")
    return targets


def _cache_ocr(directory: str) -> None:
    """Keep what `line_ocr.ocr_lines` reads in a folder, by the image's content.

    OCR takes minutes for a scanned drawing; with the cache, suggest and every
    later apply read a page only once (a page changed by removing a logo is a
    new image). `obfuscation` and this module import ocr_lines when they call
    it, so they get the cached one. The files are pickles: only use a folder
    no one else writes to.
    """
    from pickle import dump as pickle_dump
    from pickle import load as pickle_load

    from . import line_ocr

    read = getattr(line_ocr.ocr_lines, "uncached", line_ocr.ocr_lines)
    makedirs(directory, exist_ok=True)

    def ocr_lines(image):
        key = sha1(repr((image.shape, str(image.dtype), line_ocr.OCR_VERSION)).encode() + image.tobytes()).hexdigest()
        path = join(directory, key + ".pickle")
        if exists(path):
            with open(path, "rb") as file:
                return pickle_load(file)
        lines = read(image)
        with open(path, "wb") as file:
            pickle_dump(lines, file)
        return lines

    ocr_lines.uncached = read
    line_ocr.ocr_lines = ocr_lines


def _logo_references(paths) -> list:
    """LogoReferences from image files, and from every image in the folders given."""
    from .logo_obfuscation import load_logo_references

    images = []
    for path in paths:
        if isdir(path):
            images += sorted(join(path, name) for name in listdir(path)
                             if splitext(name)[1].lower() in _ai_image_extensions())
        else:
            images.append(path)
    return load_logo_references(images)


def _obfuscate_file(filepath: str, output: str, targets: dict, placeholders: dict, placeholder_format: str,
                    ocr: bool, references=(), signatures: bool = False, stamps: bool = False) -> dict:
    """Obfuscate one file: its text, then the logos, signatures and stamps found in the original.

    Text first: OCR then reads the pages `suggest` read, which an OCR cache
    has kept. Logos and signatures are found in the original, where no black
    box or placeholder hides or fakes one, and removed from the copy.

    Returns:
        dict: what was done -- "output", "mapping" (original -> replacement),
            "logos", "signatures" and "stamps" (page, box)
    """
    from .obfuscation import PLACEHOLDER_FORMAT
    from .obfuscation import obfuscate_document

    stem, extension = splitext(basename(output))
    is_pdf = extension.lower() == ".pdf"
    makedirs(dirname(abspath(output)), exist_ok=True)
    steps = []
    if references and (is_pdf or extension.lower() in _ai_image_extensions()):
        from .logo_obfuscation import obfuscate_logos

        steps.append(("logos", lambda source, target: obfuscate_logos(source, references, target, find_in=filepath,
                                                                       ocr=ocr)[1]))
    if signatures and is_pdf:
        from .signature_obfuscation import obfuscate_signatures

        steps.append(("signatures", lambda source, target: obfuscate_signatures(source, target, find_in=filepath,
                                                                                 ocr=ocr)[1]))
    if stamps and is_pdf:
        from .stamp_obfuscation import obfuscate_stamps

        steps.append(("stamps", lambda source, target: obfuscate_stamps(source, target, find_in=filepath,
                                                                         ocr=ocr)[1]))
    temporary = [join(dirname(abspath(output)), f".{stem}_step{number}{extension}") for number in range(len(steps))]
    paths = temporary + [output]
    entry = {"output": output}
    try:
        _, entry["mapping"] = obfuscate_document(filepath, targets, paths[0], placeholders=placeholders, ocr=ocr,
                                                 placeholder_format=placeholder_format or PLACEHOLDER_FORMAT)
        for number, (name, step) in enumerate(steps):
            found = step(paths[number], paths[number + 1])
            if name == "logos":
                entry["logos"] = [{"page": logo.page + 1, "box": [round(value, 1) for value in tuple(logo.box)],
                                   "score": logo.score, "reference": logo.reference} for logo in found]
            else:
                entry[name] = [{"page": page + 1, "box": [round(value, 1) for value in tuple(box)]}
                               for page, box in found]
    finally:
        for path in temporary:
            if exists(path):
                remove(path)
    return entry


# Folder codes: no I or O, which read as 1 and 0.
FOLDER_CODE_CHARS = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ"
FOLDER_CODE_LENGTH = 5

# The placeholders of a folder: its code, then a running number of nine digits.
FOLDER_PLACEHOLDER_FORMAT = "PH-{code}-{{:09d}}"


def folder_code(length: int = FOLDER_CODE_LENGTH) -> str:
    """A random code for a folder, e.g. "E3B1Q": 34**5, some 45 million, so that two projects hardly share one."""
    from secrets import choice

    return "".join(choice(FOLDER_CODE_CHARS) for _ in range(length))


def obfuscate_folder(
    folder: str,
    output_dir: str,
    targets,
    key_path: Optional[str] = None,
    code: Optional[str] = None,
    ocr: bool = True,
    logos=(),
    signatures: bool = True,
    stamps: bool = True,
    extensions: tuple = (".pdf",),
    progress: Optional[Callable[[str], None]] = None,
    max_files: Optional[int] = None,
) -> dict:
    r"""Obfuscate every document under a folder, with placeholders unique to the folder.

        key = obfuscate_folder("CMM 67-35-81", "CMM 67-35-81 obfuscated", "targets.json",
                               logos=["logos"])
        # PH-E3B1Q-000000001, PH-E3B1Q-000000002, ... in every file;
        # "CMM 67-35-81 obfuscated.key.json" maps them back:
        text = deobfuscate_text(answer_from_somewhere, key)

    The folder gets a code ("E3B1Q", `folder_code`) and every string one
    placeholder, the same in every file: PH-<code>-<nine-digit number>. Text
    that mixes several projects' placeholders can then be mapped back without
    a mix-up. The copies keep the folder's structure and file names.

    The key -- the code, every placeholder's original, and what was done to
    each file -- is written after every file, next to the output folder (not
    in it: the output is what gets shared, the key must not be). A run that
    stops is resumed by running again: files done are skipped, and their
    placeholders kept.

    Args:
        folder (str): the top folder; its subfolders are included
        output_dir (str): where the obfuscated copies go
        targets: the reviewed suggestions (a JSON path or the dict), or a
            targets dict for `obfuscate_document`
        key_path (str, optional): defaults to "<output_dir>.key.json"
        code (str, optional): the folder code; defaults to the key's, or a
            new random one
        ocr (bool, optional): read scans, images and outlined text with OCR.
            Defaults to True.
        logos (optional): reference images (files or folders) of logos to
            remove (`logo_obfuscation`)
        signatures (bool, optional): black out signature cells
            (`signature_obfuscation`). Defaults to True.
        stamps (bool, optional): black out stamps (`stamp_obfuscation`).
            Defaults to True.
        extensions (tuple, optional): the files to obfuscate. Defaults to PDFs.
        max_files (int, optional): stop after obfuscating this many files;
            the next run goes on from there
        progress (callable, optional): called with a line of text per file

    Returns:
        dict: the key
    """
    folder, output_dir = abspath(fspath(folder)), abspath(fspath(output_dir))
    if normcase(output_dir) == normcase(folder) or normcase(output_dir).startswith(normcase(folder) + sep):
        raise ValueError("output_dir must be outside the folder: the copies would be obfuscated again")
    key_path = fspath(key_path) if key_path else output_dir.rstrip("\\/") + ".key.json"
    if exists(key_path):
        with open(key_path, encoding="utf-8") as file:
            key = json_load(file)
        if code and code != key["code"]:
            raise ValueError(f"{key_path} is for code {key['code']}, not {code}: use another key file")
    else:
        key = {"code": code or folder_code(), "folder": folder, "output": output_dir, "placeholders": {}, "files": {}}
    key["placeholder_format"] = FOLDER_PLACEHOLDER_FORMAT.format(code=key["code"])
    if isinstance(targets, dict) and not ({"strings", "patterns"} & set(targets)):
        rules = targets
    else:
        rules = load_targets(targets)
    references = _logo_references(logos) if logos else ()

    def save():
        makedirs(dirname(abspath(key_path)), exist_ok=True)
        with open(key_path, "w", encoding="utf-8", newline="\n") as file:
            json_dump(key, file, ensure_ascii=False, indent=1)

    files = sorted(join(root, name) for root, _, names in walk(folder) for name in names
                   if splitext(name)[1].lower() in extensions and not name.startswith("."))
    done = 0
    for number, filepath in enumerate(files, 1):
        relative = relpath(filepath, folder).replace(sep, "/")
        output = join(output_dir, relative)
        if "mapping" in key["files"].get(relative, {}) and exists(output):
            continue
        if max_files is not None and done >= max_files:
            break
        done += 1
        if progress:
            progress(f"{number}/{len(files)} {relative}")
        try:
            entry = _obfuscate_file(filepath, output, rules, key["placeholders"], key["placeholder_format"], ocr,
                                    references, signatures, stamps)
        except Exception as error:  # one file that fails does not stop the others
            key["files"][relative] = {"error": f"{type(error).__name__}: {error}"}
            if progress:
                progress(f"{relative}: FAILED {type(error).__name__}: {error}")
        else:
            key["placeholders"].update((original, new) for original, new in entry["mapping"].items()
                                       if new.startswith("PH-"))
            key["files"][relative] = entry
        save()
    save()
    return key


def deobfuscate_text(text: str, key) -> str:
    """Put the originals back for a folder's placeholders in a text (an answer, a report, a translation).

    Args:
        key: the key `obfuscate_folder` wrote (its path or the dict)

    Placeholders of other folders, and blacked-out strings, stay as they are.
    """
    if not isinstance(key, dict):
        with open(fspath(key), encoding="utf-8") as file:
            key = json_load(file)
    originals = {placeholder: original for original, placeholder in key["placeholders"].items()}
    prefix = key["placeholder_format"].split("{")[0]
    return regex_compile(regex_escape(prefix) + r"\d{9}(?!\d)").sub(
        lambda match: originals.get(match.group(), match.group()), text)


def _ai_command_line(arguments: Optional[list] = None) -> None:
    parser = ArgumentParser(prog="python -m acd.obfuscation_ai", description=__doc__.split("\n\n")[0])
    commands = parser.add_subparsers(dest="command", required=True)
    suggest = commands.add_parser("suggest", help="have the AI suggest targets, written as JSON to review")
    suggest.add_argument("files", nargs="+")
    suggest.add_argument("-o", "--output", default="obfuscation_targets.json")
    suggest.add_argument("--base-url")
    suggest.add_argument("--model")
    suggest.add_argument("--api-key")
    suggest.add_argument("--config", help="JSON file with the provider settings")
    suggest.add_argument("--no-think", action="store_true", help='send reasoning_effort "none" (thinking models)')
    suggest.add_argument("--no-json-mode", action="store_true", help="do not ask for JSON mode (slow on some models)")
    suggest.add_argument("--instructions", default="", help="told to the model in addition")
    suggest.add_argument("--keep", action="append", default=[], help="a string never to obfuscate (repeatable)")
    suggest.add_argument("--chunk-chars", type=int, default=AI_CHUNK_CHARS)
    suggest.add_argument("--model-patterns", action="store_true", help="have the model propose patterns too")
    suggest.add_argument("--ocr", action="store_true", help="read scanned pages and images with OCR")
    apply = commands.add_parser("apply", help="obfuscate documents with a reviewed targets file")
    apply.add_argument("targets")
    apply.add_argument("files", nargs="+")
    apply.add_argument("--out-dir", required=True)
    apply.add_argument("--ocr", action="store_true")
    for command in (suggest, apply):
        command.add_argument("--ocr-cache", help="folder to keep OCR results in: a page is then read once, "
                                                 "however often suggest and apply run on it")
    apply.add_argument("--logos", nargs="+", default=[],
                       help="reference images of a logo (files or folders): every similar logo is removed")
    apply.add_argument("--summary", help="write what was obfuscated in each file as JSON (it names the originals)")
    apply.add_argument("--report", help="folder for an HTML report comparing every page with its original")
    apply.add_argument("--resume", action="store_true",
                       help="continue a batch: skip the files --summary lists as done, and keep their placeholders")
    apply.add_argument("--signatures", action="store_true", help="black out the cells signature labels name")
    apply.add_argument("--stamps", action="store_true", help="black out stamps (with --ocr also those on scans)")
    folder = commands.add_parser("folder", help="obfuscate every PDF under a folder, with placeholders unique "
                                                "to it (PH-<code>-000000001) and a key to map them back")
    folder.add_argument("folder")
    folder.add_argument("--out-dir", required=True, help="where the copies go (outside the folder)")
    folder.add_argument("--targets", required=True, help="the reviewed targets file (from suggest)")
    folder.add_argument("--key", help='the key file; defaults to "<out-dir>.key.json"')
    folder.add_argument("--code", help="the folder code; defaults to the key's, or a new random one")
    folder.add_argument("--no-ocr", action="store_true", help="do not read scans and images with OCR")
    folder.add_argument("--logos", nargs="+", default=[], help="reference images of a logo (files or folders)")
    folder.add_argument("--no-signatures", action="store_true", help="keep signature cells")
    folder.add_argument("--no-stamps", action="store_true", help="keep stamps")
    folder.add_argument("--ocr-cache", help="folder to keep OCR results in")
    folder.add_argument("--report", help="folder for an HTML report comparing every page with its original")
    folder.add_argument("--max-files", type=int, help="stop after this many files; run again to go on")
    deobfuscate = commands.add_parser("deobfuscate", help="put the originals back for a folder's placeholders")
    deobfuscate.add_argument("--key", required=True, help="the key file obfuscate wrote for the folder")
    deobfuscate.add_argument("files", nargs="*", help="text files (UTF-8); none: standard input to standard output")
    args = parser.parse_args(arguments)
    if getattr(args, "resume", False) and not args.summary:
        parser.error("--resume needs --summary")
    if getattr(args, "ocr_cache", None):
        _cache_ocr(args.ocr_cache)
    progress = lambda line: print(line, file=stderr, flush=True)  # noqa: E731

    if args.command == "folder":
        key = obfuscate_folder(args.folder, args.out_dir, args.targets, args.key, args.code, not args.no_ocr,
                               args.logos, not args.no_signatures, not args.no_stamps, progress=progress,
                               max_files=args.max_files)
        failed = [name for name, entry in key["files"].items() if "error" in entry]
        print(f"code {key['code']}: {len(key['files'])} files, {len(key['placeholders'])} placeholders, "
              f"{len(failed)} failed", file=stderr)
        if args.report:
            from .obfuscation_report import build_report

            print(f"report: {build_report(args.folder, args.out_dir, args.report, key['files'], progress=progress)}",
                  file=stderr)
        return
    if args.command == "deobfuscate":
        if not args.files:
            from sys import stdin
            from sys import stdout

            stdout.write(deobfuscate_text(stdin.read(), args.key))
        for path in args.files:
            stem, extension = splitext(path)
            with open(path, encoding="utf-8") as file:
                text = deobfuscate_text(file.read(), args.key)
            with open(stem + "_deobfuscated" + extension, "w", encoding="utf-8") as file:
                file.write(text)
            print(f"{stem}_deobfuscated{extension}", file=stderr)
        return

    if args.command == "suggest":
        options = {}
        if args.no_think:
            options["reasoning_effort"] = "none"
        if args.no_json_mode:
            options["response_format"] = None
        provider = ai_provider(args.base_url, args.model, args.api_key, args.config, options=options)
        print(f"{provider.model} at {provider.base_url}", file=stderr)
        suggestions = suggest_obfuscation_targets(
            args.files, provider, instructions=args.instructions, keep=tuple(args.keep),
            chunk_chars=args.chunk_chars, model_patterns=args.model_patterns, ocr=args.ocr,
            progress=lambda line: print(line, file=stderr, flush=True),
        )
        save_suggestions(suggestions, args.output)
        print(f"{len(suggestions['strings'])} strings, {len(suggestions['patterns'])} patterns, "
              f"{len(suggestions['rejected'])} rejected, {len(suggestions['failures'])} chunks failed "
              f"-> {args.output}", file=stderr)
    else:
        targets = load_targets(args.targets)
        makedirs(args.out_dir, exist_ok=True)
        references = _logo_references(args.logos) if args.logos else ()
        placeholders = {}
        summary = {}
        if args.resume and exists(args.summary):  # a batch run before: its placeholders, and its files done
            with open(args.summary, encoding="utf-8") as file:
                summary = json_load(file)
            for entry in summary.values():
                placeholders.update((original, new) for original, new in (entry.get("mapping") or {}).items()
                                    if "\u2588" not in new)

        def save_summary():
            if args.summary:
                makedirs(dirname(abspath(args.summary)), exist_ok=True)
                with open(args.summary, "w", encoding="utf-8") as file:
                    json_dump(summary, file, ensure_ascii=False, indent=1)

        for filepath in args.files:
            stem, extension = splitext(basename(filepath))
            output = join(args.out_dir, stem + "_obfuscated" + extension)
            if args.resume and "mapping" in summary.get(basename(filepath), {}) and exists(output):
                print(f"{output}: done before, skipped", file=stderr)
                continue
            try:
                entry = _obfuscate_file(filepath, output, targets, placeholders, None, args.ocr, references,
                                        args.signatures, args.stamps)
            except Exception as error:  # one file that fails does not stop the others
                summary[basename(filepath)] = {"error": f"{type(error).__name__}: {error}"}
                print(f"{filepath}: FAILED {type(error).__name__}: {error}", file=stderr)
                continue
            # one placeholder per string across all files (blacked-out strings need none)
            placeholders.update((original, new) for original, new in entry["mapping"].items()
                                if "\u2588" not in new)
            summary[basename(filepath)] = entry
            print(f"{output}: {len(entry['mapping'])} strings, {len(entry.get('logos', ()))} logos, "
                  f"{len(entry.get('signatures', ()))} signature areas, {len(entry.get('stamps', ()))} stamps "
                  "obfuscated", file=stderr)
            save_summary()  # after every file: a batch that stops can --resume
        save_summary()
        if args.report:
            from .obfuscation_report import build_report

            folders = {dirname(abspath(filepath)) for filepath in args.files}
            if len(folders) > 1:
                print("--report needs all files in one folder: no report written", file=stderr)
            else:
                report = build_report(folders.pop(), args.out_dir, args.report, summary, progress=progress)
                print(f"report: {report}", file=stderr)


if __name__ == "__main__":
    _ai_command_line()
