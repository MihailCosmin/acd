"""AI suggestions of obfuscation targets (`acd.obfuscation_ai`).

The provider is a small OpenAI-compatible server run in a thread, which
answers as a model would -- including the mistakes small models make -- so
every request goes over real HTTP. The live tests ask a real model; they run
only when ACD_AI_LIVE is set (with ACD_AI_BASE_URL / ACD_AI_MODEL, or the
defaults: the local Ollama server).
"""

import json
import os
import re
import threading
from http.server import BaseHTTPRequestHandler
from http.server import ThreadingHTTPServer

import pytest

from acd import obfuscation_ai
from acd.obfuscation_ai import ai_chat_json
from acd.obfuscation_ai import ai_provider
from acd.obfuscation_ai import load_targets
from acd.obfuscation_ai import save_suggestions
from acd.obfuscation_ai import suggest_obfuscation_targets


class FakeModel:
    """An OpenAI-compatible chat completions server; `answer(body)` returns
    (status, content) for every request, and `requests` records them."""

    def __init__(self, answer):
        self.answer = answer
        self.requests = []
        model = self

        class Handler(BaseHTTPRequestHandler):
            def do_POST(self):
                body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
                model.requests.append({"path": self.path, "headers": dict(self.headers), "body": body})
                status, content = model.answer(body)
                if status == 200:
                    payload = {"choices": [{"message": {"role": "assistant", "content": content}}]}
                else:
                    payload = {"error": {"message": content}}
                data = json.dumps(payload).encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def log_message(self, *args):
                pass

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.base_url = f"http://127.0.0.1:{self.server.server_address[1]}/v1"

    def provider(self, **settings):
        return ai_provider(base_url=self.base_url, model="fake", **settings)

    def close(self):
        self.server.shutdown()
        self.server.server_close()


@pytest.fixture
def fake_model():
    models = []

    def start(answer):
        models.append(FakeModel(answer))
        return models[-1]

    yield start
    for model in models:
        model.close()


@pytest.fixture
def clean_environment(monkeypatch):
    for variable in ("ACD_AI_BASE_URL", "ACD_AI_MODEL", "ACD_AI_API_KEY", "ACD_AI_TIMEOUT", "ACD_AI_CONFIG"):
        monkeypatch.delenv(variable, raising=False)
    return monkeypatch


def findings(*items):
    return 200, json.dumps({"findings": [{"text": text, "category": category} for text, category in items]})


def chunk_of(body):
    """The text a request asks about."""
    return body["messages"][1]["content"].split("<<<\n", 1)[1].rsplit("\n>>>", 1)[0]


def test_provider_defaults_to_the_local_ollama_server(clean_environment):
    provider = ai_provider()
    assert provider.base_url == "http://10.3.1.174:11434/v1"
    assert provider.model == obfuscation_ai.AI_DEFAULT_MODEL
    assert provider.api_key is None
    assert provider.options == {"reasoning_effort": "none", "response_format": None}  # Ollama: no thinking, no JSON mode


def test_provider_base_url_gets_v1_only_without_a_path(clean_environment):
    assert ai_provider("http://localhost:11434/").base_url == "http://localhost:11434/v1"
    assert ai_provider("https://api.openai.com/v1/").base_url == "https://api.openai.com/v1"
    assert ai_provider("https://api.openai.com/v1/").options == {}
    assert ai_provider("http://localhost:11434", options={"response_format": {"type": "json_object"}}).options == {
        "reasoning_effort": "none", "response_format": {"type": "json_object"}}
    assert ai_provider("http://host:8000/openai/v1").base_url == "http://host:8000/openai/v1"


def test_provider_settings_come_from_argument_then_environment_then_config(clean_environment, tmp_path):
    config = tmp_path / "ai.json"
    config.write_text(json.dumps({
        "base_url": "https://api.example.com/v1", "model": "from-config", "api_key_env": "MY_AI_KEY",
        "timeout": 30, "options": {"reasoning_effort": "none", "max_tokens": 500},
    }), encoding="utf-8")
    clean_environment.setenv("MY_AI_KEY", "secret-from-env")
    clean_environment.setenv("ACD_AI_MODEL", "from-environment")
    provider = ai_provider(config=str(config), options={"max_tokens": 900})
    assert provider.base_url == "https://api.example.com/v1"
    assert provider.model == "from-environment"
    assert provider.api_key == "secret-from-env"
    assert provider.timeout == 30
    assert provider.options == {"reasoning_effort": "none", "max_tokens": 900}
    assert ai_provider(model="from-argument", config=str(config)).model == "from-argument"

    clean_environment.setenv("ACD_AI_CONFIG", str(config))
    assert ai_provider().base_url == "https://api.example.com/v1"


def test_request_carries_model_key_json_mode_and_options(fake_model):
    model = fake_model(lambda body: (200, '{"ok": true}'))
    provider = model.provider(api_key="sk-test", options={"reasoning_effort": "none"})
    assert ai_chat_json(provider, "system", "user") == {"ok": True}
    request = model.requests[0]
    assert request["path"] == "/v1/chat/completions"
    assert request["headers"]["Authorization"] == "Bearer sk-test"
    assert request["body"]["model"] == "fake"
    assert request["body"]["temperature"] == 0
    assert request["body"]["response_format"] == {"type": "json_object"}
    assert request["body"]["reasoning_effort"] == "none"
    assert request["body"]["max_tokens"] == obfuscation_ai.AI_MAX_TOKENS
    assert [message["role"] for message in request["body"]["messages"]] == ["system", "user"]


def test_json_mode_can_be_turned_off(fake_model):
    model = fake_model(lambda body: (200, '{"ok": true}'))
    ai_chat_json(model.provider(options={"response_format": None}), "system", "user")
    assert "response_format" not in model.requests[0]["body"]


def test_no_key_no_authorization_header(fake_model):
    model = fake_model(lambda body: (200, "{}"))
    ai_chat_json(model.provider(), "system", "user")
    assert "Authorization" not in model.requests[0]["headers"]


def test_provider_without_json_mode_is_asked_again_without(fake_model):
    model = fake_model(lambda body: (400, "response_format not supported") if "response_format" in body
                       else (200, '{"findings": []}'))
    assert ai_chat_json(model.provider(), "system", "user") == {"findings": []}
    assert len(model.requests) == 2
    assert "response_format" not in model.requests[1]["body"]


@pytest.mark.parametrize("content", [
    '```json\n{"a": 1}\n```',
    '<think>Let me list them: {"wrong": 0}</think>\n{"a": 1}',
    'Here is the JSON you asked for:\n{"a": 1}\nHope this helps.',
])
def test_answer_around_the_json_is_dropped(fake_model, content):
    model = fake_model(lambda body: (200, content))
    assert ai_chat_json(model.provider(), "system", "user") == {"a": 1}


def test_answer_without_json_is_asked_for_once_more(fake_model):
    answers = iter(["Sure! The names are Heike and Jo.", '{"a": 1}'])
    model = fake_model(lambda body: (200, next(answers)))
    assert ai_chat_json(model.provider(), "system", "user") == {"a": 1}
    assert model.requests[1]["body"]["messages"][-2]["content"].startswith("Sure!")


def test_errors_and_twice_no_json_raise(fake_model):
    model = fake_model(lambda body: (500, "model not found"))
    with pytest.raises(RuntimeError, match="500"):
        ai_chat_json(model.provider(), "system", "user")
    model = fake_model(lambda body: (200, "no json here"))
    with pytest.raises(RuntimeError, match="did not answer with JSON"):
        ai_chat_json(model.provider(), "system", "user")


TEXT = """INSPECTION SHEET
Part Number: 1641B0100-01   Drawing: 1641B0230-02
Customer: Liebherr-Aerospace Lindenberg GmbH, Pfänderstraße 50-52, 88161 Lindenberg
Inspected by Heike   Wochner on 12.03.2024, checked by J. Miller
Released 20. 04. 2022   Qty 4   Box 1200x800, DCN-Nr, Palette   Page 1 of 2
Contact: heike.wochner@example.com, +49 8381 46 0
Also affected: 1641B0777-03 (not seen by the model)"""


def test_suggestions_are_checked_against_the_text(fake_model):
    model = fake_model(lambda body: findings(
        ("Heike Wochner", "person"),                # whitespace differs from the text
        ("j. miller", "person"),                    # case differs
        ("Hans Muster", "person"),                  # made up
        ("Liebherr-Aerospace Lindenberg GmbH", "company"),
        ("1641B0100-01", "part number"),            # a category the model made up
        ("1641B0230-02", "identifier"),
        ("12.03.2024", "identifier"),               # a date
        ("20. 04. 2022", "identifier"),             # a date, spaced out
        ("Drawing: 1641B0230-02", "identifier"),    # with its field label
        ("Page", "identifier"),                     # an ordinary word
        ("4", "identifier"),                        # a quantity
        ("heike.wochner@example.com", "email"),
        ("+49 8381 46 0", "phone"),
        ("1200x800", "address"),                    # a size
        ("Box 744 x 160 mm", "address"),            # not a size alone: no address either
        ("DCN-Nr", "company"),                      # a field label
        ("J. Miller 2", "person"),                  # not in the text anyway
        ("Lindenberg", "phone"),                    # no phone number
        ("Palette", "identifier"),                  # an ordinary word, no code
        ("J. Miller (Lieb", "person"),              # cut off in the middle
    ))
    suggestions = suggest_obfuscation_targets([("sheet.txt", TEXT)], model.provider())
    strings = {entry["text"]: entry for entry in suggestions["strings"]}
    assert set(strings) == {
        "Heike Wochner", "J. Miller", "Liebherr-Aerospace Lindenberg GmbH", "1641B0100-01", "1641B0230-02",
        "heike.wochner@example.com", "+49 8381 46 0",
    }
    assert strings["J. Miller"]["type"] == "blackout"
    assert strings["1641B0100-01"]["category"] == "identifier"
    assert strings["1641B0100-01"]["type"] == "placeholder"
    assert strings["Heike Wochner"]["count"] == 1
    assert strings["Heike Wochner"]["sources"] == ["sheet.txt"]
    reasons = {entry["text"]: entry["reason"] for entry in suggestions["rejected"]}
    assert reasons == {"Hans Muster": "not in the text", "12.03.2024": "date", "20. 04. 2022": "date",
                       "Page": "ordinary word", "4": "too short", "1200x800": "dimension",
                       "Box 744 x 160 mm": "not in the text", "DCN-Nr": "field label",
                       "J. Miller 2": "not in the text", "Lindenberg": "not a phone",
                       "Palette": "not an identifier"}

    # the two part numbers share a shape: the pattern also finds the third
    [pattern] = suggestions["patterns"]
    assert pattern["source"] == "shape"
    assert pattern["matches"] == ["1641B0100-01", "1641B0230-02", "1641B0777-03"]
    assert re.compile(pattern["regex"]).search("x 1641B9999-99 y")
    assert not re.compile(pattern["regex"]).search("1641B9999-999")


def test_every_spelling_is_a_target_and_findings_without_category_are_placed(fake_model):
    obfuscation = pytest.importorskip("acd.obfuscation")
    text = "Signed: Heike Wochner\nHEIKE WOCHNER, QA\nContact heike@example.com about 1641B0100-01."
    model = fake_model(lambda body: (200, json.dumps({"findings": ["Heike Wochner", "heike@example.com",
                                                                   {"text": "1641B0100-01"}]})))
    suggestions = suggest_obfuscation_targets([("a.txt", text)], model.provider())
    strings = {entry["text"]: entry for entry in suggestions["strings"]}
    assert strings["Heike Wochner"]["spellings"] == ["HEIKE WOCHNER"]
    assert strings["Heike Wochner"]["count"] == 2
    assert strings["heike@example.com"]["category"] == "email"
    assert strings["1641B0100-01"]["category"] == "identifier"
    assert strings["Heike Wochner"]["category"] == "company"  # no category given: a placeholder, not lost
    obfuscated, _ = obfuscation.obfuscate_text(text, load_targets(suggestions))
    assert "wochner" not in obfuscated.lower()


def test_keep_and_own_categories(fake_model):
    model = fake_model(lambda body: findings(("Liebherr-Aerospace Lindenberg GmbH", "company"),
                                             ("Heike Wochner", "person")))
    categories = {"person": ("names of people", "placeholder", False)}
    suggestions = suggest_obfuscation_targets([("sheet.txt", TEXT)], model.provider(), categories=categories,
                                              keep=("liebherr-aerospace lindenberg gmbh",),
                                              instructions="Customer names are not confidential.")
    assert [(entry["text"], entry["type"]) for entry in suggestions["strings"]] == [("Heike Wochner", "placeholder")]
    assert suggestions["rejected"][0]["reason"] == "to keep"
    prompt = model.requests[0]["body"]["messages"][1]["content"]
    assert "- person: names of people" in prompt
    assert "company" not in prompt.split("TEXT:")[0]
    assert "Customer names are not confidential." in prompt


def test_long_text_is_asked_chunk_by_chunk_and_repeats_once(fake_model):
    line = "Approved by Heike Wochner for order 4711-0815 of the program. " * 2
    page = "\n".join(f"{number:03d} {line}" for number in range(40))
    model = fake_model(lambda body: findings(*[(name, "person") for name in ("Heike Wochner",)
                                               if name in chunk_of(body)]))
    suggestions = suggest_obfuscation_targets([("a.txt", page), ("b.txt", page)], model.provider(), chunk_chars=1000)
    chunks = suggestions["sources"][0]["chunks"]
    assert chunks > 4
    assert suggestions["sources"][1]["chunks"] == chunks
    assert len(model.requests) == chunks  # the second document is the same text: no new requests
    assert all(len(chunk_of(request["body"])) <= 1000 for request in model.requests)
    [entry] = suggestions["strings"]
    assert entry["sources"] == ["a.txt", "b.txt"]
    assert entry["count"] == 160  # in both documents


def test_a_chunk_without_answer_is_reported_and_the_rest_goes_on(fake_model):
    page = "\n".join(f"Line {number:03d} approved by Heike Wochner, order 4711-0815-{number:03d}." for number in range(60))

    def answer(body):
        if "Line 030" in chunk_of(body):
            return 200, "I cannot help with that."  # never JSON
        return findings(("Heike Wochner", "person"))

    model = fake_model(answer)
    suggestions = suggest_obfuscation_targets([("a.txt", page)], model.provider(), chunk_chars=600)
    failures = suggestions["failures"]
    assert 1 <= len(failures) <= 2  # line 030 may be the line two chunks share
    assert all(failure["source"] == "a.txt" and "did not answer with JSON" in failure["error"]
               for failure in failures)
    assert len(model.requests) == 2 * len(failures) + suggestions["sources"][0]["chunks"] - len(failures)
    assert [entry["text"] for entry in suggestions["strings"]] == ["Heike Wochner"]


def test_unusable_provider_fails_at_once(fake_model):
    model = fake_model(lambda body: (404, "model 'fake' not found"))
    with pytest.raises(RuntimeError, match="404"):
        suggest_obfuscation_targets([("a.txt", "Line 1\n" * 2000)], model.provider())
    assert len(model.requests) == 1  # no use asking about the other chunks


def test_chunks_overlap_by_a_line():
    text = "\n".join(f"line {number} " + "x" * 50 for number in range(20))
    chunks = obfuscation_ai._ai_chunks(text, 200)
    for before, after in zip(chunks, chunks[1:]):
        assert after.split("\n")[0] == before.split("\n")[-1]
    assert all(len(chunk) <= 200 for chunk in chunks)
    assert obfuscation_ai._ai_chunks("x" * 450, 200) == ["x" * 200, "x" * 200, "x" * 50]
    assert obfuscation_ai._ai_chunks(" \n -- \n", 200) == []


def test_model_patterns_are_checked(fake_model):
    patterns = [
        {"regex": r"\d{4}B\d{4}-\d{2}", "example": "1641B0100-01", "category": "identifier"},  # good
        {"regex": r"Lindenberg\b", "example": "Lindenberg", "category": "identifier"},  # ordinary words
        {"regex": r"(\d{4}B", "example": "1641B", "category": "identifier"},  # invalid
        {"regex": r"\d+", "example": "1641B0100-01", "category": "identifier"},  # misses its example
        {"regex": r"ZZ\d{9}", "example": "", "category": "identifier"},  # matches nothing
        {"regex": r"\w*", "example": "", "category": "identifier"},  # matches the empty string
    ]
    model = fake_model(lambda body: (200, json.dumps({"findings": [], "patterns": patterns})))
    suggestions = suggest_obfuscation_targets([("sheet.txt", TEXT)], model.provider(), model_patterns=True)
    assert '"patterns"' in model.requests[0]["body"]["messages"][1]["content"]
    [kept] = suggestions["patterns"]
    assert kept["regex"] == r"\d{4}B\d{4}-\d{2}"
    assert kept["source"] == "model"
    assert kept["matches"] == ["1641B0100-01", "1641B0230-02", "1641B0777-03"]
    reasons = {entry["text"]: entry["reason"] for entry in suggestions["rejected"]}
    assert reasons[r"Lindenberg\b"].startswith("matches ordinary text")
    assert reasons[r"(\d{4}B"] == "invalid regular expression"
    assert reasons[r"\d+"] == "does not match its example"
    assert reasons[r"ZZ\d{9}"] == "matches nothing in the text"
    assert reasons[r"\w*"] == "matches nothing at all"


def test_reviewed_file_feeds_obfuscation(fake_model, tmp_path):
    obfuscation = pytest.importorskip("acd.obfuscation")
    model = fake_model(lambda body: findings(("Heike Wochner", "person"), ("1641B0100-01", "identifier"),
                                             ("1641B0230-02", "identifier"), ("J. Miller", "person")))
    path = tmp_path / "targets.json"
    save_suggestions(suggest_obfuscation_targets([("sheet.txt", TEXT)], model.provider()), path)

    # review: J. Miller may stay
    reviewed = json.loads(path.read_text(encoding="utf-8"))
    reviewed["strings"] = [entry for entry in reviewed["strings"] if entry["text"] != "J. Miller"]
    path.write_text(json.dumps(reviewed), encoding="utf-8")

    text, mapping = obfuscation.obfuscate_text(TEXT, load_targets(str(path)))
    assert "Wochner" not in text and "1641B" not in text
    assert "J. Miller" in text
    assert mapping["1641B0777-03"].startswith("PH-")  # found by the pattern only
    assert set(mapping.values()) >= {"█████ ███████"}


def test_documents_are_read_from_files(fake_model, tmp_path):
    pymupdf = pytest.importorskip("pymupdf")
    pdf = tmp_path / "sheet.pdf"
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_text((50, 72), "Inspected by Heike Wochner")
    doc.set_metadata({"author": "Hans Autor", "producer": "Crystal Reports"})
    doc.save(pdf)
    xml = tmp_path / "dm.xml"
    xml.write_bytes("<para>Geprüft: Jörg Müller</para>".encode("cp1252"))
    seen = []
    model = fake_model(lambda body: (seen.append(chunk_of(body)), (200, '{"findings": []}'))[1])
    suggestions = suggest_obfuscation_targets([pdf, str(xml)], model.provider())
    assert [source["name"] for source in suggestions["sources"]] == ["sheet.pdf", "dm.xml"]
    assert "Heike Wochner" in seen[0] and "Hans Autor" in seen[0]
    assert "Crystal Reports" not in seen[0]  # the software that made the file is no secret
    assert "Jörg Müller" in seen[1]


def test_command_line_suggest_and_apply(fake_model, tmp_path, clean_environment):
    pytest.importorskip("acd.obfuscation")
    source = tmp_path / "sheet.txt"
    source.write_text(TEXT, encoding="utf-8")
    model = fake_model(lambda body: findings(("Heike Wochner", "person")))
    targets = tmp_path / "targets.json"
    obfuscation_ai._ai_command_line(["suggest", str(source), "-o", str(targets), "--base-url", model.base_url,
                                     "--model", "fake", "--no-think", "--no-json-mode", "--keep", "J. Miller"])
    assert model.requests[0]["body"]["reasoning_effort"] == "none"
    assert "response_format" not in model.requests[0]["body"]
    assert json.loads(targets.read_text(encoding="utf-8"))["strings"][0]["text"] == "Heike Wochner"
    out = tmp_path / "out"
    obfuscation_ai._ai_command_line(["apply", str(targets), str(source), "--out-dir", str(out)])
    assert "Wochner" not in (out / "sheet_obfuscated.txt").read_text(encoding="utf-8")


LIVE = pytest.mark.skipif(not os.environ.get("ACD_AI_LIVE"), reason="set ACD_AI_LIVE to ask a real model")


@LIVE
def test_live_model_finds_names_and_part_numbers():
    provider = ai_provider()  # an Ollama base URL turns thinking and JSON mode off
    suggestions = suggest_obfuscation_targets([("sheet.txt", TEXT)], provider)
    texts = {entry["text"] for entry in suggestions["strings"]}
    assert all(obfuscation_ai._ai_literal(text).search(TEXT) for text in texts)  # checked against the text
    assert {"Heike Wochner", "1641B0100-01"} <= texts
