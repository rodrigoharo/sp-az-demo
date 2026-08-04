import re
import unicodedata
from pathlib import PurePosixPath


PROTECTED_FOLDER_SEGMENTS: dict[str, str] = {
    "links": "links",
    "document-fonts": "document-fonts",
}


ABBREVIATIONS: tuple[tuple[str, str], ...] = (
    (r"\bsimulados?-?(\d+)\b", r"sml\1"),
    (r"\bsimulados?\b", "sml"),
    (r"\blivro-do-aluno\b", "la"),
    (r"\blivro-do-professor\b", "lp"),
    (r"\bmanual-do-professor\b", "mp"),
    (r"\bquestoes\b", "qst"),
    (r"\bquestao\b", "qst"),
    (r"\bquest\b", "qst"),
    (r"\bresolucoes\b", "res"),
    (r"\bresolucao\b", "res"),
    (r"\bgabarito\b", "gab"),
    (r"\balunos?\b", "la"),
    (r"\balun\b", "la"),
    (r"\bprofessores?\b", "lp"),
    (r"\bprofessor\b", "lp"),
    (r"\bprof\b", "lp"),
    (r"\bmanual\b", "mp"),
    (r"\bcaderno\b", "cad"),
    (r"\bcapa\b", "c"),
    (r"\bmiolo\b", "m"),
    (r"\bmio\b", "m"),
    (r"\bactivity-book\b", "actbook"),
    (r"\bstudent-book\b", "sttbook"),
    (r"\besc-cris\b", "esccris"),
    (r"\bpref-vit\b", "prefvit"),
    (r"\bsim-pub\b", "simpub"),
    (r"\bportugues\b", "port"),
    (r"\blingua-portuguesa\b", "port"),
    (r"\bling-port\b", "port"),
    (r"\bmatematica\b", "mat"),
    (r"\bingles\b", "ingl"),
    (r"\bing\b", "ingl"),
    (r"\bhistoria\b", "hist"),
    (r"\bgeografia\b", "geo"),
    (r"\bciencias\b", "cnt"),
    (r"\bcn\b", "cnt"),
    (r"\blct-cht\b", "chsa"),
    (r"\bch\b", "chsa"),
    (r"\beducacao-fisica\b", "edfisic"),
    (r"\bedfisica\b", "edfisic"),
)


def normalize_dam_filename(filename: str) -> str:
    """Normalize a DAM filename without losing the original extension."""
    clean_name = PurePosixPath(str(filename).strip()).name
    if not clean_name:
        return "arquivo"

    suffix = PurePosixPath(clean_name).suffix
    stem = clean_name[: -len(suffix)] if suffix else clean_name

    normalized_stem = normalize_dam_filename_stem(stem)
    normalized_suffix = normalize_dam_extension(suffix)
    return f"{normalized_stem}{normalized_suffix}"


def normalize_dam_relative_path(relative_path: str) -> str:
    parts = [part for part in relative_path.strip("/").split("/") if part]
    if len(parts) <= 1:
        return "/".join(parts)

    folders = [normalize_dam_folder_segment(part) for part in parts[:-1]]
    return "/".join([*folders, normalize_dam_filename(parts[-1])])


def normalize_dam_folder_segment(value: str) -> str:
    text = strip_accents(str(value).strip()).lower()
    text = re.sub(r"\s*-\s*", "-", text)
    text = re.sub(r"[\s_.]+", "-", text)
    text = re.sub(r"[^a-z0-9_-]", "-", text)
    text = re.sub(r"-{2,}", "-", text)
    text = text.strip("-.")
    if text in PROTECTED_FOLDER_SEGMENTS:
        return PROTECTED_FOLDER_SEGMENTS[text]
    return text or "pasta"


def normalize_dam_name_component(value: str) -> str:
    text = strip_accents(value).lower()
    text = text.replace("º", "").replace("ª", "")
    text = re.sub(r"\s*-\s*", "-", text)
    text = re.sub(r"[\s_]+", "-", text)

    text = remove_noise_tokens(text)
    text = apply_semantic_rules(text)

    text = re.sub(r"[^a-z0-9._-]", "-", text)
    text = re.sub(r"-{2,}", "-", text)
    return text.strip("-.") or "arquivo"


def normalize_dam_filename_stem(value: str) -> str:
    text = normalize_dam_name_component(value)
    if str(value).rstrip().endswith("_") and not text.endswith("-underscore"):
        return f"{text}-underscore"
    return text


def normalize_dam_extension(suffix: str) -> str:
    if not suffix:
        return ""
    return "." + normalize_dam_name_component(suffix.lstrip("."))


def remove_noise_tokens(value: str) -> str:
    text = value
    text = re.sub(r"^brandx-?", "", text)
    text = re.sub(r"^brand-(?=[a-z]{2,})", "", text)
    text = re.sub(r"(^|-)brandx(?=-|$)", r"\1", text)
    text = re.sub(r"-?prod-?\d{2}-?\d{2}$", "", text)
    text = re.sub(r"-(?:2023|2024|2025|2026|2027)$", "", text)
    return text


def apply_semantic_rules(value: str) -> str:
    text = value
    text = re.sub(r"\b(\d{1,2})-?anos?\b", r"\1a", text)
    text = re.sub(r"\b(\d{1,2})a-?serie\b", r"\1s", text)
    text = re.sub(r"\b(\d{1,2})-?serie\b", r"\1s", text)
    text = re.sub(r"\b(\d{1,2})-?semestre\b", r"\1sem", text)
    text = re.sub(r"\bfund-?1\b", "f1", text)
    text = re.sub(r"\bfund-?2\b", "f2", text)
    text = re.sub(r"\bvolume-?(\d+)\b", r"v\1", text)
    text = re.sub(r"\bvol-?(\d+)\b", r"v\1", text)
    text = re.sub(r"\bprova-unica\b", "pu", text)
    text = re.sub(r"\bprova-?(\d+)\b", r"p\1", text)
    text = re.sub(r"\bmodulo-?(\d+)\b", r"m\1", text)
    text = re.sub(r"\bmod-?(\d+)\b", r"m\1", text)
    text = re.sub(r"\betapa-?(\d+)\b", r"e\1", text)

    for pattern, replacement in ABBREVIATIONS:
        text = re.sub(pattern, replacement, text)
    return text


def strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", str(value))
    return "".join(char for char in normalized if not unicodedata.combining(char))

