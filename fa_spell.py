# -*- coding: utf-8 -*-
"""Conservative Persian spelling review backed by the hazm lexicon (MIT).

The lexicon holds mostly base forms, so a word counts as known when it, or a
stem left after removing common inflectional/derivational affixes, is listed;
verbs are recognised from the lexicon's past#present stem pairs.

A word is reported only when every condition holds: it is unknown, has at
least four letters, appears at most twice in the document (repeated unknown
tokens are usually names or terms), and exactly one frequent lexicon word is
one edit away. Findings are review suggestions (confidence B), never applied.
"""
from collections import Counter
from functools import lru_cache
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
LEXICON_DIR = os.path.join(HERE, "lexicon")
ZWNJ = "‌"
LETTERS = "آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهیءأؤئة"
WORD_RE = re.compile("[" + LETTERS + "ۀيكٔ" + ZWNJ + "]+")
_MARKS = re.compile("[ً-ٰٕـ]")

MIN_LENGTH = 4
MAX_DOCUMENT_REPEATS = 2
MIN_SUGGESTION_FREQUENCY = 2000
# A second candidate within this ratio of the best makes the choice ambiguous.
AMBIGUITY_RATIO = 0.2

_SUFFIXES = sorted({
    "ها", "های", "هایی", "هایم", "هایت", "هایش", "هایمان", "هایتان", "هایشان",
    "ان", "یان", "گان", "ات", "جات", "ین",
    "ی", "یی", "ای", "ئی", "ام", "ات", "اش", "مان", "تان", "شان", "م", "ت", "ش",
    "تر", "ترین", "تری", "ترینی",
    "اند", "ایم", "اید", "است", "ند", "ید", "یم", "د", "ه",
    "گی", "گر", "گری", "گران", "مند", "مندی", "وار", "واری", "انه", "ناک", "ستان", "کده", "زار",
    "یت", "یه", "ی‌ها", "ی‌های",
}, key=len, reverse=True)
def _attaches(stem, suffix):
    """Whether a suffix may follow this final letter in standard spelling:
    «دانشجویان» but not «دانشیان»; «کتابم» but «پژوهی‌اش», not «پژوهیش»."""
    # A final «ه» may be pronounced («مشابهی»، «توجهم»), and ezafe after «ا/و»
    # is written «ی» («آلفای»، «دانشجوی»), so only clear mismatches are refused.
    last = stem[-1:]
    if suffix in {"یان", "یی"}:
        return last in "او"
    if suffix == "ان":
        return last not in "او"
    if suffix == "ی":
        return last != "ی"
    if suffix in {"م", "ت", "ش"}:
        return last not in "اوی"
    if suffix == "ای":
        return last in "هی"
    return True


_PLURALS = {"ها", "های", "هایی", "ان", "یان", "ات", "هایش", "هایشان"}
_INFLECTIONS = sorted(_PLURALS | {"ی", "ای", "یی", "تر", "ترین"}, key=len, reverse=True)
# Noun/adjective prefixes; verbal prefixes are handled by _is_verb.
_PREFIXES = ("نا", "بی", "هم", "با", "غیر", "پر", "کم", "ضد", "فرا", "پیش", "پس", "خود")


# Standard Persian keyboard (ISIRI 9147) rows; a substitution is a plausible
# typo only between neighbouring keys or letters that differ just by dots.
_KEY_ROWS = ("ضصثقفغعهخحجچ", "شسیبلاتنمکگ", "ظطزرذدپو")
_DOT_GROUPS = ("بپتث", "جچحخ", "دذ", "رزژ", "سش", "صض", "طظ", "عغ", "فق", "کگ", "یئ", "اآأ", "وؤ", "هة")


def _confusable_pairs():
    position = {}
    for row, (letters, stagger) in enumerate(zip(_KEY_ROWS, (0, 0.25, 0.75))):
        for column, letter in enumerate(letters):
            position[letter] = (column + stagger, row)
    pairs = set()
    for a, (ax, ay) in position.items():
        for b, (bx, by) in position.items():
            if a != b and ((ax - bx) ** 2 + (ay - by) ** 2) ** 0.5 <= 1.2:
                pairs.add((a, b))
    for group in _DOT_GROUPS:
        pairs.update((a, b) for a in group for b in group if a != b)
    return frozenset(pairs)


CONFUSABLE = _confusable_pairs()


def normalize(word):
    word = _MARKS.sub("", word).replace("ي", "ی").replace("ك", "ک").replace("ۀ", "ه")
    return word.strip(ZWNJ)


@lru_cache(maxsize=1)
def lexicon():
    """Return (frequency by word, strong words, past stems, present stems).

    About 160k entries have neither a frequency nor a tag, and some of them
    are typos. They are accepted when written exactly, but only "strong"
    entries (counted or tagged) may serve as stems of derived forms.
    """
    words, strong = {}, set()
    with open(os.path.join(LEXICON_DIR, "fa_words.dat"), encoding="utf-8") as stream:
        for line in stream:
            parts = line.rstrip("\n").split("\t")
            if not parts or not parts[0]:
                continue
            frequency = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
            word = normalize(parts[0])
            words[word] = max(frequency, words.get(word, 0))
            if frequency > 0 or (len(parts) > 2 and parts[2] not in {"", "0"}):
                strong.add(word)
    past, present = set(), set()
    with open(os.path.join(LEXICON_DIR, "fa_verbs.dat"), encoding="utf-8") as stream:
        for line in stream:
            if "#" in line:
                left, right = (normalize(stem) for stem in line.strip().split("#", 1))
                if left:
                    past.add(left)
                if right:
                    present.add(right)
    return words, frozenset(strong), frozenset(past), frozenset(present)


_PAST_ENDINGS = ("", "م", "ی", "یم", "ید", "ند")
_PRESENT_ENDINGS = ("م", "ی", "د", "یم", "ید", "ند")
_PREVERBS = ("بر", "در", "باز", "فرو", "فرا", "ور", "وا")
# Short present stems are common bare subjunctives («کند»، «شود»); other
# two-letter stems need «می/ب/ن» so typos such as «تنیم» are not verbs.
_BARE_SHORT_PRESENT = {"کن", "شو", "ده", "زن", "رو", "گو", "آ", "بر", "دار"}
_TWO_LETTER_HEADS = {"پر", "کم", "بر", "در", "هم", "سر", "بی", "نا", "پس", "خوش", "بد", "نو", "دو", "سه", "تک"}


def _is_verb(word):
    _, _, past, present = lexicon()
    candidates = [(word, False)]
    for preverb in _PREVERBS:
        if word.startswith(preverb) and len(word) - len(preverb) >= 3:
            candidates.append((word[len(preverb):], True))
    for form, prefixed in candidates:
        form = form.replace(ZWNJ, "")
        for head in ("نمی", "می", "ن", "ب", ""):
            if head and not form.startswith(head):
                continue
            rest = form[len(head):]
            for ending in _PAST_ENDINGS:
                if rest.endswith(ending) and (stem := rest[:len(rest) - len(ending)] if ending else rest) in past:
                    if head in {"", "ن", "می", "نمی"}:
                        return True
            if head in {"", "ن"} and rest.endswith(("ن", "ه")) and rest[:-1] in past:
                return True     # infinitive «شدن»، participle «کرده»
            for ending in _PRESENT_ENDINGS:
                if rest.endswith(ending) and (stem := rest[:-len(ending)]) in present:
                    if head or prefixed or len(stem) >= 3 or stem in _BARE_SHORT_PRESENT:
                        return True
            if head == "ب" and rest in present:
                return True     # imperative «بکن»
    return False


def _known_stem(word, depth=0):
    """Strong lexicon word, a verb form, or one more strong affix removal."""
    words, strong, _, _ = lexicon()
    if word in strong or _is_verb(word):
        return True
    if depth >= 2:
        return False
    for suffix in _SUFFIXES:
        if word.endswith(suffix) and len(word) - len(suffix) >= 3:
            stem = word[:-len(suffix)]
            if not _attaches(stem, suffix):
                continue
            if _known_stem(stem, depth + 1):
                return True
            # Unranked entries may carry a plural, never a one-letter suffix.
            if suffix in _PLURALS and stem in words:
                return True
            # خواننده → خوانندگان / آسوده → آسودگی
            if suffix in {"گان", "گی", "گر"} and stem + "ه" in strong:
                return True
    for prefix in _PREFIXES:
        if word.startswith(prefix) and len(word) - len(prefix) >= 3 and _known_stem(word[len(prefix):], depth + 1):
            return True
    return False


def _frequent(word):
    words, _, _, _ = lexicon()
    return words.get(word, 0) >= 1000


@lru_cache(maxsize=200000)
def is_known(word):
    words, _, _, _ = lexicon()
    if word in words or _known_stem(word):
        return True
    if ZWNJ in word:
        return all(part in words or _known_stem(part) or part in _SUFFIXES or part in _PREFIXES
                   for part in word.split(ZWNJ) if part)
    # Closed compounds: «برخط»، «پرخطا»، «بیدارشدن»، «آزادکردن».
    for cut in range(2, len(word) - 1):
        left, right = word[:cut], word[cut:]
        head_ok = left in _TWO_LETTER_HEADS if len(left) == 2 else _frequent(left)
        tail_ok = (_frequent(right) if len(right) >= 3 or len(left) == 2 else False) or (
            len(right) >= 3 and (_is_verb(right) or _known_stem(right)))
        if head_ok and tail_ok:
            return True
    return False


def _edits(word):
    """Single realistic typing slips: drop, double, swap, or a confusable key."""
    splits = [(word[:i], word[i:]) for i in range(len(word) + 1)]
    for left, right in splits:
        if right:
            yield left + right[1:]
            for letter in LETTERS:
                if (right[0], letter) in CONFUSABLE:
                    yield left + letter + right[1:]
        if len(right) > 1:
            yield left + right[1] + right[0] + right[2:]
        # The writer may have dropped any letter (deletions above cover an
        # extra letter).
        for letter in LETTERS:
            yield left + letter + right


def _base_frequency(word):
    """Frequency of a lexicon word, or of the stem a known inflection reduces to."""
    words, strong, _, _ = lexicon()
    if word in strong:
        return words[word]
    # Only plain inflections of a real stem; derived splits such as
    # «کو+گران» would turn any short word into a frequent suggestion.
    for suffix in _INFLECTIONS:
        stem = word[:-len(suffix)]
        if word.endswith(suffix) and len(stem) >= 4 and stem in strong and _attaches(stem, suffix):
            return words[stem]
    return 0


@lru_cache(maxsize=50000)
def suggestion(word):
    """Return the confident correction for an unknown word, or None.

    Two forms of the same word («پژوهش / پژوهشی» for «پژوهیش») are offered
    together; any other close competition means no suggestion at all.
    """
    if len(word) < MIN_LENGTH or ZWNJ in word or is_known(word):
        return None
    scored = {}
    for candidate in set(_edits(word)):
        frequency = _base_frequency(candidate)
        if frequency:
            scored[candidate] = frequency
    if not scored:
        return None
    ranked = sorted(scored.items(), key=lambda item: -item[1])
    best, frequency = ranked[0]
    if frequency < MIN_SUGGESTION_FREQUENCY:
        return None
    rivals = [candidate for candidate, score in ranked[1:] if score >= frequency * AMBIGUITY_RATIO]
    if not rivals:
        return best
    if len(rivals) == 1 and scored[rivals[0]] == frequency and (
            rivals[0].startswith(best) or best.startswith(rivals[0])):
        return " / ".join(sorted((best, rivals[0]), key=len))
    return None


def document_counts(texts):
    counts = Counter()
    for text in texts:
        counts.update(normalize(match.group()) for match in WORD_RE.finditer(text))
    return counts


def find_misspellings(text, counts=None, protected=(), terms=()):
    """Yield (start, end, original, suggestion) for confident misspellings."""
    terms = {normalize(term) for term in terms}
    for match in WORD_RE.finditer(text):
        word = normalize(match.group())
        if not word or word in terms:
            continue
        if counts is not None and counts.get(word, 0) > MAX_DOCUMENT_REPEATS:
            continue
        if any(not (match.end() <= span.start or match.start() >= span.end) for span in protected):
            continue
        fix = suggestion(word)
        if fix:
            yield match.start(), match.end(), match.group(), fix
