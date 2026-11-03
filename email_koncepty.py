"""Monthly report e-mails: copy the templates into analytik@ Drafts with the month changed.

Templates: Outlook folder "Šablony reportů" in martin.koprnicky@vlmedia.cz, one e-mail per report,
any month written in them. Result: a copy of each in the Drafts ("Koncepty") of analytik@vlmedia.cz
with the month changed in the subject and the body (every Czech case form, original capitalisation,
the year right after the month too), no Re:/Fw: prefix and analytik@ as the sender. Layout, images,
links and recipients stay as they are. A template without a body is skipped.

ONLY DRAFTS, NOTHING DELETED. Every Outlook object goes through OutlookGuard, which raises before
any send-like or delete-like method can be reached; no such call exists in this file. Templates and
existing drafts are only read; the only write is the new draft.

Needs classic Outlook with both mailboxes in its folder list. Starter cell: Tableau_Asyn_MM.ipynb.
"""
import html
import re

import win32com.client

TEMPLATE_MAILBOX = "martin.koprnicky@vlmedia.cz"
TEMPLATE_FOLDER = "Šablony reportů"
TARGET_MAILBOX = "analytik@vlmedia.cz"
OL_FOLDER_DRAFTS = 16
OL_MAIL = 43

# Czech month forms: 0 = nominative/accusative ("za září"), 1 = genitive ("do září", "1. září"),
# 2 = dative/locative ("v září"), 3 = instrumental ("před zářím").
MONTHS = [
    ("leden", "ledna", "lednu", "lednem"),
    ("únor", "února", "únoru", "únorem"),
    ("březen", "března", "březnu", "březnem"),
    ("duben", "dubna", "dubnu", "dubnem"),
    ("květen", "května", "květnu", "květnem"),
    ("červen", "června", "červnu", "červnem"),
    ("červenec", "července", "červenci", "červencem"),
    ("srpen", "srpna", "srpnu", "srpnem"),
    ("září", "září", "září", "zářím"),
    ("říjen", "října", "říjnu", "říjnem"),
    ("listopad", "listopadu", "listopadu", "listopadem"),
    ("prosinec", "prosince", "prosinci", "prosincem"),
]
FORMS = {}  # "října" -> {1}: which case forms a word can be (září is 0, 1 and 2 at once)
for _forms in MONTHS:
    for _i, _form in enumerate(_forms):
        FORMS.setdefault(_form, set()).add(_i)

GEN_WORDS = {"do", "od", "z", "ze", "během", "konce", "začátku", "polovina", "poloviny"}
LOC_WORDS = {"v", "ve", "o", "po", "při", "k", "ke"}

_month_alt = "|".join(sorted(FORMS, key=len, reverse=True))
# month word, then optionally the year after it (spaces, &nbsp; or tags in between, or nothing)
MONTH_RE = re.compile(
    rf"\b(?P<month>{_month_alt})(?![^\W\d_])"
    r"(?:(?P<gap>(?:\s|\xa0|&nbsp;|&#160;|<[^>]*>){0,20})(?P<year>(?:19|20)\d\d)\b)?",
    re.IGNORECASE)
ENTITY_RE = re.compile(r"&(#\d+|#x[0-9a-fA-F]+|[A-Za-z]+);")


class OutlookGuard:
    """Wraps an Outlook COM object. Any send-like or delete-like attribute raises before Outlook sees it.

    Everything returned (attributes, method results) is wrapped again, so every item, folder
    or copy reached from the Outlook application is covered.
    """
    _BLOCKED = ("send", "submit", "forward", "reply", "delete", "harddelete", "remove", "clear")

    def __init__(self, obj):
        object.__setattr__(self, "_obj", obj)

    @staticmethod
    def _wrap(value):
        return OutlookGuard(value) if hasattr(value, "_oleobj_") else value

    def __getattr__(self, name):
        if name.lower().startswith(self._BLOCKED):
            raise PermissionError(f"Zakázáno ({name}): skript nic neodesílá ani nemaže, jen vytváří koncepty.")
        value = getattr(self._obj, name)
        if callable(value) and not hasattr(value, "_oleobj_"):
            def call(*args):
                return self._wrap(value(*(a._obj if isinstance(a, OutlookGuard) else a for a in args)))
            return call
        return self._wrap(value)

    def __setattr__(self, name, value):
        setattr(self._obj, name, value)


def _match_case(new, old):
    if old.isupper():
        return new.upper()
    if old[0].isupper():
        return new[0].upper() + new[1:]
    return new


def _word_before(text, pos):
    """The word right before pos, with tags and entities removed ("v", "za", "1.", ...)."""
    before = re.sub(r"<[^>]*>", " ", text[max(0, pos - 300):pos])
    before = html.unescape(before).replace("\xa0", " ").split()
    return before[-1].lower() if before else ""


def _case_form(word, text, pos):
    forms = FORMS[word.lower()]
    if len(forms) == 1:
        return next(iter(forms))
    prev = _word_before(text, pos)
    if prev in LOC_WORDS and 2 in forms:
        return 2
    if (prev in GEN_WORDS or re.fullmatch(r"\d{1,2}\.", prev)) and 1 in forms:
        return 1
    return min(forms)


def _decode_entities(text):
    """&#345; -> ř etc., so month words written as entities are found. Keeps markup-safe ones."""
    def repl(m):
        ch = html.unescape(m.group(0))
        return m.group(0) if len(ch) != 1 or ch in "<>&\"'\xa0" else ch
    return ENTITY_RE.sub(repl, text)


def change_month(text, month, year, is_html=False):
    """Every month word in text -> the same case form of `month` (1-12), the year after it -> `year`.

    Returns (new_text, changes) with changes = list of "ZÁŘÍ 2026 -> ŘÍJEN 2026".
    In HTML only visible text is changed: tags, attributes and links are left alone.
    """
    if is_html:
        text = _decode_entities(text)
        body_start = text.lower().find("<body")
    changes = []

    def repl(m):
        start = m.start()
        if is_html and (start < body_start or text.rfind("<", 0, start) > text.rfind(">", 0, start)):
            return m.group(0)                        # inside <head> or inside a tag
        old_word = m.group("month")
        new_word = _match_case(MONTHS[month - 1][_case_form(old_word, text, start)], old_word)
        new = new_word + (m.group("gap") + str(year) if m.group("year") else "")
        if new != m.group(0):
            old_label = re.sub(r"<[^>]*>|&nbsp;|&#160;|\xa0|\s+", " ", m.group(0)).split()
            new_label = re.sub(r"<[^>]*>|&nbsp;|&#160;|\xa0|\s+", " ", new).split()
            changes.append(f"{' '.join(old_label)} -> {' '.join(new_label)}")
        return new

    return MONTH_RE.sub(repl, text), changes


def _mailbox(ns, address):
    for i in range(1, ns.Folders.Count + 1):
        folder = ns.Folders.Item(i)
        if folder.Name.lower() == address.lower():
            return folder
    names = [ns.Folders.Item(i).Name for i in range(1, ns.Folders.Count + 1)]
    raise RuntimeError(f"Schránka {address} není v Outlooku. Vidím jen: {names}")


def _drafts(root):
    try:
        return root.Store.GetDefaultFolder(OL_FOLDER_DRAFTS)
    except Exception:
        return root.Folders.Item("Koncepty")


def _real_attachments(item):
    """Names of attachments that are not pictures (an old PDF left in a template).

    Pictures are skipped: the templates carry dozens of hidden inline images (logo, icons).
    """
    names = [item.Attachments.Item(i).FileName for i in range(1, item.Attachments.Count + 1)]
    return [n for n in names if not re.search(r"\.(png|gif|jpe?g|bmp|svg|emz|wmz)$", n, re.I)]


def strip_reply_prefix(subject):
    """'Re: Fw: REPORT – …' -> 'REPORT – …' (Re, Fw, Fwd and the Czech Odp, Př, any number of them)."""
    return re.sub(r"^\s*(?:(?:re|fw|fwd|odp|př|pr)\s*:\s*)+", "", subject, flags=re.I)


def _summary(changes):
    counts = {}
    for c in changes:
        counts[c] = counts.get(c, 0) + 1
    return ", ".join(f"{c} ({n}×)" if n > 1 else c for c, n in counts.items()) or "nic"


def vytvor_koncepty(mesic, rok, ostry_beh=False, jen=None):
    """Copy every template into analytik@ Drafts with the month set to mesic/rok. Never sends.

    mesic, rok: target month (int or "10"/"2026" as in my_dataT).
    ostry_beh=False: dry run, only prints what would change; nothing is created.
    jen: create only templates whose subject contains this text (first real test on one e-mail).
    """
    mesic, rok = int(mesic), int(rok)
    outlook = OutlookGuard(win32com.client.Dispatch("Outlook.Application"))
    ns = outlook.GetNamespace("MAPI")
    templates = _mailbox(ns, TEMPLATE_MAILBOX).Folders.Item(TEMPLATE_FOLDER)
    target = _mailbox(ns, TARGET_MAILBOX)
    drafts = _drafts(target)
    # the path says "\\Analytik VLMEDIA\Koncepty", so compare the mailbox itself, not its name
    if drafts.Store.StoreID != target.Store.StoreID:
        raise RuntimeError(f"Cílová složka {drafts.FolderPath} není ve schránce {TARGET_MAILBOX}.")

    existing = {drafts.Items.Item(i).Subject for i in range(1, drafts.Items.Count + 1)}
    items = [templates.Items.Item(i) for i in range(1, templates.Items.Count + 1)]
    items = sorted((it for it in items if it.Class == OL_MAIL), key=lambda it: it.Subject)
    mode = "OSTRÝ BĚH" if ostry_beh else "ZKOUŠKA (nic se nevytvoří)"
    print(f"{mode}: {len(items)} šablon z {templates.FolderPath} -> {drafts.FolderPath}, "
          f"měsíc {MONTHS[mesic - 1][0]} {rok}\n")

    created = 0
    for tpl in items:
        if jen and jen.lower() not in tpl.Subject.lower():
            continue
        subject, subj_changes = change_month(strip_reply_prefix(tpl.Subject), mesic, rok)
        body, body_changes = change_month(tpl.HTMLBody, mesic, rok, is_html=True)
        print(f"• {subject}")
        print(f"    předmět: {_summary(subj_changes)} | tělo: {_summary(body_changes)}")
        warnings = []
        if strip_reply_prefix(tpl.Subject) != tpl.Subject:
            print("    - Re:/Fw: z předmětu odstraněno")
        if not MONTH_RE.search(tpl.Subject):
            warnings.append("v předmětu jsem nenašel měsíc")
        no_body = not (tpl.Body or "").strip()
        if no_body:
            warnings.append("šablona nemá tělo: PŘESKOČENO")
        elif not MONTH_RE.search(_decode_entities(tpl.HTMLBody)):
            warnings.append("v těle jsem nenašel měsíc")
        if tpl.Sent:
            warnings.append("šablona je odeslaný/přijatý e-mail, ne koncept: PŘESKOČENO")
        attachments = _real_attachments(tpl)
        if attachments:
            warnings.append(f"šablona má přílohu {attachments} (zkopíruje se)")
        if subject in existing:
            warnings.append("koncept s tímto předmětem už v analytik@ je (nebo je šablona dvakrát): "
                            "PŘESKOČENO")
        for w in warnings:
            print(f"    ! {w}")
        if tpl.Sent or no_body or subject in existing:
            continue
        existing.add(subject)
        if not ostry_beh:
            continue

        copy = tpl.Copy()
        try:
            draft = copy.Move(drafts)
        except Exception as exc:
            raise RuntimeError(f"Kopii se nepodařilo přesunout do {drafts.FolderPath} ({exc}). "
                               f"Kopie zůstala ve složce {TEMPLATE_FOLDER}.") from exc
        draft.Subject = subject
        draft.HTMLBody = body
        draft.SentOnBehalfOfName = TARGET_MAILBOX
        draft.Save()
        created += 1
        problems = []
        if draft.Sent:
            problems.append("NENÍ koncept")
        if not (draft.Body or "").strip():
            problems.append("nemá tělo")
        if draft.SentOnBehalfOfName.lower() not in (TARGET_MAILBOX, "analytik vlmedia"):
            problems.append(f"odesílatel je {draft.SentOnBehalfOfName}")
        print(f"    ✓ koncept vytvořen, od: {draft.SentOnBehalfOfName}"
              + (f"  !!! ZKONTROLUJ: {', '.join(problems)}" if problems else ""))

    print(f"\nHotovo: {created} konceptů vytvořeno." if ostry_beh
          else "\nZkouška hotová. Když sedí, spusť s ostry_beh=True.")


if __name__ == "__main__":
    tests = [
        ("REPORT – Facebook Monetizace Nezpravodajské weby za září 2026", 10, 2026,
         "REPORT – Facebook Monetizace Nezpravodajské weby za říjen 2026"),
        ("za měsíc ZÁŘÍ 2026", 10, 2026, "za měsíc ŘÍJEN 2026"),
        ("za měsíc ŘÍJEN2026", 11, 2026, "za měsíc LISTOPAD2026"),
        ("v září, do září, 1. září, za září", 10, 2026, "v říjnu, do října, 1. října, za říjen"),
        ("v listopadu, za listopad 2026", 12, 2026, "v prosinci, za prosinec 2026"),
        ("Prosinec 2026", 1, 2027, "Leden 2027"),
        ("Zářijový report, záříčko", 10, 2026, "Zářijový report, záříčko"),
    ]
    for src, m, y, want in tests:
        got, _ = change_month(src, m, y)
        assert got == want, (src, got, want)
    for src in ("Re: REPORT – X", "RE:Fw: REPORT – X", "Odp: Př: REPORT – X", "REPORT – X"):
        assert strip_reply_prefix(src) == "REPORT – X", src
    assert strip_reply_prefix("Report – Reels") == "Report – Reels"
    h = ('<head><title>září</title></head><body><a href="x?zari=září">zde</a>'
         '<b>za měsíc Z&#193;Ř&#205;</b>&nbsp;<span>2026</span></body>')
    got, ch = change_month(h, 10, 2026, is_html=True)
    assert got == ('<head><title>září</title></head><body><a href="x?zari=září">zde</a>'
                   '<b>za měsíc ŘÍJEN</b>&nbsp;<span>2026</span></body>'), got
    assert ch == ["ZÁŘÍ 2026 -> ŘÍJEN 2026"], ch
    print("email_koncepty selftest: OK")
