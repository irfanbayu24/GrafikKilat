"""
matchlib.py — pencocokan fleksibel untuk Cek Absensi (tanpa Streamlit, sehingga mudah diuji).

Kasus: File 1 punya kolom UNIT dan NAMA terpisah, File 2 (daftar hadir) hanya punya SATU kolom teks bebas
seperti "BKT - Budi", "Budi S (Biro Keuangan)", atau "dr. Budi iPhone", dan nama boleh tidak lengkap.

Cara kerja (per kata, bukan per kalimat):
  1. Teks dipecah menjadi kata (tanpa huruf besar/aksen/tanda baca); sapaan, gelar, dan label perangkat dibuang.
  2. UNIT dikenali bila kata unit (atau singkatannya, mis. "Pusat Data dan Informasi" -> "pdi") ada di teks.
     Kata unit dibuang dari teks; sisanya dianggap NAMA. Urutan "Unit - Nama" maupun "Nama - Unit" sama saja.
  3. NAMA cocok bila tiap kata dari File 2 ada di nama lengkap File 1: persis, awalan ("Bud" ~ "Budi"),
     inisial ("S" ~ "Santoso"), atau salah ketik 1 huruf (kata ≥ 5 huruf).
  4. Kandidat terbaik dipilih (unit lebih lengkap > nama persis > tanpa kata asing > bobot kata). Bila beberapa
     kandidat seimbang, hasilnya AMBIGU dan tidak ditebak.

Status File 1: Hadir · Hadir (belum pasti) · Belum hadir · Nama kosong.
Status entri File 2: Cocok · Ambigu · Tidak cocok · Kosong. Tiap kecocokan diberi keyakinan Tinggi/Sedang/Rendah.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

import pandas as pd

from checklib import key_text

DEFAULT_NOISE = (
    "pak bu bapak ibu mas mbak kak kakak bang dr drs dra ir prof hj kh ustadz ustadzah "          # sapaan & gelar
    "si kom kes sh mh se st mt msi mm spd mpd skm mkes apt amd skom mkom ssos "
    "iphone ipad android samsung redmi oppo vivo xiaomi realme host cohost guest tamu zoom user"   # label rapat/perangkat
)
_STOP = frozenset({"dan", "di", "of", "the", "ke", "untuk", "pada", "dengan", "yang", "dalam"})
_TOKEN = re.compile(r"[^\W_]+", re.UNICODE)

CONF_ORDER = {"Tinggi": 3, "Sedang": 2, "Rendah": 1, "": 0}


# ---------------------------------------------------------------------------
# Tokenisasi
# ---------------------------------------------------------------------------
def fold(text) -> str:
    """Huruf kecil + buang aksen (é -> e)."""
    t = unicodedata.normalize("NFKD", str(text))
    return "".join(ch for ch in t if not unicodedata.combining(ch)).casefold()


def parse_noise(text: str) -> frozenset:
    return frozenset(_TOKEN.findall(fold(text)))


def tokenize(text, noise: frozenset = frozenset()) -> list:
    return [t for t in _TOKEN.findall(fold(text)) if t not in noise]


def _acronym(tokens: list) -> str:
    core = [t for t in tokens if t not in _STOP]
    return "".join(t[0] for t in core) if len(core) >= 2 else ""


# ---------------------------------------------------------------------------
# Pencocokan kata
# ---------------------------------------------------------------------------
def _within_one_edit(a: str, b: str) -> bool:
    """Jarak edit ≤ 1: satu huruf beda/tambah/kurang, atau dua huruf bersebelahan tertukar."""
    if a == b:
        return True
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    if la == lb:
        diff = [i for i in range(la) if a[i] != b[i]]
        return len(diff) == 1 or (len(diff) == 2 and diff[1] == diff[0] + 1
                                  and a[diff[0]] == b[diff[1]] and a[diff[1]] == b[diff[0]])
    if la > lb:
        a, b = b, a
    i = 0
    while i < len(a) and a[i] == b[i]:
        i += 1
    return a[i:] == b[i + 1:]


def _token_weight(r: str, n: str, typo: bool) -> float:
    """Seberapa cocok kata File 2 (r) terhadap kata nama lengkap File 1 (n)."""
    if r == n:
        return 1.0
    if len(r) == 1:
        return 0.5 if n.startswith(r) else 0.0           # inisial
    if n.startswith(r):
        return 0.8 if len(r) >= 3 else 0.6               # awalan
    if typo and len(r) >= 5 and len(n) >= 5 and _within_one_edit(r, n):
        return 0.7                                        # salah ketik
    return 0.0


def _match_tokens(R: list, N: tuple, typo: bool):
    """Cocokkan kata File 2 (R) ke kata nama File 1 (N). None bila tidak cukup kuat (mis. hanya inisial)."""
    used = [False] * len(N)
    weights, extras = [], 0
    for r in sorted(R, key=len, reverse=True):            # kata panjang dulu agar tidak berebut dengan inisial
        best_j, best_w = -1, 0.0
        for j, n in enumerate(N):
            if not used[j]:
                w = _token_weight(r, n, typo)
                if w > best_w:
                    best_j, best_w = j, w
        if best_j < 0:
            extras += 1
        else:
            used[best_j] = True
            weights.append(best_w)
    if not weights or max(weights) < 0.6:
        return None
    return weights, extras


# ---------------------------------------------------------------------------
# Matcher
# ---------------------------------------------------------------------------
@dataclass
class Outcome:
    status: str                      # Cocok | Ambigu | Tidak cocok | Kosong
    persons: list = field(default_factory=list)       # indeks Person terbaik
    confidence: str = ""
    note: str = ""
    tokens: tuple = ()
    ranked: list = field(default_factory=list)        # [(indeks Person, kunci skor)] untuk penguji


@dataclass
class Person:
    row: int                         # posisi baris di df1
    unit_key: str
    name_tokens: tuple


class Matcher:
    def __init__(self, df1: pd.DataFrame, unit_col, name_col: str, *, noise: frozenset,
                 unit_overlap: float = 1.0, name_only_fallback: bool = False, max_extra: int = 0, typo: bool = True):
        self.noise, self.unit_overlap = noise, unit_overlap
        self.fallback, self.max_extra, self.typo = name_only_fallback, max_extra, typo
        self.use_unit = bool(unit_col)
        self.persons: list = []
        self.blank_rows: list = []
        self.units: dict = {}                 # unit_key -> (kata inti, semua kata, singkatan)
        self.by_unit_prefix: dict = {}        # (unit_key, 2 huruf awal) -> [indeks Person]
        self.by_prefix: dict = {}             # 2 huruf awal -> [indeks Person]
        names = df1[name_col].tolist()
        units = df1[unit_col].tolist() if unit_col else [None] * len(df1)
        for row, (u, n) in enumerate(zip(units, names)):
            nt = tuple(tokenize(key_text(n), noise))
            if not nt:
                self.blank_rows.append(row)
                continue
            ut = tokenize(key_text(u), noise) if self.use_unit else []
            uk = " ".join(ut)
            if uk not in self.units:
                self.units[uk] = ([t for t in ut if t not in _STOP] or ut, frozenset(ut), _acronym(ut))
            pi = len(self.persons)
            self.persons.append(Person(row, uk, nt))
            for h in {t[:2] for t in nt if len(t) >= 2}:
                self.by_unit_prefix.setdefault((uk, h), []).append(pi)
                self.by_prefix.setdefault(h, []).append(pi)

    # -- unit ---------------------------------------------------------------------------------
    def _unit_pools(self, t2: list, t2set: set) -> list:
        """[(unit_key|None, kata nama tersisa, pecahan unit cocok, unit_terverifikasi, longgar)]

        'longgar' = jalur fallback (unit tak terbaca): kata asing (mis. singkatan unit) boleh ada, tetapi seluruh
        nama lengkap di File 1 wajib ikut tercocokkan sebagai pengganti bukti unit yang hilang.
        """
        if not self.use_unit:
            return [(None, t2, 1.0, None, False)]
        pools, verified_any = [], False
        for uk, (core, allt, acr) in self.units.items():
            if not core:                                   # unit File 1 kosong: tidak bisa diverifikasi
                pools.append((uk, t2, 0.0, False, False))
                continue
            hits = [t for t in core if t in t2set]
            frac = len(hits) / len(core)
            removed = set(hits)
            if frac < self.unit_overlap or not hits:
                if acr and acr in t2set:                   # singkatan unit, mis. "pdi"
                    frac, removed = 1.0, {acr}
                else:
                    continue
            removed |= {t for t in allt if t in t2set and t in _STOP}
            pools.append((uk, [t for t in t2 if t not in removed], frac, True, False))
            verified_any = True
        if not verified_any and self.fallback:
            pools = [(None, t2, 0.0, False, True)]
        return pools

    # -- satu teks ----------------------------------------------------------------------------
    def match_text(self, text) -> Outcome:
        t2 = tokenize(key_text(text), self.noise)
        if not t2:
            return Outcome("Kosong")
        scored = []
        for uk, R, frac, verified, loose in self._unit_pools(t2, set(t2)):
            heads = {r[:2] for r in R if len(r) >= 2}
            cand: set = set()
            for h in heads:
                cand.update(self.by_prefix.get(h, ()) if uk is None else self.by_unit_prefix.get((uk, h), ()))
            for pi in cand:
                res = _match_tokens(R, self.persons[pi].name_tokens, self.typo)
                if res is None:
                    continue
                weights, extras = res
                if loose:
                    if len(weights) < len(self.persons[pi].name_tokens):
                        continue                      # fallback: nama lengkap harus tercocokkan semua
                elif extras > self.max_extra:
                    continue
                exact = sorted(R) == sorted(self.persons[pi].name_tokens)
                scored.append(((round(frac, 3), exact, -extras, round(sum(weights), 3)), pi, verified))
        if not scored:
            return Outcome("Tidak cocok", tokens=tuple(t2))
        scored.sort(key=lambda s: s[0], reverse=True)
        top_key = scored[0][0]
        top = [s for s in scored if s[0] == top_key]
        persons = sorted({s[1] for s in top})
        frac, exact, neg_extras, _ = top_key
        verified = top[0][2]
        ranked = [(s[1], s[0]) for s in scored[:5]]
        if len(persons) > 1:
            return Outcome("Ambigu", persons, "Rendah", "beberapa nama sama-sama cocok", tuple(t2), ranked)
        if verified is True and exact and neg_extras == 0:
            conf = "Tinggi"
        elif neg_extras == 0 and verified is not False:
            conf = "Sedang"
        else:
            conf = "Rendah"
        notes = []
        if verified is False:
            notes.append("unit tidak terbaca di File 2 (dicocokkan lewat nama saja)")
        if neg_extras:
            notes.append("ada kata yang tidak dikenali")
        return Outcome("Cocok", persons, conf, "; ".join(notes), tuple(t2), ranked)


# ---------------------------------------------------------------------------
# Jalankan untuk seluruh data + susun hasil
# ---------------------------------------------------------------------------
@dataclass
class MatchResult:
    people: pd.DataFrame        # semua baris File 1 + status
    hadir: pd.DataFrame
    belum_pasti: pd.DataFrame
    belum_hadir: pd.DataFrame
    blank_1: pd.DataFrame
    detail: pd.DataFrame        # satu baris per entri File 2
    only_in_2: pd.DataFrame     # baris asli File 2 yang tidak cocok
    summary: dict = field(default_factory=dict)


def _free(existing, name: str) -> str:
    cand, n = name, 2
    while cand in set(existing):
        cand, n = f"{name} ({n})", n + 1
    return cand


def run(matcher: Matcher, df1: pd.DataFrame, df2: pd.DataFrame, text_col: str, unit_col, name_col: str) -> MatchResult:
    texts = df2[text_col].tolist()
    outcomes = [matcher.match_text(t) for t in texts]
    unit_vals = df1[unit_col].tolist() if unit_col else [None] * len(df1)
    name_vals = df1[name_col].tolist()

    def label(pi: int) -> tuple:
        row = matcher.persons[pi].row
        return key_text(unit_vals[row]), key_text(name_vals[row])

    unique: dict = {}      # posisi baris df1 -> [(teks, keyakinan)]
    maybe: dict = {}
    drows = []
    for i, (text, oc) in enumerate(zip(texts, outcomes)):
        unit, nama, cand = "", "", ""
        if oc.status == "Cocok":
            row = matcher.persons[oc.persons[0]].row
            unique.setdefault(row, []).append((key_text(text), oc.confidence))
            unit, nama = label(oc.persons[0])
        elif oc.status == "Ambigu":
            for pi in oc.persons:
                maybe.setdefault(matcher.persons[pi].row, []).append(key_text(text))
            cand = "; ".join(f"{n} ({u})" if u else n for u, n in (label(pi) for pi in oc.persons[:6]))
        drows.append({"Baris di File 2": i + 1, "Teks di File 2": key_text(text), "Status": oc.status,
                      "Keyakinan": oc.confidence, "Unit (File 1)": unit, "Nama (File 1)": nama,
                      "Kandidat (jika ambigu)": cand, "Catatan": oc.note})
    detail = pd.DataFrame(drows, columns=["Baris di File 2", "Teks di File 2", "Status", "Keyakinan", "Unit (File 1)",
                                          "Nama (File 1)", "Kandidat (jika ambigu)", "Catatan"])

    blank = set(matcher.blank_rows)
    status, count, seen, conf = [], [], [], []
    for r in range(len(df1)):
        if r in blank:
            status.append("Nama kosong"); count.append(0); seen.append(""); conf.append("")
        elif r in unique:
            status.append("Hadir"); count.append(len(unique[r]))
            seen.append("; ".join(t for t, _ in unique[r][:3]))
            conf.append(max((c for _, c in unique[r]), key=lambda c: CONF_ORDER[c]))
        elif r in maybe:
            status.append("Hadir (belum pasti)"); count.append(0)
            seen.append("; ".join(maybe[r][:3]) + " (ambigu)"); conf.append("Rendah")
        else:
            status.append("Belum hadir"); count.append(0); seen.append(""); conf.append("")
    people = df1.copy()
    people[_free(df1.columns, "Status kehadiran")] = status
    people[_free(df1.columns, "Jumlah entri cocok di File 2")] = count
    people[_free(df1.columns, "Teks di File 2")] = seen
    people[_free(df1.columns, "Keyakinan")] = conf
    status_col = people.columns[len(df1.columns)]
    by = lambda s: people[people[status_col] == s].reset_index(drop=True)   # noqa: E731

    only_mask = [oc.status == "Tidak cocok" for oc in outcomes]
    n_valid = len(df1) - len(blank)
    st = pd.Series(status)
    hadir_n = int((st == "Hadir").sum())
    summary = {
        "total_1": len(df1), "valid_1": n_valid, "hadir": hadir_n,
        "belum_pasti": int((st == "Hadir (belum pasti)").sum()), "belum_hadir": int((st == "Belum hadir").sum()),
        "blank_1": len(blank), "total_2": len(df2),
        "cocok": sum(o.status == "Cocok" for o in outcomes), "ambigu": sum(o.status == "Ambigu" for o in outcomes),
        "tidak_cocok": sum(only_mask), "kosong": sum(o.status == "Kosong" for o in outcomes),
        "pct_hadir": round(hadir_n / n_valid * 100, 1) if n_valid else 0.0,
    }
    return MatchResult(people, by("Hadir"), by("Hadir (belum pasti)"), by("Belum hadir"), by("Nama kosong"),
                       detail, df2.loc[only_mask].reset_index(drop=True), summary)


def build_report(res: MatchResult) -> dict:
    s = res.summary
    rows = [
        ("Total baris File 1", s["total_1"]), ("Baris File 1 tanpa nama (tidak dicocokkan)", s["blank_1"]),
        ("Hadir", s["hadir"]), ("Hadir (belum pasti, ambigu)", s["belum_pasti"]), ("Belum hadir", s["belum_hadir"]),
        ("Persentase hadir (pasti, dari baris yang punya nama, %)", s["pct_hadir"]),
        ("Total entri File 2", s["total_2"]), ("Entri cocok", s["cocok"]), ("Entri ambigu", s["ambigu"]),
        ("Entri tidak cocok (hanya di File 2)", s["tidak_cocok"]), ("Entri kosong", s["kosong"]),
    ]
    sheets = {
        "Ringkasan": pd.DataFrame(rows, columns=["Keterangan", "Jumlah"]),
        "Hadir": res.hadir, "Hadir (belum pasti)": res.belum_pasti, "Belum hadir": res.belum_hadir,
        "Tidak cocok (File 2)": res.only_in_2, "Rincian pencocokan": res.detail,
    }
    if len(res.blank_1):
        sheets["Nama kosong File 1"] = res.blank_1
    return sheets
