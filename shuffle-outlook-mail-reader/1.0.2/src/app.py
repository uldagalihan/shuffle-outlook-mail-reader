#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import datetime as _dt
import re
import unicodedata
import requests
from walkoff_app_sdk.app_base import AppBase

GRAPH = "https://graph.microsoft.com/v1.0"

# ==== Username normalizasyonu için TR map ====
_TR_MAP = str.maketrans({
    "ç": "c", "ğ": "g", "ı": "i", "ö": "o", "ş": "s", "ü": "u",
    "Ç": "c", "Ğ": "g", "İ": "i", "I": "i", "Ö": "o", "Ş": "s", "Ü": "u",
})


def _normalize_person_key(s: str) -> str:
    """
    Person parametresine yazılan değeri (full name veya username)
    QRadar tarafındaki username formatına (alihan.uludag gibi) normalize eder.
    """
    s = (s or "").strip()
    if not s:
        return ""

    # Eğer içinde boşluk varsa full name gibi düşün → ilk ve son token
    if " " in s:
        toks = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü]+", s)
        if len(toks) >= 2:
            use = f"{toks[0]} {toks[-1]}"
        else:
            use = s
    else:
        # Zaten username gibi (alihan.uludag) kabul et
        use = s

    use = use.translate(_TR_MAP).lower()
    try:
        use = unicodedata.normalize("NFKD", use).encode("ascii", "ignore").decode("ascii")
    except Exception:
        pass

    use = re.sub(r"[^a-z0-9 .]+", "", use)
    use = re.sub(r"\s+", " ", use).strip()
    use = use.replace(" ", ".")
    use = re.sub(r"\.+", ".", use).strip(".")
    return use


def _parse_excluded_persons(persons_str: str):
    """
    Outlook action parametresinden gelen Person listesini set'e çevirir.
    Ör: "alihan.uludag, emre.uludag" -> {"alihan.uludag", "emre.uludag"}
    """
    if not persons_str:
        return set()

    parts = re.split(r"[,\n;]+", persons_str)
    out = set()
    for p in parts:
        key = _normalize_person_key(p)
        if key:
            out.add(key)
    return out


class OutlookGraphAppOnly(AppBase):
    __version__ = "1.1.1"   # versiyonu artırdım
    app_name = "Outlook Graph AppOnly"

    def __init__(self, redis=None, logger=None, **kwargs):
        super().__init__(redis=redis, logger=logger, **kwargs)

    # === Auth ===
    def _token(self, tenant_id, client_id, client_secret):
        url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        r = requests.post(url, data=data, timeout=30)
        r.raise_for_status()
        return r.json()["access_token"]

    # === HTTP GET helper (Prefer text body) ===
    def _get(self, url, tok, params=None, prefer_text_body=True):
        headers = {"Authorization": f"Bearer {tok}"}
        if prefer_text_body:
            headers["Prefer"] = 'outlook.body-content-type="text"'

        if self.logger:
            try:
                prepped = requests.Request("GET", url, params=params).prepare()
                self.logger.info(f"[Graph GET] url={prepped.url} params={json.dumps(params, ensure_ascii=False)}")
            except Exception:
                pass

        r = requests.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        return r.json()

    # === Mesajları subject = '...' ile getir (body dahil) ===
    def _fetch_by_exact_subject(self, tenant_id, client_id, client_secret, mailbox, subject, top=None):
        tok = self._token(tenant_id, client_id, client_secret)
        url = f"{GRAPH}/users/{mailbox}/messages"
        safe_subject = subject.replace("'", "''")

        filter_expr = f"receivedDateTime ge 1900-01-01T00:00:00Z and subject eq '{safe_subject}'"
        params = {
            "$select": "id,sender,subject,receivedDateTime,body,uniqueBody,bodyPreview",
            "$filter": filter_expr,
            "$orderby": "receivedDateTime desc",
        }
        if top is not None:
            try:
                t = max(1, min(1000, int(top)))
            except Exception:
                t = 10
            params["$top"] = t

        data = self._get(url, tok, params=params, prefer_text_body=True)
        return data.get("value", [])

    # === Body metnini al (uniqueBody varsa onu tercih et) ===
    @staticmethod
    def _get_body_text(item):
        body = ""
        if isinstance(item, dict):
            ub = item.get("uniqueBody") or {}
            b = item.get("body") or {}
            body = (ub.get("content") or b.get("content") or item.get("bodyPreview") or "")
        body = body.replace("\r\n", "\n").replace("\r", "\n")
        body = re.sub(r"[ \t]+", " ", body)
        body = re.sub(r"\n{2,}", "\n", body)
        return body.strip()

    # === İsim normalize ===
    @staticmethod
    def _clean_name(name):
        if not name:
            return ""
        name = re.sub(r"\s+", " ", name)
        return name.strip(" \t\r\n-–—.")

    # === Yardımcılar: token sınıfları ===
    @staticmethod
    def _is_all_caps_word(tok: str) -> bool:
        letters = re.sub(r"[^A-Za-zÇĞİÖŞÜçğıöşü]", "", tok)
        return len(letters) >= 2 and letters == letters.upper()

    @staticmethod
    def _is_title_like(tok: str) -> bool:
        return bool(re.match(r"^[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+$", tok))

    # === NEW HIRE: Ad Soyad yakalama ===
    @staticmethod
    def _extract_name_new_hire(text):
        m = re.search(r"CEP\s*TELEFONU\b.*?\b(\d{3,})\b(?P<rest>.*)", text, flags=re.IGNORECASE | re.DOTALL)
        if not m:
            m = re.search(r"S[İI]C[İI]L\s*NO\b.*?\b(\d{3,})\b(?P<rest>.*)", text, flags=re.IGNORECASE | re.DOTALL)
        rest = m.group("rest") if m else text

        tokens = re.findall(r"[A-Za-zÇĞİÖŞÜçğıöşü'’\-]+", rest)
        connectors = {"de","da","van","von","bin","ibn","al","el","oğlu","oglu","del","di"}
        name_tokens = []

        for tok in tokens:
            low = tok.lower()
            if OutlookGraphAppOnly._is_all_caps_word(tok):
                break
            if OutlookGraphAppOnly._is_title_like(tok) or low in connectors:
                name_tokens.append(tok)
                if len(name_tokens) >= 6:
                    break
                continue
            if name_tokens:
                break

        while name_tokens and name_tokens[-1].lower() in connectors:
            name_tokens.pop()

        name = OutlookGraphAppOnly._clean_name(" ".join(name_tokens))

        if not name:
            m2 = re.search(
                r"ADI\s*SOYADI\s*[:\-]?\s*(?P<name>(?:[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+(?:\s+[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+){1,5}))",
                text, flags=re.IGNORECASE
            )
            if m2:
                name = OutlookGraphAppOnly._clean_name(m2.group("name"))

        if len(name.split()) < 2:
            return ""
        return name

    # === Yardımcı: tarih yakala ===
    @staticmethod
    def _extract_first_date(text):
        if not text:
            return None
        t = text.replace("\r\n", "\n").replace("\r", "\n")
        t = re.sub(r"\s+", " ", t)

        ctx_pat = re.compile(r"tarihi\s+itibari\s+ile.{0,30}", re.IGNORECASE)
        date_pats = [
            r"(?P<d>\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b)",
            r"(?P<d>\b\d{4}-\d{1,2}-\d{1,2}\b)",
        ]

        mctx = ctx_pat.search(t)
        search_ranges = []
        if mctx:
            s, e = mctx.start(), mctx.end()
            search_ranges.append(t[max(0, s-40): min(len(t), e+40)])
        search_ranges.append(t)

        def _parse_candidate(s2):
            m = re.search(r"\b(\d{1,2})[./](\d{1,2})[./](\d{2,4})\b", s2)
            if m:
                d, M, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                if y < 100:
                    y += 2000
                try:
                    return _dt.date(y, M, d)
                except ValueError:
                    pass
            m = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", s2)
            if m:
                y, M, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
                try:
                    return _dt.date(y, M, d)
                except ValueError:
                    pass
            return None

        for seg in search_ranges:
            for p in date_pats:
                m = re.search(p, seg)
                if m:
                    got = _parse_candidate(m.group("d"))
                    if got:
                        return got
        return None

    # === TERMINATION Regex ===
    @staticmethod
    def _extract_name_termination(text):
        txt = (text or "").replace("\r\n", "\n").replace("\r", "\n")
        txt = re.sub(r"[ \t]+", " ", txt)
        txt = re.sub(r"\n+", " ", txt).strip()
        name_token = r"(?:[A-ZÇĞİÖŞÜ][A-Za-zÇĞİÖŞÜçğıöşü'’\-]+|[A-ZÇĞİÖŞÜ]{2,})"
        name_pattern = rf"(?P<name>{name_token}(?:\s+{name_token}){{1,5}})"
        pat_with_label = re.compile(
            rf"sicil\w*\s+ile\s+çalışan\s+{name_pattern}\s+isimli\s+çalışan\s+için\b",
            flags=re.IGNORECASE
        )
        pat_plain = re.compile(rf"\b{name_pattern}\s+için\b", flags=re.IGNORECASE)
        pat_without_label = re.compile(
            rf"sicil\w*\s+ile\s+çalışan\s+{name_pattern}\s+için\b", flags=re.IGNORECASE
        )
        for pat in (pat_with_label, pat_without_label, pat_plain):
            m = pat.search(txt)
            if m:
                raw = m.group("name").strip()
                raw = re.sub(r"\s+isimli\s+çalışan\b.*$", "", raw, flags=re.IGNORECASE).strip()
                connectors = {"de","da","van","von","bin","ibn","al","el","oğlu","oglu","del","di","di’"}
                toks = raw.split()
                while toks and toks[-1].lower() in connectors:
                    toks.pop()
                name = " ".join(toks)
                if len(name.split()) >= 2:
                    return OutlookGraphAppOnly._clean_name(name)
        return ""

    # === ISO yardımcıları ===
    @staticmethod
    def _to_iso_date(d: _dt.date) -> str:
        return d.isoformat() if d else ""

    @staticmethod
    def _to_iso_dt(dt: _dt.datetime) -> str:
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=_dt.timezone.utc)
        return dt.astimezone(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    @staticmethod
    def _midnight_utc_after_days(d: _dt.date, days: int) -> str:
        if not d:
            return ""
        target = _dt.datetime(d.year, d.month, d.day, tzinfo=_dt.timezone.utc) + _dt.timedelta(days=days)
        target = target.replace(hour=0, minute=0, second=0, microsecond=0)
        return OutlookGraphAppOnly._to_iso_dt(target)

    # === Zaman yardımcıları ===
    @staticmethod
    def _now_utc():
        return _dt.datetime.now(_dt.timezone.utc)

    @staticmethod
    def _parse_iso_utc(s: str):
        try:
            if s.endswith("Z"):
                s = s[:-1] + "+00:00"
            return _dt.datetime.fromisoformat(s)
        except Exception:
            return None

    @classmethod
    def _filter_ready(cls, items):
        """
        activate_at yoksa: hemen hazır sayar.
        activate_at varsa: activate_at <= now (UTC) olanları bırakır.
        """
        out = []
        now = cls._now_utc()
        for it in (items or []):
            if not isinstance(it, dict):
                continue
            act = (it.get("activate_at") or "").strip()
            if not act:
                out.append(it); continue
            dt = cls._parse_iso_utc(act)
            if not dt or dt <= now:
                out.append(it)
        return out

    # === ACTION 1: New Hire -> İSİM LİSTESİ ===
    def list_new_hire_messages(self, tenant_id, client_id, client_secret, mailbox, top=None):
        subject = "[Kurum Dışı] Şirkete Yeni Katılım - New Comer"
        items = self._fetch_by_exact_subject(tenant_id, client_id, client_secret, mailbox, subject, top)
        names = []
        for it in items:
            body = self._get_body_text(it)
            name = self._extract_name_new_hire(body)
            if name:
                names.append(name)
                if self.logger:
                    self.logger.info(f"[NEW_HIRE] matched name: {name}")
            else:
                if self.logger:
                    prev = (it.get("bodyPreview") or "")[:120]
                    self.logger.info(f"[NEW_HIRE] no match. preview={prev}")
        return {"success": True, "names": names}

    # === ACTION 2: Termination -> Detaylı JSON (+ activate filtre + exclude_persons) ===
    def list_termination_messages(
        self,
        tenant_id,
        client_id,
        client_secret,
        mailbox,
        top=None,
        only_ready=True,
        exclude_persons=None   # <<< YENİ PARAM
    ):
        """
        exclude_persons:
          Örnek input: "alihan.uludag, emre.uludag"
          Açıklama (Shuffle description): "İşe devam ettiği için mailden ismi çekilmeyecek kişileri buraya yazın."
        """
        subject = "[Kurum Dışı] Çalışan İlişik Kesme Bildirimi"
        items = self._fetch_by_exact_subject(tenant_id, client_id, client_secret, mailbox, subject, top)
        out_items = []
        names = []

        excluded = _parse_excluded_persons(exclude_persons)

        for it in items:
            body = self._get_body_text(it)
            name = self._extract_name_termination(body)
            term_date = self._extract_first_date(body)  # datetime.date veya None

            # Eğer isim var ve excluded listesindeyse, tamamen atla
            if name:
                uname = _normalize_person_key(name)
                if uname and uname in excluded:
                    if self.logger:
                        self.logger.info(f"[TERMINATION] excluded person: name={name} uname={uname}")
                    continue

            received_iso = it.get("receivedDateTime", "") or ""
            activate_at = self._midnight_utc_after_days(term_date, 3) if term_date else ""

            rec = {
                "name": name or "",
                "mail_received_at": received_iso,
                "termination_date": self._to_iso_date(term_date) if term_date else "",
                "activate_at": activate_at
            }
            out_items.append(rec)

        # Activate filtresi
        if only_ready:
            ready = self._filter_ready(out_items)
        else:
            ready = out_items

        # names: boş olmayan + uniq
        seen = set()
        for r in ready:
            nm = (r.get("name") or "").strip()
            if nm and nm not in seen:
                seen.add(nm)
                names.append(nm)

        if self.logger:
            self.logger.info(f"[TERMINATION] total={len(out_items)} ready={len(ready)} names={len(names)}")

        return {"success": True, "items": ready, "names": names}

    # === ACTION 3: Sadece 'hazır' isim listesi (QRadar'a direkt ver) ===
    def list_ready_termination_names(
        self,
        tenant_id,
        client_id,
        client_secret,
        mailbox,
        top=None,
        exclude_persons=None   # <<< aynı param buraya da
    ):
        res = self.list_termination_messages(
            tenant_id,
            client_id,
            client_secret,
            mailbox,
            top=top,
            only_ready=True,
            exclude_persons=exclude_persons
        )
        return {"success": res.get("success", True), "names": res.get("names", [])}


if __name__ == "__main__":
    OutlookGraphAppOnly.run()
