from __future__ import annotations

import re
import shutil
import subprocess
from html import unescape
from pathlib import Path


DEFAULT_SQL_DATA_DIR = r"C:\Program Files\Microsoft SQL Server\MSSQL15.SQLEXPRESS\MSSQL\DATA"
SYSTEM_DATABASE_FILES = {"master", "model", "msdb", "tempdb"}
_NON_PRODUCT_KEYWORDS = (
    "rabatt",
    "vorauszahlungsrabatt",
    "skonto",
    "nachlass",
    "fracht",
    "lieferkosten",
    "porto",
    "verpackung",
    "montagekosten",
    "zwischensumme",
    "zwischentotal",
    "gesamtsumme",
    "stornierung",
    "unkosten",
    "abholung",
    "lieferung erfolgt",
    "lieferschein",
    "auftragsnummer",
    "projektnummer",
    "text auf blende",
    "pro verankerung",
    "expresslieferung",
    "warenverteilung",
    "beachten sie",
    "bitte beachten",
    "achtung",
    "hinweis",
    "zur info",
    "zur information",
)
_NON_PRODUCT_PREFIXES = (
    "zwi",
    "gesa",
    "storn",
    "raba",
    "skon",
)
_HTML_TAG_RE = re.compile(r"<[^>]+>")
_MULTISPACE_RE = re.compile(r"\s+")
_QTY_PREFIX_RE = re.compile(
    r"^\s*(\d+)\s*(?:x|stk\.?|st(?:ue|ü)ck)?\s*[:\-]?\s*(.+)$",
    flags=re.IGNORECASE,
)
_WEIGHT_SUFFIX_RE = re.compile(r"\(\s*[\d.,]+\s*kg\s*\)\s*$", flags=re.IGNORECASE)
_WEIGHT_SUFFIX_CAPTURE_RE = re.compile(r"\(\s*([\d.,]+)\s*kg\s*\)\s*$", flags=re.IGNORECASE)
_PRODUCT_SPLIT_RE = re.compile(r";\s*(?=\d+\s*(?:x|stk\.?|stueck)\b)", flags=re.IGNORECASE)


def _decode_sqlcmd_output(raw: bytes | None) -> str:
    payload = raw or b""
    if not payload:
        return ""
    for encoding in ("utf-8", "cp850", "cp1252"):
        try:
            return payload.decode(encoding)
        except UnicodeDecodeError:
            continue
    return payload.decode("latin-1", errors="replace")


def _normalize_product_title(text: str) -> str:
    value = unescape(str(text or ""))
    value = _HTML_TAG_RE.sub(" ", value)
    value = _MULTISPACE_RE.sub(" ", value).strip(" -;,.")
    if value.endswith("("):
        value = value[:-1].strip()
    return value


def _is_non_product_line(title: str) -> bool:
    value = str(title or "").casefold()
    if not value:
        return True
    if value in {"produkt", "-", "-- neue seite --"}:
        return True
    if value.startswith("produkt "):
        return True
    if any(keyword in value for keyword in _NON_PRODUCT_KEYWORDS):
        return True
    if any(value.startswith(prefix) for prefix in _NON_PRODUCT_PREFIXES):
        return True
    if len(value) < 3:
        return True
    return not any(ch.isalpha() for ch in value)


def _split_quantity_and_title(text: str) -> tuple[int, str]:
    value = _normalize_product_title(text)
    if not value:
        return 0, ""
    match = _QTY_PREFIX_RE.match(value)
    if match:
        qty = int(match.group(1) or 1)
        title = _normalize_product_title(match.group(2))
        return max(1, qty), title
    return 1, value


def _extract_product_titles(raw_products: str) -> str:
    raw = str(raw_products or "").strip()
    if not raw:
        return ""
    parts = [segment.strip() for segment in _PRODUCT_SPLIT_RE.split(raw)]
    cleaned_items = []
    for part in parts:
        text = _normalize_product_title(part)
        if not text:
            continue
        weight_match = _WEIGHT_SUFFIX_CAPTURE_RE.search(text)
        weight_text = ""
        if weight_match:
            weight_text = str(weight_match.group(1) or "").replace(",", ".").strip()
            text = _WEIGHT_SUFFIX_RE.sub("", text).strip()
        qty, title = _split_quantity_and_title(text)
        # Avoid duplicate quantities like "7x 2x Universal-..."
        _nested_qty, nested_title = _split_quantity_and_title(title)
        if nested_title:
            title = nested_title
        if _is_non_product_line(title):
            continue
        item = f"{max(1, int(qty or 1))}x {title}"
        if weight_text:
            item = f"{item} ({weight_text} kg)"
        cleaned_items.append(item)
    if cleaned_items:
        return "; ".join(cleaned_items)
    return ""


def infer_database_name_from_data_dir(data_dir: str) -> str:
    path = Path(str(data_dir or "").strip() or DEFAULT_SQL_DATA_DIR)
    if not path.exists() or not path.is_dir():
        return ""

    candidates = []
    for file in path.glob("*.mdf"):
        stem = file.stem
        if stem.lower() in SYSTEM_DATABASE_FILES:
            continue
        try:
            size = file.stat().st_size
        except OSError:
            size = 0
        candidates.append((size, stem))

    if not candidates:
        return ""
    candidates.sort(reverse=True)
    return candidates[0][1]


def fetch_order_rows(
    *,
    server_instance: str,
    database: str,
    limit: int = 10000,
) -> list[dict]:
    sqlcmd = shutil.which("sqlcmd")
    if not sqlcmd:
        raise RuntimeError("sqlcmd wurde nicht gefunden. Bitte SQL Server Command Line Tools installieren.")

    server = str(server_instance or r".\SQLEXPRESS").strip() or r".\SQLEXPRESS"
    db = str(database or "").strip()
    if not db:
        raise RuntimeError("Kein SQL-Datenbankname konfiguriert.")

    row_limit = max(1, min(int(limit or 5000), 50000))
    metadata_query = """
SET NOCOUNT ON;
SELECT c.name
FROM sys.columns c
INNER JOIN sys.objects o ON o.object_id = c.object_id
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
WHERE s.name = 'dbo'
  AND o.name = 'WW_Pos'
ORDER BY c.column_id;
""".strip()
    metadata_proc = subprocess.run(
        [sqlcmd, "-S", server, "-d", db, "-f", "65001", "-w", "65535", "-s", "|", "-h", "-1", "-Q", metadata_query],
        capture_output=True,
        text=False,
        timeout=60,
        check=False,
    )
    metadata_stdout = _decode_sqlcmd_output(metadata_proc.stdout)
    pos_columns = {line.strip() for line in metadata_stdout.splitlines() if line.strip()}

    quantity_col = next((name for name in ("Menge", "Anzahl", "Qty") if name in pos_columns), None)
    weight_col = next(
        (name for name in ("Gewicht", "Gewichte", "GewichtKG", "GewichtKg", "Bruttogewicht", "Nettogewicht") if name in pos_columns),
        None,
    )
    title_columns = [
        name
        for name in (
            "Bezeichnung",
            "Artikelbezeichnung",
            "ArtikelBez",
            "ArtikelText",
            "Kurztext",
            "Titel",
            "Name",
        )
        if name in pos_columns
    ]
    if not title_columns and "Bezeichnung" in pos_columns:
        title_columns = ["Bezeichnung"]
    position_type_col = next(
        (name for name in ("Positionsart", "PosTyp", "Positionstyp", "Typ", "Art", "Zeilentyp") if name in pos_columns),
        None,
    )
    text_flag_col = next(
        (name for name in ("IstText", "Textzeile", "IsText", "NurText", "Kommentar", "IsComment") if name in pos_columns),
        None,
    )
    qty_expr = f"TRY_CONVERT(decimal(18, 3), p2.[{quantity_col}])" if quantity_col else "CAST(1 AS decimal(18, 3))"
    qty_norm_expr = f"CASE WHEN {qty_expr} IS NULL OR {qty_expr} <= 0 THEN CAST(1 AS decimal(18, 3)) ELSE {qty_expr} END"
    if title_columns:
        title_nullif = ", ".join(f"NULLIF(p2.[{col}], '')" for col in title_columns)
        title_expr = f"LTRIM(RTRIM(COALESCE({title_nullif}, '')))"
    else:
        title_expr = "LTRIM(RTRIM(COALESCE(NULLIF(p2.Bezeichnung, ''), '')))"
    keytext_expr = f"LOWER(CONCAT({title_expr}, ' ', COALESCE(p2.Zusatztext, '')))"
    text_line_conditions = [
        f"{title_expr} = ''",
        f"LOWER({title_expr}) IN ('produkt', 'text', 'hinweis', 'kommentar', 'bemerkung')",
        f"{keytext_expr} LIKE '%rabatt%'",
        f"{keytext_expr} LIKE '%vorauszahlungsrabatt%'",
        f"{keytext_expr} LIKE '%skonto%'",
        f"{keytext_expr} LIKE '%nachlass%'",
        f"{keytext_expr} LIKE '%zwischensumme%'",
        f"{keytext_expr} LIKE '%zwischentotal%'",
        f"{keytext_expr} LIKE '%gesamtsumme%'",
        f"{keytext_expr} LIKE '%gesamttotal%'",
        f"{keytext_expr} LIKE '%beachten sie%'",
        f"{keytext_expr} LIKE '%bitte beachten%'",
        f"{keytext_expr} LIKE '%hinweis%'",
        f"{keytext_expr} LIKE '%zur info%'",
        f"{keytext_expr} LIKE '%zur information%'",
    ]
    if position_type_col:
        text_line_conditions.append(
            f"LOWER(CONVERT(nvarchar(100), COALESCE(p2.[{position_type_col}], ''))) IN ('text', 'txt', 'hinweis', 'kommentar', 'bemerkung', 'note', 'notes')"
        )
    if text_flag_col:
        text_line_conditions.append(f"TRY_CONVERT(int, COALESCE(p2.[{text_flag_col}], 0)) = 1")
    include_product_line_expr = f"NOT ({' OR '.join(text_line_conditions)})"
    if weight_col:
        unit_weight_expr = f"TRY_CONVERT(decimal(18, 3), p2.[{weight_col}])"
        line_weight_expr = f"({qty_norm_expr} * COALESCE({unit_weight_expr}, CAST(0 AS decimal(18, 3))))"
        weight_suffix_expr = (
            f"CASE WHEN {unit_weight_expr} IS NOT NULL "
            f"THEN CONCAT(' (', CONVERT(nvarchar(32), CAST(ROUND({line_weight_expr}, 2) AS decimal(18, 2))), ' kg)') "
            f"ELSE '' END"
        )
        total_weight_expr = f"CONVERT(nvarchar(32), CAST(ROUND(SUM({line_weight_expr}), 2) AS decimal(18, 2)))"
    else:
        weight_suffix_expr = "''"
        total_weight_expr = "''"
    query = f"""
SET NOCOUNT ON;
SELECT TOP {row_limit}
  CONVERT(nvarchar(50), k.Ident) AS ImportID,
  REPLACE(REPLACE(REPLACE(COALESCE(CONVERT(nvarchar(50), k.Kopf), ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Auftragsnummer,
  COALESCE(CONVERT(varchar(10), k.Datum, 23), '') AS Bestelldatum,
  REPLACE(REPLACE(REPLACE(
      LTRIM(RTRIM(COALESCE(NULLIF(delivery_addr.Firma, ''), NULLIF(CONCAT(COALESCE(delivery_addr.Vorname, ''), ' ', COALESCE(delivery_addr.Nachname, '')), ' '), 'Unbekannt')))
      , CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Name,
  REPLACE(REPLACE(REPLACE(COALESCE(delivery_addr.Strasse, ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Strasse,
  REPLACE(REPLACE(REPLACE(COALESCE(delivery_addr.PLZ, ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS PLZ,
  REPLACE(REPLACE(REPLACE(COALESCE(delivery_addr.Ort, ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Ort,
  REPLACE(REPLACE(REPLACE(COALESCE(delivery_addr.Land, ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Land,
  REPLACE(REPLACE(REPLACE(
      LTRIM(RTRIM(COALESCE(NULLIF(order_addr.Firma, ''), NULLIF(CONCAT(COALESCE(order_addr.Vorname, ''), ' ', COALESCE(order_addr.Nachname, '')), ' '), '')))
      , CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS AuftragName,
  REPLACE(REPLACE(REPLACE(COALESCE(order_addr.PLZ, ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS AuftragPLZ,
  REPLACE(REPLACE(REPLACE(COALESCE(order_addr.Ort, ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS AuftragOrt,
  REPLACE(REPLACE(REPLACE(COALESCE(mail.Kontakt, 'N/A'), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Email,
  REPLACE(REPLACE(REPLACE(COALESCE(phone.Kontakt, 'N/A'), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Telefon,
  REPLACE(REPLACE(REPLACE(COALESCE(k.Gewichte, 'N/A'), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Gewicht,
  REPLACE(REPLACE(REPLACE(COALESCE(pos.Produkte, ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Produkte,
  REPLACE(REPLACE(REPLACE(COALESCE(pos.ProduktgewichtTotal, ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS ProduktgewichtTotal,
  REPLACE(REPLACE(REPLACE(COALESCE(k.Notiz, ''), CHAR(13), ' '), CHAR(10), ' '), '|', '/') AS Notizen,
  COALESCE(delivery.Liefercode, '') AS Liefercode
FROM dbo.WW_Kopf k
LEFT JOIN dbo.AVE_Stamm order_addr ON order_addr.Ident = k.AdressID
LEFT JOIN dbo.AVE_Stamm delivery_addr ON delivery_addr.Ident = COALESCE(k.LieferadressID, k.AdressID)
OUTER APPLY (
    SELECT
        STRING_AGG(
            CAST(CONCAT(
                CONVERT(nvarchar(16), CAST(ROUND({qty_norm_expr}, 0) AS int)),
                'x ',
                REPLACE(
                    REPLACE(
                        REPLACE(
                            REPLACE(
                                {title_expr},
                                ';',
                                ','
                            ),
                            CHAR(13),
                            ' '
                        ),
                        CHAR(10),
                        ' '
                    ),
                    '|',
                    '/'
                ),
                {weight_suffix_expr}
            ) AS nvarchar(max)),
            '; '
        ) AS Produkte,
        {total_weight_expr} AS ProduktgewichtTotal
    FROM dbo.WW_Pos p2
    WHERE p2.KopfID = k.Ident
      AND UPPER(COALESCE(p2.PosCode, '')) = 'ART'
      AND {include_product_line_expr}
) pos
OUTER APPLY (
    SELECT TOP 1
        CASE
            WHEN src.keytext LIKE '%fracht_m_vert_mont%' THEN 'Fracht_m_vert_mont'
            WHEN src.keytext LIKE '%fracht_m_vert%' THEN 'Fracht_m_vert'
            WHEN src.keytext LIKE '%fracht_o_vert%' THEN 'Fracht_o_vert'
            WHEN src.keytext LIKE '%fracht mit spediteur%' OR src.keytext LIKE '%lieferung erfolgt ab gawela mit spediteur%' THEN 'fracht_mit_spediteur'
            WHEN src.keytext LIKE '%fracht-tresor-bordstein%' THEN 'Fracht-Tresor-Bordstein'
            WHEN src.keytext LIKE '%fracht-tresor-verwendung%' THEN 'Fracht-Tresor-verwendung'
            WHEN src.keytext LIKE '%selbstabholung%' THEN 'Selbstabholung'
            WHEN src.keytext LIKE '%post%' THEN 'post'
            ELSE NULL
        END AS Liefercode
    FROM (
        SELECT LOWER(COALESCE(p2.Bezeichnung, '')) AS keytext
        FROM dbo.WW_Pos p2
        WHERE p2.KopfID = k.Ident
    ) src
    WHERE src.keytext LIKE '%fracht_m_vert_mont%'
       OR src.keytext LIKE '%fracht_m_vert%'
       OR src.keytext LIKE '%fracht_o_vert%'
       OR src.keytext LIKE '%fracht mit spediteur%'
       OR src.keytext LIKE '%lieferung erfolgt ab gawela mit spediteur%'
       OR src.keytext LIKE '%fracht-tresor-bordstein%'
       OR src.keytext LIKE '%fracht-tresor-verwendung%'
       OR src.keytext LIKE '%selbstabholung%'
       OR src.keytext LIKE '%post%'
    ORDER BY CASE
        WHEN src.keytext LIKE '%fracht_m_vert_mont%' THEN 1
        WHEN src.keytext LIKE '%fracht_m_vert%' THEN 2
        WHEN src.keytext LIKE '%fracht_o_vert%' THEN 3
        WHEN src.keytext LIKE '%fracht mit spediteur%' OR src.keytext LIKE '%lieferung erfolgt ab gawela mit spediteur%' THEN 4
        WHEN src.keytext LIKE '%fracht-tresor-bordstein%' THEN 5
        WHEN src.keytext LIKE '%fracht-tresor-verwendung%' THEN 6
        WHEN src.keytext LIKE '%selbstabholung%' THEN 7
        WHEN src.keytext LIKE '%post%' THEN 8
        ELSE 99
    END
) delivery
OUTER APPLY (
    SELECT TOP 1 sk.Nummer AS Kontakt
    FROM dbo.AVE_StammKontakt sk
    INNER JOIN dbo.AVE_Kontakttyp kt ON kt.Ident = sk.KontaktID
    WHERE sk.AdressID = COALESCE(k.LieferadressID, k.AdressID)
      AND kt.Art IN ('T', 'M')
      AND sk.Nummer IS NOT NULL
    ORDER BY sk.Sortierung, sk._UDatum DESC
) phone
OUTER APPLY (
    SELECT TOP 1 sk.Nummer AS Kontakt
    FROM dbo.AVE_StammKontakt sk
    INNER JOIN dbo.AVE_Kontakttyp kt ON kt.Ident = sk.KontaktID
    WHERE sk.AdressID = COALESCE(k.LieferadressID, k.AdressID)
      AND kt.Art = 'E'
      AND sk.Nummer IS NOT NULL
    ORDER BY sk.Sortierung, sk._UDatum DESC
) mail
WHERE k.Kopf IS NOT NULL
  AND ISNULL(k.Archiv, 0) = 0
  AND ISNULL(k.Statuscode, '') = 'O'
  AND k.Typ = 'SALES'
ORDER BY k.Datum DESC, k.Kopf DESC;
""".strip()

    proc = subprocess.run(
        [sqlcmd, "-S", server, "-d", db, "-f", "65001", "-w", "65535", "-y", "0", "-s", "|", "-Q", query],
        capture_output=True,
        text=False,
        timeout=120,
        check=False,
    )
    stdout_text = _decode_sqlcmd_output(proc.stdout)
    stderr_text = _decode_sqlcmd_output(proc.stderr)
    if proc.returncode != 0:
        detail = (stderr_text or stdout_text or "").strip()
        raise RuntimeError(detail or "SQL-Abfrage konnte nicht ausgeführt werden.")

    rows = []
    for line in stdout_text.splitlines():
        raw = line.strip()
        if not raw:
            continue
        parts = raw.split("|")
        if len(parts) < 18:
            continue
        if parts[0].strip() == "ImportID":
            continue
        rows.append(
            {
                "ImportID": parts[0].strip(),
                "Auftragsnummer": parts[1].strip(),
                "Bestelldatum": parts[2].strip(),
                "Name": parts[3].strip(),
                "Strasse": parts[4].strip(),
                "PLZ": parts[5].strip(),
                "Ort": parts[6].strip(),
                "Land": parts[7].strip(),
                "AuftragName": parts[8].strip(),
                "AuftragPLZ": parts[9].strip(),
                "AuftragOrt": parts[10].strip(),
                "Email": parts[11].strip() or "N/A",
                "Telefon": parts[12].strip() or "N/A",
                "Gewicht": parts[13].strip() or "N/A",
                "Produkte": _extract_product_titles(parts[14]),
                "ProduktgewichtTotal": parts[15].strip(),
                "Notizen": parts[16].strip(),
                "Liefercode": parts[17].strip(),
            }
        )
    return rows
