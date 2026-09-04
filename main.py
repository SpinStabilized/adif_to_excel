#!/usr/bin/env python3
"""Turn an ADIF log into a QSL-label spreadsheet.

Reads an ADIF file, looks up each worked station's mailing address on QRZ.com,
and writes an .xlsx that a label printer can merge against.

    python main.py log.adi                    # writes log.xlsx beside the input
    python main.py log.adi -o cards.xlsx      # explicit output path
    python main.py log.adi --no-lookup        # layout only, no QRZ queries spent
    python main.py log.adi --delimiter ', '   # one-line addresses
    python main.py log.adi -v                 # per-callsign progress

QRZ credentials come from QRZ_CALL and QRZ_PASSWORD, in the environment or in a
.env file in the working directory. The XML API needs a QRZ subscription.

Sections below, in order: configuration, QRZ lookups, address building,
spreadsheet building, command line.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from dataclasses import dataclass
from enum import StrEnum
from functools import lru_cache
from pathlib import Path
from typing import Callable

import adiftools.adiftools as adiftools
import pandas as pd
import pycountry
import qrzlib
from dotenv import load_dotenv
from i18naddress import InvalidAddressError, format_address
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

__version__ = "0.2.0"

log = logging.getLogger("adif_to_excel")


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

#: Environment variables (or .env keys) holding QRZ credentials.
ENV_CALLSIGN = "QRZ_CALL"
ENV_PASSWORD = "QRZ_PASSWORD"

SHEET_NAME = "QSL Labels"

#: Output column order. These header names are what a label template merges
#: against, so renaming one here means editing the template to match.
COLUMNS = [
    "CALL",
    "QSO_DATE_OFF",
    "TIME_OFF",
    "FREQ",
    "RST",
    "MODE",
    "Note",
    "QSL_PSE",
    "QSL_TNX",
    "address",
    "status",
]

#: Tick mark written into the QSL_PSE / QSL_TNX columns.
CHECKED = "\u2714"

#: Rendered column width in characters. Anything absent gets DEFAULT_WIDTH.
COLUMN_WIDTHS = {
    "CALL": 10,
    "QSO_DATE_OFF": 13,
    "TIME_OFF": 9,
    "FREQ": 10,
    "RST": 9,
    "MODE": 8,
    "Note": 20,
    "QSL_PSE": 9,
    "QSL_TNX": 9,
    "address": 42,
    "status": 22,
}
DEFAULT_WIDTH = 14

#: QRZ country strings that pycountry guesses badly or not at all. Add entries
#: as you run into them; keys are compared upper-cased. These are cheaper and
#: more reliable than pycountry's fuzzy search.
COUNTRY_OVERRIDES: dict[str, str] = {
    "UNITED STATES": "US",
    "UNITED STATES OF AMERICA": "US",
    "UK": "GB",
    "ENGLAND": "GB",
    "SCOTLAND": "GB",
    "WALES": "GB",
    "NORTHERN IRELAND": "GB",
    "FED. REP. OF GERMANY": "DE",
    "CZECH REPUBLIC": "CZ",
    "SLOVAK REPUBLIC": "SK",
    "MACEDONIA": "MK",
    "MOLDOVA": "MD",
    "RUSSIA": "RU",
    "ASIATIC RUSSIA": "RU",
    "EUROPEAN RUSSIA": "RU",
    "SOUTH KOREA": "KR",
    "TAIWAN": "TW",
    "VIETNAM": "VN",
    "LAOS": "LA",
    "BURMA": "MM",
    "IRAN": "IR",
    "SYRIA": "SY",
    "BOLIVIA": "BO",
    "VENEZUELA": "VE",
    "TANZANIA": "TZ",
    "IVORY COAST": "CI",
    "CAPE VERDE": "CV",
    "EAST TIMOR": "TL",
    "SWAZILAND": "SZ",
    "CANARY ISLANDS": "ES",
    "BALEARIC ISLANDS": "ES",
    "AZORES": "PT",
    "MADEIRA ISLANDS": "PT",
    "CORSICA": "FR",
    "SARDINIA": "IT",
    "SICILY": "IT",
    "HAWAII": "US",
    "ALASKA": "US",
    "PUERTO RICO": "PR",
    "SAINT HELENA": "SH",
    "REP. OF KOREA": "KR",
    "REP. OF SOUTH AFRICA": "ZA",
    "CONGO (DEM. REP.)": "CD",
    "DEM. REP. OF THE CONGO": "CD",
    "BOSNIA-HERZEGOVINA": "BA",
}

#: Matches a callsign inside a free-text QSL-manager field.
#:
#: The prefix is one to three alphanumerics so that both one-letter prefixes
#: (W3HNK, M0OXO, K4GMX) and digit-bearing prefixes (9A1CMA, 4X4ABC) match.
#: Requiring two characters before the digit silently skips the single-letter
#: prefixes, which covers some of the busiest QSL managers there are.
CALLSIGN_RE = re.compile(r"\b([A-Z0-9]{1,3}\d[A-Z]{1,4})\b")

#: Matches a preferred name in quotes inside a QRZ name field, as in
#: 'Brian J. "Bri" McLaughlin'. Straight and curly double quotes both count.
#:
#: Single quotes are deliberately not matched: apostrophes are common in
#: surnames (O'Brien, D'Angelo) and a pair of them would swallow the name.
NICKNAME_RE = re.compile(r"[\"\u201c\u201d]\s*([^\"\u201c\u201d]+?)\s*[\"\u201c\u201d]")

#: QSL-manager values that describe a routing method rather than a person.
NON_MANAGER_TOKENS = frozenset(
    {
        "",
        "NONE",
        "NO",
        "N",
        "BURO",
        "BUREAU",
        "VIA BURO",
        "LOTW",
        "EQSL",
        "DIRECT",
        "DIRECT ONLY",
        "OQRS",
        "CLUBLOG",
        "CLUB LOG",
        "QRZ",
        "SEE QRZ",
        "SEE QRZ.COM",
        "HOME CALL",
        "SASE",
    }
)

#: Takes a callsign, returns a QRZ record or None. Injected rather than called
#: directly so that --no-lookup is a one-line substitution.
CallsignLookup = Callable[[str], "qrzlib.QRZRecord | None"]


# ---------------------------------------------------------------------------
# QRZ lookups
# ---------------------------------------------------------------------------
#
# qrzlib already caches every lookup in ~/.local/qrz-cache.sqlite3 with a
# three-year TTL, including negative results for callsigns that were not found.
# There is no reason to keep a second dictionary cache on top of it, so this
# script does not. Delete that file to force a refresh.


class MissingCredentials(RuntimeError):
    """Raised when the QRZ username or password is not configured."""


class QrzClient:
    """Authenticated QRZ.com XML API client.

    Authentication is deferred until the first lookup, so constructing the
    client never touches the network.
    """

    def __init__(self, username: str, password: str) -> None:
        self._username = username
        self._password = password
        self._session: qrzlib.QRZ | None = None

    @classmethod
    def from_env(cls, dotenv_path: str | None = None) -> "QrzClient":
        """Build a client from QRZ_CALL / QRZ_PASSWORD.

        Looks for a .env in the working directory first, then one sitting beside
        this script. That second lookup is what lets a wrapper on PATH invoke the
        tool from anywhere without exporting credentials globally. Real
        environment variables override both.

        Raises:
            MissingCredentials: if either value is absent or blank.
        """
        load_dotenv(dotenv_path)
        beside_script = Path(__file__).resolve().with_name(".env")
        if beside_script.is_file():
            load_dotenv(beside_script)  # does not overwrite anything already set

        username = (os.getenv(ENV_CALLSIGN) or "").strip()
        password = os.getenv(ENV_PASSWORD) or ""
        if not username or not password:
            raise MissingCredentials(
                f"Set {ENV_CALLSIGN} and {ENV_PASSWORD} in the environment, in a .env "
                f"file in this directory, or in {beside_script}. Run with --no-lookup "
                "to skip QRZ entirely and produce a sheet with blank addresses."
            )
        return cls(username, password)

    def _connect(self) -> qrzlib.QRZ:
        if self._session is None:
            log.debug("Authenticating to QRZ as %s", self._username)
            session = qrzlib.QRZ()
            session.authenticate(self._username, self._password)
            self._session = session
        return self._session

    def __call__(self, callsign: str) -> qrzlib.QRZRecord | None:
        """Look up a callsign, returning None if QRZ has no record for it.

        Transport and XML failures are logged and swallowed rather than raised.
        One flaky lookup partway through a hundred-QSO log should leave a single
        blank address behind, not throw away the whole run.
        """
        callsign = callsign.strip().upper()
        if not callsign:
            return None
        try:
            return self._connect().get_call(callsign)
        except qrzlib.QRZ.NotFound:
            log.info("%s is not listed on QRZ", callsign)
        except (qrzlib.QRZ.SessionError, qrzlib.QRZ.XMLError, ValueError) as err:
            log.warning("QRZ lookup failed for %s: %s", callsign, err)
        return None


# ---------------------------------------------------------------------------
# Address building
# ---------------------------------------------------------------------------
#
# Two decisions live here. Who gets the card: if QRZ lists a QSL manager, it
# goes to the manager with a "QSL for <call>" attention line, otherwise to the
# operator. And how it is laid out: country-specific ordering from
# google-i18n-address, with a plain fallback when that library rejects a record.


class AddressStatus(StrEnum):
    """Why an address looks the way it does. Written to the sheet for triage."""

    OK = "ok"
    VIA_MANAGER = "via manager"
    NOT_ON_QRZ = "not on QRZ"
    NO_STREET_ADDRESS = "no street address on QRZ"
    UNKNOWN_COUNTRY = "country not recognized"
    MANAGER_NOT_FOUND = "QSL manager not on QRZ"


@dataclass(frozen=True)
class MailingAddress:
    """A formatted address plus the reason it turned out that way.

    Every path through resolve_mailing_address returns one of these, so a caller
    never has to work out whether it got an address, a placeholder or None.
    """

    text: str
    status: AddressStatus

    @property
    def is_mailable(self) -> bool:
        return self.status in (AddressStatus.OK, AddressStatus.VIA_MANAGER)


def normalize_country(country_name: str) -> str:
    """Tidy up the cosmetic noise in a country string before matching it.

    Logs carry things like "Turks and  Caicos Islands" (a doubled space) and
    "Trinidad & Tobago". Neither matches anything in pycountry, not even its
    fuzzy search, purely because of punctuation. This collapses whitespace,
    spells out "&", and expands a leading "St." to "Saint".

    Keys in COUNTRY_OVERRIDES are compared against the output of this function,
    so write them normalized: single spaces, "and" rather than "&", "Saint"
    spelled out.
    """
    text = country_name.replace("&", " and ")
    text = re.sub(r"\bSt\.?\s+", "Saint ", text, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", text).strip()


@lru_cache(maxsize=None)
def country_to_iso2(country_name: str | None) -> str | None:
    """Map a QRZ country string to an ISO 3166-1 alpha-2 code.

    Tries three things in order, stopping at the first hit:

    1. COUNTRY_OVERRIDES, for names pycountry does not carry at all
       ("England", "Fed. Rep. of Germany").
    2. An exact pycountry lookup, which matches names, official names and the
       alpha-2/alpha-3 codes, case-insensitively.
    3. A fuzzy search, which is slow and occasionally confidently wrong, so it
       is the last resort rather than the first move.

    Cached, because the same handful of countries repeats across a log.
    """
    if not country_name or not country_name.strip():
        return None

    name = normalize_country(country_name)
    if name.upper() in COUNTRY_OVERRIDES:
        return COUNTRY_OVERRIDES[name.upper()]

    try:
        return pycountry.countries.lookup(name).alpha_2
    except LookupError:
        pass

    try:
        match = pycountry.countries.search_fuzzy(name)[0]
    except LookupError:
        log.warning("No ISO country code for %r (add it to COUNTRY_OVERRIDES)", country_name)
        return None

    log.debug("Fuzzy-matched country %r to %s (%s)", country_name, match.alpha_2, match.name)
    return match.alpha_2


def extract_manager_call(qslmgr: str | None, station_call: str = "") -> str | None:
    """Pull a manager callsign out of QRZ's free-text qslmgr field.

    Returns None when the field is empty, describes a routing method rather than
    a person ("LOTW", "BURO"), or simply repeats the station's own call.
    """
    if not isinstance(qslmgr, str):
        return None
    text = qslmgr.strip().upper()
    if not text or text in NON_MANAGER_TOKENS:
        return None
    match = CALLSIGN_RE.search(text)
    if not match:
        return None
    manager = match.group(1)
    return None if manager == station_call.strip().upper() else manager


def _tidy(value: str | None) -> str:
    """Collapse runs of whitespace and trim. The same noise that breaks country
    matching also shows up in street addresses, where it just looks sloppy on a
    printed label."""
    return re.sub(r"\s+", " ", value).strip() if value else ""


def apply_nickname(text: str) -> str:
    """Reduce a name to its quoted preferred form, if it has one.

    QRZ records often carry a nickname in quotes: 'Brian J. "Bri"'. Someone who
    writes that on their profile is telling you what to call them, so the
    nickname replaces everything before it and the quotes come off:

        'Brian J. "Bri"'              -> 'Bri'
        'Brian J. "Bri" McLaughlin'   -> 'Bri McLaughlin'
        'Brian J. McLaughlin'         -> 'Brian J. McLaughlin'

    A quote character that is not part of a matched pair (an empty '""', an
    unclosed one) is dropped rather than printed on a label.

    Delete the call to this in _recipient_name if you would rather address cards
    formally.
    """
    match = NICKNAME_RE.search(text)
    if not match:
        return _tidy(re.sub(r"[\"\u201c\u201d]", " ", text))
    return _tidy(f"{match.group(1)} {text[match.end():]}")


def _recipient_name(record: qrzlib.QRZRecord, callsign: str) -> str:
    """Best available addressee: "First Last (CALL)", or just the call.

    QRZ splits names into fname and name; name_fmt is its own pre-formatted
    version, used only as a fallback. Either can carry a quoted nickname, which
    apply_nickname promotes to the mailing name. A record with no name at all
    still has a perfectly good addressee in the callsign.
    """
    given = apply_nickname(record.fname or "")
    surname = _tidy(record.name)
    name = _tidy(f"{given} {surname}") or apply_nickname(record.name_fmt or "")
    return f"{name} ({callsign})" if name else callsign


def _format_lines(record: qrzlib.QRZRecord, iso2: str, name: str, attention: str) -> list[str]:
    """Lay the address out per the destination country's conventions."""
    fields = {
        "name": _tidy(name),
        "company_name": _tidy(attention),
        "street_address": _tidy(record.addr1),
        "city": _tidy(record.addr2),
        "country_area": _tidy(record.state),
        "postal_code": _tidy(record.zip),
        "country_code": iso2,
    }
    fields = {k: v for k, v in fields.items() if v}

    try:
        lines = format_address(fields, latin=True).split("\n")
    except InvalidAddressError as err:
        log.debug("i18n formatting rejected %s (%s); using plain layout", record.call, err)
        lines = [
            line
            for line in (
                _tidy(name),
                _tidy(attention),
                _tidy(record.addr1),
                _tidy(record.addr2),
                _tidy(record.state),
                _tidy(record.zip),
            )
            if line
        ]

    country = pycountry.countries.get(alpha_2=iso2)
    country_line = (country.name if country else iso2).upper()
    if not lines or lines[-1].strip().upper() != country_line:
        lines.append(country_line)
    return lines


def resolve_mailing_address(
    callsign: str,
    lookup: CallsignLookup,
    *,
    qsl_via: str = "",
    delimiter: str = "\n",
) -> MailingAddress:
    """Return the postal address for a callsign, following a QSL manager if listed.

    A manager can come from two places. The log's own QSL_VIA field is checked
    first, because it records the route for that specific contact, which is what
    you want for a DXpedition or contest call whose QRZ profile has since moved
    on. QRZ's qslmgr field is the fallback.

    Args:
        callsign: the worked station's call, as it appears in the log.
        lookup: a callable returning a QRZRecord or None.
        qsl_via: the QSL_VIA field from the log, if the log carried one. Free
            text is fine; anything without a callsign in it is ignored.
        delimiter: what joins the address lines. "\\n" suits an Excel cell with
            wrapping turned on; ", " suits a single-line label field.

    Returns:
        A MailingAddress whose text is empty whenever the status is anything
        other than ok or via manager.
    """
    callsign = callsign.strip().upper()
    record = lookup(callsign)
    if record is None:
        return MailingAddress("", AddressStatus.NOT_ON_QRZ)

    manager_call = extract_manager_call(qsl_via, callsign)
    if manager_call:
        log.debug("%s: the log routes via %s", callsign, manager_call)
    else:
        manager_call = extract_manager_call(record.qslmgr, callsign)

    if manager_call:
        manager_record = lookup(manager_call)
        if manager_record is None:
            log.warning("%s lists manager %s, who is not on QRZ", callsign, manager_call)
            return MailingAddress("", AddressStatus.MANAGER_NOT_FOUND)
        source, status = manager_record, AddressStatus.VIA_MANAGER
        name = _recipient_name(manager_record, manager_call)
        attention = f"QSL for {callsign}"
    else:
        source, status = record, AddressStatus.OK
        name = _recipient_name(record, callsign)
        attention = ""

    if not source.addr1:
        return MailingAddress("", AddressStatus.NO_STREET_ADDRESS)

    iso2 = country_to_iso2(source.country)
    if not iso2:
        return MailingAddress("", AddressStatus.UNKNOWN_COUNTRY)

    return MailingAddress(delimiter.join(_format_lines(source, iso2, name, attention)), status)


# ---------------------------------------------------------------------------
# Spreadsheet building
# ---------------------------------------------------------------------------


def _column(frame: pd.DataFrame, *names: str) -> pd.Series:
    """Return the first of the named columns present in the frame, as strings.

    ADIF makes several of these fields optional and different loggers emit
    different subsets, so falling through candidates (and finally to blanks)
    keeps a missing QSO_DATE_OFF from raising a KeyError mid-run.
    """
    for name in names:
        if name in frame.columns:
            return frame[name].fillna("").astype(str).str.strip()
    log.debug("None of %s present in the log; leaving the column blank", names)
    return pd.Series([""] * len(frame), index=frame.index, dtype="object")


def _format_date(value: str) -> str:
    """20260829 becomes 2026-08-29; anything unexpected passes through."""
    if len(value) >= 8 and value[:8].isdigit():
        return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    return value


def _format_time(value: str) -> str:
    """1453 or 145307 becomes 14:53; anything unexpected passes through.

    ADIF permits both HHMM and HHMMSS, which is why this slices rather than
    parsing with a fixed format string. An empty value stays empty rather than
    becoming a misleading 00:00.
    """
    if not value:
        return ""
    padded = value.zfill(4)[:4]
    return f"{padded[:2]}:{padded[2:4]}" if padded.isdigit() else value


def _combine_rst(sent: pd.Series, received: pd.Series) -> pd.Series:
    """Join sent and received reports as 599/599, or blank if neither exists."""
    joined = sent + "/" + received
    return joined.where(joined != "/", "")


def build_label_rows(
    log_frame: pd.DataFrame,
    lookup: CallsignLookup,
    *,
    delimiter: str = "\n",
) -> pd.DataFrame:
    """Build the label sheet from a parsed ADIF dataframe.

    Each QSO becomes one row, so working the same station twice produces two
    cards. Addresses are resolved once per distinct call-and-route pair, not
    once per QSO.

    Args:
        log_frame: the dataframe from adiftools.ADIFParser.read_adi.
        lookup: a callable returning a QRZRecord or None.
        delimiter: what joins address lines inside the cell.

    Returns:
        A dataframe whose columns are exactly COLUMNS.
    """
    calls = _column(log_frame, "CALL").str.upper()
    routes = _column(log_frame, "QSL_VIA")

    # Keyed on (call, route) rather than call alone, because QSL_VIA belongs to
    # the QSO, not the station. The same call worked on two DXpeditions can have
    # two different managers.
    keys = list(zip(calls, routes))
    addresses: dict[tuple[str, str], MailingAddress] = {}
    unique_keys = list(dict.fromkeys(key for key in keys if key[0]))
    for index, key in enumerate(unique_keys, start=1):
        call, route = key
        log.info("Resolving address %d/%d: %s", index, len(unique_keys), call)
        addresses[key] = resolve_mailing_address(
            call, lookup, qsl_via=route, delimiter=delimiter
        )

    return pd.DataFrame(
        {
            "CALL": calls,
            "QSO_DATE_OFF": _column(log_frame, "QSO_DATE_OFF", "QSO_DATE").map(_format_date),
            "TIME_OFF": _column(log_frame, "TIME_OFF", "TIME_ON").map(_format_time),
            "FREQ": _column(log_frame, "FREQ"),
            "RST": _combine_rst(_column(log_frame, "RST_SENT"), _column(log_frame, "RST_RCVD")),
            "MODE": _column(log_frame, "MODE"),
            "Note": "",
            "QSL_PSE": CHECKED,
            "QSL_TNX": CHECKED,
            "address": [addresses[k].text if k in addresses else "" for k in keys],
            "status": [str(addresses[k].status) if k in addresses else "" for k in keys],
        },
        columns=COLUMNS,
    )


def write_workbook(rows: pd.DataFrame, destination: Path) -> None:
    """Write the rows to an .xlsx with widths and wrapping already set.

    Multi-line addresses are unreadable in Excel unless the cell wraps, so this
    formatting pass is doing real work rather than decoration.
    """
    destination.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(destination, engine="openpyxl") as writer:
        rows.to_excel(writer, index=False, sheet_name=SHEET_NAME)
        worksheet = writer.sheets[SHEET_NAME]

        for position, name in enumerate(rows.columns, start=1):
            letter = get_column_letter(position)
            worksheet.column_dimensions[letter].width = COLUMN_WIDTHS.get(name, DEFAULT_WIDTH)
            if name == "address":
                for cell in worksheet[letter][1:]:  # skip the header row
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

        worksheet.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

EXIT_OK = 0
EXIT_ERROR = 1


def default_output_path(source: Path) -> Path:
    """/logs/2026-08.adi becomes /logs/2026-08.xlsx.

    Same directory, same stem, Excel extension. Uses with_suffix so a stem
    containing dots ("field-day.v2.adi") keeps everything but the last suffix.
    """
    return source.with_suffix(".xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="adif-to-excel",
        description=(
            "Read an ADIF log, look up each station's mailing address on QRZ.com, "
            "and write a spreadsheet ready for a QSL label printer."
        ),
        epilog=(
            f"QRZ credentials are read from {ENV_CALLSIGN} and {ENV_PASSWORD}, either "
            "in the environment or in a .env file in the working directory."
        ),
    )
    parser.add_argument("input", type=Path, help="ADIF file to read (.adi)")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        metavar="XLSX",
        help="output workbook (default: the input path with a .xlsx extension)",
    )
    parser.add_argument(
        "--delimiter",
        default="\n",
        help=(
            "what joins address lines inside a cell (default: a newline). "
            "Use --delimiter ', ' for label templates that want one line."
        ),
    )
    parser.add_argument(
        "--no-lookup",
        action="store_true",
        help="skip QRZ entirely and leave addresses blank, for testing the layout",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="show per-callsign progress and debug detail",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def _summarize(rows: pd.DataFrame) -> None:
    """Log one line per status so problems are visible without opening Excel."""
    mailable = {AddressStatus.OK, AddressStatus.VIA_MANAGER}
    for status, count in rows["status"].value_counts().items():
        level = logging.INFO if status in mailable else logging.WARNING
        log.log(level, "%4d %s", count, status or "no address")


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    # qrzlib calls logging.basicConfig() at import time, which would make ours a
    # no-op and leave its format in place, so take the root logger back.
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)-8s %(message)s",
        force=True,
    )

    source: Path = args.input
    if not source.is_file():
        log.error("No such file: %s", source)
        return EXIT_ERROR

    destination: Path = args.output or default_output_path(source)
    if destination.resolve() == source.resolve():
        log.error("Refusing to overwrite the input file: %s", source)
        return EXIT_ERROR

    if args.no_lookup:
        log.info("Running with --no-lookup: addresses will be blank")
        lookup: CallsignLookup = lambda callsign: None  # noqa: E731
    else:
        try:
            lookup = QrzClient.from_env()
        except MissingCredentials as err:
            log.error("%s", err)
            return EXIT_ERROR

    log.info("Reading %s", source)
    log_frame = adiftools.ADIFParser().read_adi(str(source))
    if log_frame.empty:
        log.error("No QSOs found in %s", source)
        return EXIT_ERROR
    log.info("Found %d QSO(s)", len(log_frame))

    rows = build_label_rows(log_frame, lookup, delimiter=args.delimiter)
    _summarize(rows)

    write_workbook(rows, destination)
    log.info("Wrote %d row(s) to %s", len(rows), destination)
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
