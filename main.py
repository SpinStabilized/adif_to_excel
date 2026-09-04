import adiftools.adiftools as adiftools
import qrzlib
import pandas as pd
from dotenv import load_dotenv
import os
import re
import pycountry
from i18naddress import format_address, InvalidAddressError

# --- 1. country name -> ISO alpha-2 -------------------------------------

_COUNTRY_OVERRIDES = {
    "UNITED STATES": "US",
    "UK": "GB",
    "ENGLAND": "GB",
    # add QRZ-specific oddities here as you encounter them
}

_CALLSIGN_RE = re.compile(r'\b([A-Z0-9]{2,3}\d[A-Z]{1,4})\b')

manager_cache: dict[str, dict[str, str]] = {}
call_cache: dict[str, dict[str, str]] = {}

_session = None

def _get_session():
    global _session
    if _session is None:
        load_dotenv()
        _session = qrzlib.QRZ()
        _session.authenticate(os.getenv("QRZ_CALL"), os.getenv("QRZ_PASSWORD"))
    return _session

def country_to_iso2(country_name):
    if not isinstance(country_name, str) or not country_name.strip():
        return None
    key = country_name.strip().upper()
    if key in _COUNTRY_OVERRIDES:
        return _COUNTRY_OVERRIDES[key]
    try:
        return pycountry.countries.search_fuzzy(country_name)[0].alpha_2
    except LookupError:
        return None

def extract_manager_call(qslmgr):
    if not isinstance(qslmgr, str):
        return None
    text = qslmgr.strip().upper()
    if text in ("", "NONE", "BURO", "LOTW", "EQSL", "DIRECT"):
        return None
    match = _CALLSIGN_RE.search(text)
    return match.group(1) if match else None

def resolve_manager_info(mgr_call, cache):
    if mgr_call not in cache:
        cache[mgr_call] = get_address(mgr_call)  # {} if not found
    return cache[mgr_call] or None

def _field(response, name, default=""):
    value = getattr(response, name, default)
    return value if value is not None else default

def build_mailing_address(call, info, manager_cache, delimiter="\n"):
    mgr_call = extract_manager_call(info.get('qslmgr'))

    if mgr_call:
        mgr_info = resolve_manager_info(mgr_call, manager_cache)
        if not mgr_info:
            return f"[MANUAL LOOKUP NEEDED: manager {mgr_call} not found on QRZ]"
        src = mgr_info
        recipient = f"{src.get('fname','')} {src.get('name2','')}".strip() or src.get('name')
        recipient_call = mgr_call
        attn = f"QSL for {call}"
    else:
        src = info
        recipient = f"{info.get('fname','')} {info.get('name2','')}".strip() or info.get('name')
        recipient_call = call
        attn = info.get('attn')

    if recipient and recipient_call:
        recipient = f"{recipient} ({recipient_call})"

    iso2 = country_to_iso2(src.get('country'))
    if not iso2 or not src.get('addr1'):
        return None

    addr = {
        'name': recipient or None,
        'company_name': attn or None,
        'street_address': src.get('addr1'),
        'city': src.get('addr2'),
        'country_area': src.get('state'),
        'postal_code': src.get('zip'),
        'country_code': iso2,
    }
    addr = {k: v for k, v in addr.items() if v}

    try:
        formatted = format_address(addr, latin=True)
        lines = formatted.split("\n")
    except InvalidAddressError:
        lines = [v for v in (src.get('addr1'), src.get('addr2'),
                              src.get('state'), src.get('zip')) if v]

    country_line = pycountry.countries.get(alpha_2=iso2).name.upper()
    if not lines or lines[-1].strip().upper() != country_line:
        lines.append(country_line)

    return delimiter.join(lines)

def get_call_info(call: str) -> dict[str, str]:
    """Fetch and cache QRZ info for a callsign (avoids re-querying repeat contacts)."""
    if call not in call_cache:
        call_cache[call] = get_address(call)
    return call_cache[call]

def get_address(call: str) -> dict[str, str]:

    session = _get_session()
    call_info: dict[str, str] = {}

    try:
        response = session.get_call(call)
        if response:
            call_info = {
                'name': _field(response, 'name_fmt'),
                'addr1': _field(response, 'addr1'),
                'addr2': _field(response, 'addr2'),
                'state': _field(response, 'state'),
                'zip': _field(response, 'zip'),
                'country': _field(response, 'country'),
                'attn': _field(response, 'attn'),
                'qslmgr': _field(response, 'qslmgr'),
                'fname': _field(response, 'fname'),
                'name2': _field(response, 'name'),
            }
    except session.NotFound as err:
        print(f'Error in get_address({call}): {err}')

    return call_info

def main():
    adi = adiftools.ADIFParser()

    df_adi = adi.read_adi('test.adi') # Use your own adi file
    df_xl = df_adi[['CALL', 'QSO_DATE_OFF', 'TIME_OFF', 'FREQ', 'RST_SENT', 'RST_RCVD', 'MODE']]
    df_xl["Note"] = ""
    df_xl['QSL_PSE'] = "✔"
    df_xl['QSL_TNX'] = "✔"

    df_xl['QSO_DATE_OFF'] = pd.to_datetime(df_xl['QSO_DATE_OFF'], format='%Y%m%d').dt.strftime('%Y-%m-%d')
    df_xl['TIME_OFF'] = df_xl['TIME_OFF'].astype(str).str.zfill(4)
    df_xl['TIME_OFF'] = pd.to_datetime(df_xl['TIME_OFF'], format='%H%M').dt.strftime('%H:%M')

    rst_pos = df_xl.columns.get_loc('RST_SENT')
    df_xl['RST'] = df_xl['RST_SENT'].fillna('') + '/' + df_xl['RST_RCVD'].fillna('')
    df_xl = df_xl.drop(columns=['RST_SENT', 'RST_RCVD'])
    df_xl.insert(rst_pos, 'RST', df_xl.pop('RST'))

    df_xl['address'] = [
        build_mailing_address(call, get_call_info(call), manager_cache)
        for call in df_xl['CALL']
    ]
    df_xl.to_excel('output.xlsx', index=False)

if __name__ == "__main__":
    main()
