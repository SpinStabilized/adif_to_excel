import adiftools.adiftools as adiftools
import qrzlib
import pandas as pd
from dotenv import load_dotenv
import os

def get_address(calls: list[str]) -> dict[str, str]:

    load_dotenv()
    qrz_call: str = os.getenv("QRZ_CALL")
    qrz_api_key: str = os.getenv("QRZ_API_KEY")

    session: qrzlib.QRZ = qrzlib.QRZ()
    session.authenticate(qrz_call, qrz_api_key)

    addresses: dict[str, str] = {}

    for call in calls:
        call_info: dict = {'name':'', 'addr1':'', 'addr2':'', 'state':'', 'zip':'', 'country':''}

        try:
            response = session.get_call(call)
            if response:
                call_info = {
                    'name':response.name_fmt,
                    'addr1':response.addr1,
                    'addr2':response.addr2,
                    'state':response.state,
                    'zip':response.zip,
                    'country':response.country,
                }
        except session.NotFound as err:
            print(f'Error in get_address({call}, {session}): {err}')
        finally:
            addresses[call] = call_info

    return addresses

def main():
    adi = adiftools.ADIFParser()

    df_adi = adi.read_adi('test.adi') # Use your own adi file
    df_xl = df_adi[['CALL', 'QSO_DATE_OFF', 'TIME_OFF', 'FREQ', 'RST_SENT', 'RST_RCVD', 'MODE']]
    df_xl["Note"] = ""
    addresses: dict[str, str] = get_address(df_xl['CALL'].tolist())
    mapped = df_xl['CALL'].map(addresses).apply(pd.Series)
    df_xl = pd.concat([df_xl, mapped], axis=1)
    df_xl.to_excel('output.xlsx', index=False)

if __name__ == "__main__":
    main()
