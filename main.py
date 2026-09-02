import adiftools.adiftools as adiftools


def main():
    adi = adiftools.ADIFParser()

    df_adi = adi.read_adi('test.adi') # Use your own adi file
    df_xl = df_adi[['CALL', 'QSO_DATE_OFF', 'TIME_OFF', 'FREQ', 'RST_SENT', 'RST_RCVD', 'MODE']]
    df_xl["Note"] = ""
    df_xl.to_excel('output.xlsx', index=False)

if __name__ == "__main__":
    main()
