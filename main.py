import adiftools.adiftools as adiftools


def main():
    adi = adiftools.ADIFParser()

    df_adi = adi.read_adi('test.adi') # Use your own adi file
    df_xl = df_adi[['FREQ', 'RST_SENT', 'RST_RCVD', 'CALL', 'QSO_DATE_OFF', 'TIME_OFF', 'MODE']]
    df_xl.to_excel('output.xlsx', index=False)

if __name__ == "__main__":
    main()
