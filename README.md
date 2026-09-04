# adif-to-excel

Reads an ADIF log, looks up each worked station's mailing address on QRZ.com,
and writes a spreadsheet that a QSL label printer can merge against. This is
for my own use of a Niimbot thermal printer for QSO labels and address labels
but others might find it useful as well.

Everything lives in `main.py`, in five labelled sections: configuration, QRZ
lookups, address building, spreadsheet building, and the command line.

## Setup

```bash
uv sync
```

QRZ's XML API needs a subscription. Put your login in a `.env` file next to the
script (already covered by `.gitignore`):

```
QRZ_CALL=N0CALL
QRZ_PASSWORD=your-qrz-password
```

Real environment variables override the `.env` file if both are set.

## Use

```bash
uv run main.py log.adi                    # writes log.xlsx beside the input
uv run main.py log.adi -o cards.xlsx      # explicit output path
uv run main.py log.adi --no-lookup        # layout only, no QRZ queries spent
uv run main.py log.adi --delimiter ', '   # one-line addresses
uv run main.py log.adi -v                 # per-callsign progress
```

With no `-o`, output goes to the same directory with the same stem and an
`.xlsx` extension. The script refuses to write over its own input.

## Output columns

| Column                | Contents                                        |
| --------------------- | ----------------------------------------------- |
| `CALL`                | worked station                                  |
| `QSO_DATE_OFF`        | `YYYY-MM-DD`, from `QSO_DATE_OFF` or `QSO_DATE` |
| `TIME_OFF`            | `HH:MM`, from `TIME_OFF` or `TIME_ON`           |
| `FREQ`                | MHz, as logged                                  |
| `RST`                 | sent/received, e.g. `599/599`                   |
| `MODE`                | as logged                                       |
| `Note`                | left blank for you to fill in                   |
| `QSL_PSE` / `QSL_TNX` | tick marks                                      |
| `address`             | formatted mailing address, one QSO per row      |
| `status`              | why the address looks the way it does           |

Header names are what a label template merges against, so renaming one means
editing the template to match. They live in `COLUMNS` near the top of `main.py`.

## The `status` column

Sort or filter on this to find the cards needing hand work:

| Status                     | Meaning                                                                            |
| -------------------------- | ---------------------------------------------------------------------------------- |
| `ok`                       | mail directly to the operator                                                      |
| `via manager`              | QRZ listed a QSL manager; the card is addressed to them with a "QSL for CALL" line |
| `not on QRZ`               | no QRZ record, or the lookup failed                                                |
| `no street address on QRZ` | record exists but `addr1` is empty                                                 |
| `country not recognized`   | QRZ's country string did not map to an ISO code                                    |
| `QSL manager not on QRZ`   | a manager was named but could not be resolved                                      |

If you hit `country not recognized`, add the country string to
`COUNTRY_OVERRIDES` in `main.py`.

## Caching

`qrzlib` already caches lookups in `~/.local/qrz-cache.sqlite3` for three years,
including negative results. Re-running on the same log costs no QRZ queries.
Delete that file to force a refresh.
