import aiofiles
import asyncio
import pandas as pd
import re

HEADER_0 = 'No. Urut'
HEADER_1 = 'No. Kartu'
HEADER_2 = 'Waktu Txn'
HEADER_3 = 'Nominal Transaksi'
HEADER_4 = 'Terminal / ANT'
HEADER_5 = 'Keterangan'


async def readFile(path, df):
    async with aiofiles.open(path) as f:
        is_2nd_line = False
        tmp = []
        async for line in f:
            if "KODE LAPORAN :" in line:
                continue

            if "PERUSAHAAN   :" in line:
                continue

            if line.isspace():
                continue

            if "--------" in line:
                continue

            if "No.    No.Kartu           " in line:
                continue

            if "Urut                      Txn" in line:
                continue

            if "       SUB TOTAL NOMINAL TOP UP   :" in line:
                continue

            if "       GRAND TOTAL NOMINAL TOP UP :" in line:
                continue

            if not is_2nd_line:
                match = re.search(
                        r'\s*(\d+)\s+(\d+)\s+(\d{2}/\d{2}/\d{2})\s+Rp\s+([\d,.]+)\s+([A-Z0-9 -]+)\s+\s+([\sA-Z0-9]+)',
                        line
                        )
                tmp = [
                        int(match.group(1)),
                        match.group(2),
                        match.group(3),
                        int(match.group(4).replace(",", "").replace(".", "")),
                        match.group(5).strip(),
                        match.group(6).strip()
                ]

                is_2nd_line = True
            else:
                match = re.search(
                        r'\s*(\d{2}:\d{2}:\d{2})\s+([A-Z ]+[A-Z, ]+)\s+([A-Z0-9]+)\s*',
                        line
                        )
                tmp[2] = f"{tmp[2]} {match.group(1)}"
                tmp[4] = f"{tmp[4]}{match.group(2)}"
                tmp[5] = f"{tmp[5]}{match.group(3)}"

                tmpDf = pd.DataFrame(data={
                    HEADER_0: [tmp[0]],
                    HEADER_1: [tmp[1]],
                    HEADER_2: [tmp[2]],
                    HEADER_3: [tmp[3]],
                    HEADER_4: [tmp[4]],
                    HEADER_5: [tmp[5]]
                    })
                df = df._append(tmpDf, ignore_index=True)

                tmp = []
                is_2nd_line = False

    return df


async def main():
    df = pd.DataFrame(
        {
            HEADER_0: pd.Series(dtype=int),
            HEADER_1: pd.Series(dtype=str),
            HEADER_2: pd.Series(dtype=str),
            HEADER_3: pd.Series(dtype=int),
            HEADER_4: pd.Series(dtype=str),
            HEADER_5: pd.Series(dtype=str)
        }
    )

    df = await readFile('input.txt', df)

    df.to_excel("output.xlsx", index=False)

if __name__ == "__main__":
    asyncio.run(main())
