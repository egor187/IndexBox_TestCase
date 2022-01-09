import sqlite3
from pathlib import Path

import pandas

from helpers import dict_factory, calc_cagr, word_writer, NEW_FACTOR


pandas.set_option("display.max_columns", 20)
pandas.set_option("display.width", 500)

con = sqlite3.connect("test.db")
con.row_factory = dict_factory
cur = con.cursor()


def main():
    avg_query = """
        select factor, year, sum(res) as summary from testidprod group by factor, year 
        having partner is null and state is null and bs == 0 and factor == 1 or factor == 2;
    """

    cur.execute(avg_query)
    avg_output = cur.fetchall()
    con.close()

    df = pandas.DataFrame(
        {
            (elem["factor"], elem["year"]): {"World": elem["summary"]}
            for elem in avg_output
        }
    )
    df.columns.names = ["Factor", "Year"]
    for year in [
        year["year"] for year in filter(lambda x: x["factor"] == 1, avg_output)
    ]:
        df[NEW_FACTOR, year] = df[2][year] / df[1][year]

    df = df.round(2)
    df.to_excel("report.xlsx")
    cagr = calc_cagr(df)

    df = df.drop(columns=[1, 2])
    word_writer(df, Path("report.docx"), float(cagr))


if __name__ == "__main__":
    main()
