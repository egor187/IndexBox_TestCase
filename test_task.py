import sqlite3

from pandas import read_sql_query

con = sqlite3.connect("test.db")
cur = con.cursor()


query = """
    SELECT * FROM testidprod WHERE partner is NULL AND state is NULL and bs == 0 and factor == 0 or factor == 1;
"""

cur.execute(query)
output = cur.fetchall()

tables = read_sql_query(query, con)


print(tables)

