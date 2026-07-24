import sqlite3

conn = sqlite3.connect(r"C:\Users\EspenGabrielsen\code\ProduksjonsKalkyle\produksjonskalkyle_copy.db")

#resp = conn.execute("select * from products where item_no = 'JD29198'")
#resp = conn.execute("select * from routing_lines where item_no = 'JD29198'")
resp = conn.execute("SELECT * from work_centers")

for line in resp:
    print(line)

conn.close()

