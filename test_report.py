from report import clean, summarize

rows = [
    {"date": "2026-01-03", "product": " laptop ", "region": "cairo", "quantity": "2", "unit_price": "$1,000"},
    {"date": "03/01/2026", "product": "Laptop", "region": "Cairo", "quantity": "2", "unit_price": "1000"},  # duplicate
    {"date": "bad", "product": "Mouse", "region": "Giza", "quantity": "1", "unit_price": "5"},
    {"date": "2026-01-04", "product": "Mouse", "region": "Giza", "quantity": "", "unit_price": "5"},
]
good, problems = clean(rows)
assert len(good) == 1 and good[0][1:] == ("Laptop", "Cairo", 2, 1000.0)
assert [p[1] for p in problems] == ["duplicate row", "invalid date, quantity or price", "invalid date, quantity or price"]
assert summarize(good, 1) == [("Laptop", 2000.0)]
print("ok")
