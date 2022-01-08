def dict_factory(cursor, row):
    """Func for sqlite row_factory to get dicts instead tuples"""
    d = {}
    for idx, col in enumerate(cursor.description):
        d[col[0]] = row[idx]
    return d
