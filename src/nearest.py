
import xlwings as xw

def main():
    wb = xw.books["HawkFallsCogi.xlsx"]

    docs = [[str(int(x[0])), int(x[-1]), -1*x[6]] for x in wb.sheets["mb51"].range("A2:P2").expand("down").value]
    find = [[int(x[-1]), x[0]] for x in wb.sheets["ReverseFinal"].range("I2:M2").expand("down").value]

    res = [[x[2]] for x in nearest_neighbor(find, docs)]
    wb.sheets["ReverseFinal"].range("N2").value = res


def nearest_neighbor(to_find, vals):
    for line in to_find:
        program, qty = line

        nearest = qty * 4
        nearest_at = -1
        for i, (_doc, prog, q) in enumerate(vals):
            if program == prog and abs(qty-q) < nearest:
                nearest = abs(qty-q)
                nearest_at = i

        if nearest_at > -1:
            line.append(vals[nearest_at][0])
            del vals[nearest_at]
        else:
            line.append(None)

    return to_find


if __name__ == "__main__":
    main()