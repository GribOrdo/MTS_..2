def nds_count(summ, typ, nds):
    # вычислить НДС из суммы с НДС
    if typ == 1:
        return (summ / (1 + nds) - summ) * (-1)
    # накинуть НДС на сумму без НДС
    if typ == 0:
        return summ * (1 + nds)


print(f'{nds_count(204817.72, 1, 0.22):.2f}')
