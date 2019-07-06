from SimpleExcel import Excel

if __name__ == "__main__":
    e = Excel('./test.xlsx', 0)
    print('A1: {value}'.format(value=e.read([1,1])))
    print('A2: {value}'.format(value=e.read(Excel.convert('A2'))))
    print(e.read_range([1,1], [11,2]))
    print('B2: {value}'.format(value=e.write([2,1], 'Test')))
    print('B2: {value}'.format(value=e.read([2,1])))
