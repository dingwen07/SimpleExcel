import sys
import os

try:
    from SimpleExcel import Excel
except:
    sys.path.append(os.getcwd())
    from SimpleExcel import Excel



if __name__ == "__main__":
    e = Excel('./example/example.xlsx', 0)
    print(e.read((1, 2)))
    #       read  1  B
