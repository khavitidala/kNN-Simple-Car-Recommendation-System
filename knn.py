from copy import deepcopy
from xlwt import Workbook
import xlrd
import math
import xlwt

def baca_data(path):
    data = xlrd.open_workbook(path)
    wsData = data.sheet_by_index(0)
    df = []
    for i in range(1, 18):
        tmp = []
        d = wsData.cell(i, 0)
        tmp.append(d.value)
        d = wsData.cell(i, 1)
        tmp.append(d.value)
        d = wsData.cell(i, 2)
        tmp.append(d.value)
        d = wsData.cell(i, 3)
        tmp.append(d.value)
        d = wsData.cell(i, 4)
        tmp.append(d.value)
        d = wsData.cell(i, 5)
        tmp.append(d.value)
        df.append(tmp)
    return df

def normalisasi(x, up, down):
    return float(float(x-down)/float(up-down))

def prapemrosesan(training, testing):
    norm = training[:]
    norm.append(testing)

    low = 0
    high = 0
    for row in range(len(norm)):
        if norm[row][5] >= high:
            high = norm[row][5]
        if norm[row][5] <= low:
            low = norm[row][5]
    for row in range(len(norm)):
        for col in range(1, len(norm[0])):
            if col != len(norm[0])-1:
                norm[row][col] = normalisasi(float(norm[row][col]), float(10), float(0))
            else:
                norm[row][col] = normalisasi(float(norm[row][col]), float(high), float(low))
    trains = norm[:-1]
    tests = norm[-1]
    return trains, tests

def input_test():
    test = []
    test.append("sample_test")
    t = float(input("Ukuran (1-10) : "))
    test.append(t)
    t = float(input("Kenyamanan (1-10) : "))
    test.append(t)
    t = float(input("Irit (1-10) : "))
    test.append(t)
    t = float(input("Kecepatan (1-10) : "))
    test.append(t)
    t = float(input("Harga (dalam ratus juta) : "))
    test.append(t)
    return test

"""## Calculate Distance

### Euclidean distance
"""

def euclidean(x1, x2):
    a = 0
    i = 1
    l = len(x2)
    while i < l:
        a += ((x1[i]-x2[i])**2)
        i += 1
    return math.sqrt(a)

"""### Manhattan distance

"""

def manhattan(x1, x2):
    a = 0
    i = 1
    l = len(x2)
    while i < l:
        a += abs(x1[i]-x2[i])
        i += 1
    return a

"""### Minkowski distance

"""

def minkowski(x1, x2, h):
  a = 0
  i = 1
  l = len(x2)
  while i < l:
      a += (abs(x1[i]-x2[i])**h)
      i += 1
  return math.pow(a, 1/h)

"""### Supremum distance"""

def supremum(x1, x2):
    a = []
    i = 1
    l = len(x2)
    while i < l:
        a.append(abs(x1[i]-x2[i]))
        i += 1
    return max(a)

def kalkulasi(train, test, ntrain, ntest):
    df = []
    for i in range(len(train)-1):
        try:
            t = {}
            t['train'] = train[i]
            t['norm'] = ntrain[i]
            t['dist'] = {}
            t['dist']['euclidean'] = euclidean(ntrain[i], ntest)
            t['dist']['manhattan'] = manhattan(ntrain[i], ntest)
            t['dist']['minkowski'] = minkowski(ntrain[i], ntest, 1.5)
            t['dist']['supremum'] = supremum(ntrain[i], ntest)
            df.append(t)
        except IndexError:
            print(i)
            return
    return df

def knn(df):
    knn = {}
    knn['eu'] = sorted(df, key = lambda i: (i['dist']['euclidean'], i['train'][0]))
    knn['ma'] = sorted(df, key = lambda i: (i['dist']['manhattan'], i['train'][0]))
    knn['mi'] = sorted(df, key = lambda i: (i['dist']['minkowski'], i['train'][0]))
    knn['su'] = sorted(df, key = lambda i: (i['dist']['supremum'], i['train'][0]))

    return knn['eu'][:3], knn['ma'][:3], knn['mi'][:3], knn['su'][:3]

def main():
    train = baca_data('mobil.xls')
    test = input_test()
    ntrain, ntest = prapemrosesan(deepcopy(train), deepcopy(test))
    df = kalkulasi(deepcopy(train), deepcopy(test), deepcopy(ntrain), deepcopy(ntest))
    eu, ma, mi, su = knn(deepcopy(df))
    wb = Workbook()
    sheet1 = wb.add_sheet('rekomendasi')
    sheet1.write(0, 0, eu[0]['norm'][0])
    sheet1.write(1, 0, eu[1]['norm'][0])
    sheet1.write(2, 0, eu[2]['norm'][0])
    wb.save('rekomendasi.xls')
    

if __name__ == "__main__": 
    main()