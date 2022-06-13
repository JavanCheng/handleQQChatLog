import csv

filename = './1-12.xlsx'
a = []
new = [[]]
m = 0
n = 0

with open(filename) as f1:
    reader = csv.reader(f1)
    lists = list(reader)
    i = 0
    while i < len(lists):
        if lists[i] == new[len(new) - 1]:
            i += 1
            continue
        new.append(lists[i])
        i += 1
    new.pop(0)
    print(new)
    for i in range(0, len(new)):
        print(i, new[i])
    # for i in range(0, len(lists)):
    #     j = 1
    #     if len(lists[i]):
    #         while lists[i + j]:
    #             j += 1
    #     b = [lists[i]]
    #     for k in range(1, j):
    #         if i + k < len(lists):
    #             b.extend(lists[i + k])
    #         else:
    #             b.extend(lists[i])
    #
    #     a.append(b)
    # lists = new
# while m < len(new):
#     while len(new[m + n]):
#         n += 1
#     b = []
#     # print('n', n)
#     for k in range(m, m + n):
#         # print('k', k)
#         # print('1:', new[k])
#         b.append(new[k])
#     print(b[1:])
#     m += n + 1
    # print('m', m)

    n = 0
# print(a)
    # print(list(reader))
    # for i in lists:
    #     print(i)
    #     for j in i:
    #         j.split('?')
    # print(list(reader))

    # with open('exam.csv', 'w', newline='') as f:
    #     writer = csv.writer(f);
    #     for row in new:
    #         print(row)
    #         # if len(row) > 0:
    #         newRow = row[0].split('?')
    #         writer.writerow(newRow)
