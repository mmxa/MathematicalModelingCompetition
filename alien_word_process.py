import string
import copy
import time
import xlwt
import json
import matplotlib.pyplot as plt

Excel_Generation = True


def search_algorithm(word, len_word, index=0):
    #首先对文本进行遍历，得到不同的三个字母组合在文本中出现的次数
    result = dict()
    index_end = len(word) - len_word
    while index <= index_end:
        temp = word[index:index + len_word]         # 获取三个相连的字符串
        if temp in result:
            result[temp].add_index(index)
        else:
            result[temp] = Records(index)
        index += 1
    return result


class Records:
    def __init__(self, index, times=1):
        self.times = times
        self.index = [index]

    def add_index(self, index):
        self.times += 1
        self.index.append(index)


class Cost:
    def __init__(self, index,times=0):
        self.times = times
        self.index = index


def Generate_Excel(result, naming='default_name', row1=0):
    excelTable = xlwt.Workbook()
    sheet1 = excelTable.add_sheet(naming, cell_overwrite_ok=True)
    sheet1.write(0, 0, 'word')
    sheet1.write(0, 1, 'times')
    sheet1.write(0, 2, 'position')
    for k in result:
        row1 += 1
        sheet1.write(row1, 0, k)
        sheet1.write(row1, 1, result[k].times)
        sheet1.write(row1, 2, result[k].index)
    excelTable.save(naming + '.xlsx')


def calc_algorithm(data_set, word, len_word=18, index=0):
    #  通过遍历由连续的18个单词组成的字符串，得到其中三个连续子字符出现次数的函数
    result = dict()
    index_end = len(word) - len_word
    while index <= index_end:
        temp = word[index:index + len_word]  # 获取18个相连的字符串
        result[temp] = Cost(index)
        for i in range(len_word - 3):
            index_word = temp[i: i+3]
            result[temp].times += data_set[index_word].times
        index += 1
    return result


def main():
    print(__file__ + " start!!")
    start = time.time()
    f = open('name_of_words10086.txt')
    word = f.read()
    result_search = search_algorithm(word, len_word=3)
    for k in result_search:
        print("word: ", k, " times: ", result_search[k].times, " position： ", result_search[k].index)
    result_fun = calc_algorithm(result_search, word)
    for k in result_fun:
        print("word: ", k, " times: ", result_fun[k].times," position: ", result_fun[k].index)
    print('Running time: %s Seconds' % (time.time() - start))
    #if Excel_Generation is True:
        #Generate_Excel(result_search, 'words_frequency99')
        #Generate_Excel(result_fun, 'cost_of_words10086')
    total = 0
    rows = 0
    cost_data = []
    index_value = []
    for k in result_fun:                # 获取结果中各索引位置对应的cost值
        cost_data.append(result_fun[k].times)
        total += result_fun[k].times
        index_value.append(rows)
        rows += 1

    diffe = []
    normal = []
    interest_index = []
    interest_index_res = []
    lamb = 0.5
    mean_value = total / len(cost_data)
    for k in range(len(cost_data)-22):
        diffe.append(cost_data[k] - cost_data[k+21])
        norm = (cost_data[k] - cost_data[k+21]) / mean_value
        if norm > lamb:
            interest_index.append(k)
        normal.append(norm)


    #plt.scatter(index_value, cost_data, s=10, c='g')
    print(interest_index)
    for k in range(len(interest_index)-1):
        delta = interest_index[k+1] - interest_index[k]
        if delta > 10:
            if interest_index[k] not in interest_index_res:
                interest_index_res.append(int(interest_index[k]))
            if interest_index[k+1] not in interest_index_res:
                interest_index_res.append(int(interest_index[k+1]))

    print(len(interest_index))
    print(interest_index_res)
    print(len(interest_index_res))
    res_string = []
    for k in range(len(interest_index_res)-1):
        res_string.append(word[int(interest_index_res[k]-5):int(interest_index_res[k]+25)])
    with open("final_of_words.json", "w") as f:
        json.dump(res_string, f)
    print(res_string)

def    relative(str1, str2):
    for k1 in range(29):
        for k2 in range(29):
            if str1[k1] == str2[k2]:
                resutl[k1] += 1
                str3[k2] = str2[k2]
        if result[k1] >= 7:
            res[str1] = result[k1]

    plt.scatter(index_value[:len(normal)], normal, s=10, c='b')
    plt.savefig('tu1.png')
    plt.show()


if __name__ == '__main__':
    main()
