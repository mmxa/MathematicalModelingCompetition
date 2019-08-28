import string
import random
import copy


class random_txt_generation():

    def __init__(self):
        self.index_used = []                # 插入位置记录到列表中，避免重复
        self.saved_words = []               # 原始字符串记录
        self.final_words = []               # 记录替换后的15-21个字母组成的字符串
        self.final_words_pos = []           # 记录替换字符的位置
        self.final_txt = []                 # 记录替换后的文本字符串

    def random_str(self, length):
        rand_word = []
        rand_word += [random.choice('abcdefghijklmnopqrst') for i in range(length)]
        words = "".join(rand_word)          # 根据给定的字符串长度生成一个由字母a-t组成的字符串
        return words

    def alternate(self, orig):              # 对初始字符串进行随机替换
        used_index = []
        k = 0
        orig_str = copy.deepcopy(orig)
        num_alt = random.randint(0, 4)      # 从0到4之间随机选取替换字母次数
        if num_alt == 0:                    # 若替换次数为0，则返回原字符串
            self.final_words.append(orig_str)
            return orig_str
        else:
            while k < num_alt:
                index_alt = random.randint(0, len(orig_str) - 1)
                # 随机生成字符串替换的位置
                if index_alt in used_index:
                    continue
                # 若替换位置已经使用过，则跳过并进行下一次循环，k不计数
                else:
                    used_index.append(index_alt)
                # 将随机替换位置记录，以防止再次替换
                    orig_str = orig_str[:index_alt] + random.choice('abcdefghijklmnopqrst') + orig_str[index_alt + 1:]
                # 若替换位置符合条件，则对该位置的字母进行随机替换
                    k += 1
            self.final_words.append(orig_str)
                # 记录替换后的字符串并返回
        return orig_str

    def insert_word(self, txt_ori, orig):
        orig_str = copy.deepcopy(orig)
        fig = True
        k = 0
        str_modified = txt_ori
        word_times = random.randint(40, 60)
        print(word_times)
        # 随机生成替换字符串的次数
        while k < word_times:
            index_alt = int(random.randint(0, len(txt_ori) - len(orig_str)))
            # 随机产生进行字符串替换的位置
            fig = True
            for q in self.index_used:
                if abs(q - index_alt) < len(orig_str):
                    fig = False
                #若产生的随机插入位置距离已插入字符位置距离较近，将标志位切换为False

            if fig is not False:
                str_modified = str_modified[:index_alt] + self.alternate(orig_str) + \
                               str_modified[index_alt + len(orig_str):]
                self.index_used.append(index_alt)
                k = k + 1
                self.final_words_pos.append(index_alt)
                # 若产生的替换位置满足间隔要求，则对字符串数据进行插入
        self.final_txt.append(str_modified)


def main():
    generating = random_txt_generation()
    word_num = random.randint(15, 21)
    # 生成随机字符的长度
    orig = generating.random_str(word_num)
    # 生成随机字符串，作为待提取数据的原始数据

    text_ori = generating.random_str(random.randint(150000, 180000))
    orig2 = generating.random_str(word_num)
    generating.saved_words.append(orig)
    generating.saved_words.append(orig2)
    generating.insert_word(text_ori, orig)
    generating.insert_word(text_ori, orig2)
    # 向随机文本中插入按规则进行替换的字符串数据
    # print(generating.final_txt)
    x = str(123)
    with open("name_of_words" + x + ".txt", "w") as f:
        f.write(generating.final_txt[0])
    # 将生成的文本保存到txt文件中
    print(generating.saved_words)
    print(generating.final_words)
    print(generating.final_words_pos)


if __name__ == '__main__':
    main()
