#-*- coding: UTF-8 -*-  
"""Calculate the crossword and export image and text files."""

# Authors: David Whitlock <alovedalongthe@gmail.com>, Bryan Helmig
# Crossword generator that outputs the grid and clues as a pdf file and/or
# the grid in png/svg format with a text file containing the words and clues.
# Copyright (C) 2010-2011 Bryan Helmig
# Copyright (C) 2011-2016 David Whitlock
#
# Genxword is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# Genxword is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with genxword.  If not, see <http://www.gnu.org/licenses/gpl.html>.

import random, re, time, string
from operator import itemgetter
from collections import defaultdict
import json
from copy import copy as duplicate
from copy import deepcopy as deepduplicate

import sys
import math
import time
import json
reload(sys)
sys.setdefaultencoding('utf8')
import xlwt #处理excel
import xlrd
import itertools 
from xlutils.copy import copy;

PY2 = sys.version_info[0] == 2
if PY2:
    import codecs
    from functools import partial
    open = partial(codecs.open, encoding='utf-8')

class Crossword(object):
    calculatetimes = 0
    def __init__(self, rows, cols, empty=' ', available_words=[], add_world_list=[]):
        self.rows = rows
        self.cols = cols
        self.empty = empty
        self.available_words = available_words
        self.add_world_list = split_words(add_world_list)
        self.let_coords = defaultdict(list)
        self.result = {}
        self.min_col = 15
        self.max_col = 0
        self.min_row = 15
        self.max_row = 0

    def prep_grid_words(self):
        self.current_wordlist = []
        self.let_coords.clear()
        self.grid = [[self.empty]*self.cols for i in range(self.rows)]
        self.available_words = [word[:2] for word in self.available_words]
        return self.first_word(self.available_words[0])
    
    def make_min_max(self, row, col):
        if col < self.min_col:
            self.min_col = col
        elif col > self.max_col:
            self.max_col = col
        if row < self.min_row:
            self.min_row = row
        elif row > self.max_row:
            self.max_row = row
        
#func to calculate the answer 
    def compute_crossword(self, time_permitted=1.00, add_size=0):
        self.best_wordlist = []
        #answer单词列表长度
        wordlist_length = len(self.available_words)
        time_permitted = float(time_permitted)
        start_full = float(time.time())
        while (float(time.time()) - start_full) < time_permitted:
            self.result.clear()
            #只有在第一个单词无法放置的时候才会返回0
            if self.prep_grid_words() == 0:
                break
            [self.add_words(word) for i in range(2) for word in self.available_words
             if word not in self.current_wordlist]
            if len(self.current_wordlist) > len(self.best_wordlist):
                self.best_wordlist = list(self.current_wordlist)
                self.best_grid = list(self.grid)
            if len(self.best_wordlist) == wordlist_length:
                self.min_col = 15
                self.max_col = 0
                self.min_row = 15
                self.max_row = 0
                '''
                print self.best_wordlist
                for best_word in self.best_wordlist:
                    word_dict = {}
                    if len(best_word) == 4:
                        word_dict["answer"] = str(best_word[0])
                        word_dict["firstLetterRow"] = best_word[1] + add_size
                        word_dict["firstLetterCol"] = best_word[2] + add_size
                        vertical = best_word[3]
                        self.make_min_max(best_word[1] + add_size, best_word[2] + add_size)
                    else:
                        word_dict["answer"] = str(best_word[0])
                        word_dict["firstLetterRow"] = best_word[2] + add_size
                        word_dict["firstLetterCol"] = best_word[3] + add_size
                        vertical = best_word[4] 
                        self.make_min_max(best_word[2] + add_size, best_word[3] + add_size)
                    if vertical:
                        ishorizontal = 0 #col
                    else:
                        ishorizontal = 1 #row
                    word_dict["isHorizontal"] = ishorizontal
                    if self.result.has_key('data'):
                        data_list = self.result['data']
                    else:
                        data_list = []
                    data_list.append(word_dict)
                    self.result["data"] = data_list
                '''
                break
        
        if len(self.best_wordlist) == wordlist_length:
            #找到答案后best_grid输出
            #answer = '\n'.join([''.join([u'{} '.format(c) for c in self.best_grid[r]])
            #                for r in range(len(self.best_grid))])
            #print answer
            self.best_grid_buffer = deepduplicate(self.best_grid)
            self.add_world_list = sort_add_word(self.add_world_list)
            #print self.add_world_list
            i = 0
            alone_word_result = []
            while i < len(self.add_world_list):
                alone_add_data = []
                #结果数据插入单词
                alone_add_data.append(self.add_world_list[i][0])

                wordindex = 0
                resultlist = calculate_ilegalrect(self.best_grid_buffer)
                #字母矩阵输出
                #answer = '\n'.join([''.join([u'{} '.format(c) for c in self.best_grid_buffer[r]])
                #               for r in range(len(self.best_grid_buffer))])
                #print answer
                #矩阵格式化输出
                numrect = '\n'.join([''.join([u'{} '.format(c) for c in resultlist[r]])
                            for r in range(len(resultlist))])
                #print numrect
                is_insert_vert = judge_insert_type(resultlist)
                #结果数据插入插入类型
                alone_add_data.append(int(not is_insert_vert))

                is_insert_first = judge_insert_direction(is_insert_vert, resultlist)

                insert_pos =  self.insert_aloneword(self.add_world_list[i][0], is_insert_vert, is_insert_first, resultlist)

                if insert_pos != None:
                    #结果数据插入首字母坐标
                    alone_add_data.append(insert_pos[0][0])
                    alone_add_data.append(insert_pos[0][1])
                    for coord in insert_pos:
                        self.best_grid_buffer[coord[0]][coord[1]] = self.add_world_list[i][0][wordindex]
                        wordindex += 1
                    i+=1
                    alone_word_result.append(alone_add_data)
                else:
                    self.best_grid_buffer = []
                    self.best_grid = self.extend_best_grid()
                    self.best_grid_buffer = deepduplicate(self.best_grid)
                    alone_add_data = []
                    alone_word_result = []
                    i = 0

            self.min_col = 15
            self.max_col = 0
            self.min_row = 15
            self.max_row = 0
            for best_word in self.best_wordlist:
                word_dict = {}
                if len(best_word) == 4:
                    word_dict["answer"] = str(best_word[0])
                    word_dict["firstLetterRow"] = best_word[1] + add_size
                    word_dict["firstLetterCol"] = best_word[2] + add_size
                    vertical = best_word[3]
                    self.make_min_max(best_word[1] + add_size, best_word[2] + add_size)
                else:
                    word_dict["answer"] = str(best_word[0])
                    word_dict["firstLetterRow"] = best_word[2] + add_size
                    word_dict["firstLetterCol"] = best_word[3] + add_size
                    vertical = best_word[4] 
                    self.make_min_max(best_word[2] + add_size, best_word[3] + add_size)
                if vertical:
                    ishorizontal = 0 #col
                else:
                    ishorizontal = 1 #row
                word_dict["isHorizontal"] = ishorizontal
                if self.result.has_key('data'):
                    data_list = self.result['data']
                else:
                    data_list = []
                data_list.append(word_dict)
                self.result["data"] = data_list

            #独立词插入data数据
            for alone_word in alone_word_result:
                alone_word_dict = {}
                alone_word_dict["answer"] = str(alone_word[0])
                alone_word_dict["firstLetterRow"] = alone_word[2]
                alone_word_dict["firstLetterCol"] = alone_word[3]
                alone_word_dict["isHorizontal"] = alone_word[1]
                if self.result.has_key('data'):
                    data_list = self.result['data']
                else:
                    data_list = []
                data_list.append(alone_word_dict)
                self.result["data"] = data_list

        #print numrect
      	 #print answer + '\n\n' + str(len(self.best_wordlist)) + ' out of ' + str(wordlist_length)
            #最终答案输出
            answer = '\n'.join([''.join([u'{} '.format(c) for c in self.best_grid_buffer[r]])
                                for r in range(len(self.best_grid_buffer))])
            print answer
            self.result["size"] = self.cols

            return len(self.best_wordlist)

    def check_best_pos(self):
        listlen = len(self.best_grid)
        isright = True
        hasaletter = False

        for i in range(listlen):
            if self.best_grid[0][i] != '-' or self.best_grid[listlen-1][i] != '-':
                hasaletter = True

        if not hasaletter:
            print "need move ho"
            for i in range(listlen-1):
                for j in range(listlen):
                    self.best_grid[i][j] = self.best_grid[i+1][j]

            for i in range(listlen):
                self.best_grid[listlen-1][i] = '-'

        hasaletter = False

        for i in range(listlen):
            if self.best_grid[i][0] != '-' or self.best_grid[i][listlen-1] != '-':
                hasaletter = True

        if not hasaletter:
            print "need move ve"
            for i in range(listlen):
                for j in range(listlen-1):
                    self.best_grid[i][j] = self.best_grid[i][j+1]

            for i in range(listlen):
                self.best_grid[i][listlen-1] = '-'

    def extend_best_grid(self):
        self.cols = len(self.best_grid) + 1
        self.rows = self.cols
        resultlist = calculate_ilegalrect(self.best_grid)
        rectlen = len(self.best_grid)
        new_grid = [['-' for col in range(rectlen+1)] for row in range(rectlen+1)]
        mostid = most_letter_id(resultlist)
        best_word_len = range(len(self.best_wordlist))

        if mostid != 0:
            if mostid == 1:
                for i in range(rectlen):
                    for j in range(rectlen):
                        new_grid[i+1][j] = self.best_grid[i][j]
                for i in best_word_len:
                    if len(self.best_wordlist[i]) == 4:
                        self.best_wordlist[i][1] += 1
                    else:
                        self.best_wordlist[i][2] += 1
            if mostid == 2:
                for i in range(rectlen):
                    for j in range(rectlen):
                        new_grid[i+1][j+1] = self.best_grid[i][j]
                for i in best_word_len:
                    if len(self.best_wordlist[i]) == 4:
                        self.best_wordlist[i][1] += 1
                        self.best_wordlist[i][2] += 1
                    else:
                        self.best_wordlist[i][2] += 1
                        self.best_wordlist[i][3] += 1
            if mostid == 3:
                for i in range(rectlen):
                    for j in range(rectlen):
                        new_grid[i][j+1] = self.best_grid[i][j]
                for i in best_word_len:
                    if len(self.best_wordlist[i]) == 4:
                        self.best_wordlist[i][2] += 1
                    else:
                        self.best_wordlist[i][3] += 1
            if mostid == 4:
                for i in range(rectlen):
                    for j in range(rectlen):
                        new_grid[i][j] = self.best_grid[i][j]
        return new_grid
    #插入独立词语

    def insert_aloneword(self, word, is_insert_vert, is_insert_first, numlist = []):
        wordlen = len(word)
        if is_insert_first:
            for i in range(len(numlist)):
                for j in range(len(numlist[i])):
                    if numlist[i][j] == 1:
                        if is_insert_vert:
                            Vresultval = self.find_pos_vert(i, j, 1, wordlen, numlist)
                            if Vresultval != None:
                                return Vresultval
                        else:
                            Hresultval = self.find_pos_horiz(i, j, 1, wordlen, numlist)
                            if Hresultval != None:
                                return Hresultval
        else:
            numlist_length = len(numlist)
            for i in range(numlist_length):
                for j in range(numlist_length):
                    if numlist[numlist_length-i-1][numlist_length-j-1] == 1:
                        if is_insert_vert:
                            Vresultval = self.find_pos_vert(numlist_length-i-1, numlist_length-j-1, 1, wordlen, numlist)
                            if Vresultval != None:
                                return Vresultval
                        else:
                            Hresultval = self.find_pos_horiz(numlist_length-i-1, numlist_length-j-1, 1, wordlen, numlist)
                            if Hresultval != None:
                                return Hresultval

    def find_pos_horiz(self, i, j, count, wordlen, numlist = []):
        if j >= len(numlist) or numlist[i][j] == 0:
            return

        if count == wordlen:
            return [[i, j]]
        else:
            resultval = self.find_pos_horiz(i, j+1, count+1, wordlen, numlist)
            if resultval != None:
                resList = []
                resList.append([i, j])
                resList.extend(resultval)
                return resList

    def find_pos_vert(self, i, j, count, wordlen, numlist = []):
        if i >= len(numlist) or numlist[i][j] == 0:
            return

        if count == wordlen:
            return [[i, j]]
        else:
           resultval = self.find_pos_vert(i+1, j, count+1, wordlen, numlist)
           if resultval != None:
                resList = []
                resList.append([i, j])
                resList.extend(resultval)
                return resList

    def get_coords(self, word):
        """Return possible coordinates for each letter."""
        word_length = len(word[0])
        coordlist = []
        #在这里检测出共用字母的位置 l,v是共用字母在当前单词的编号和let_coords里面的信息
        temp_list =  [(l, v) for l, letter in enumerate(word[0])
                      for k, v in self.let_coords.items() if k == letter]
        for coord in temp_list:
            letc = coord[0]
            for item in coord[1]:
                (rowc, colc, vertc) = item
                if vertc:
                    if colc - letc >= 0 and (colc - letc) + word_length <= self.cols:
                        row, col = (rowc, colc - letc)
                        score = self.check_score_horiz(word, row, col, word_length)
                        if score:
                            coordlist.append([rowc, colc - letc, 0, score])
                else:
                    if rowc - letc >= 0 and (rowc - letc) + word_length <= self.rows:
                        row, col = (rowc - letc, colc)
                        score = self.check_score_vert(word, row, col, word_length)
                        if score:
                            coordlist.append([rowc - letc, colc, 1, score])
        if coordlist:
            return max(coordlist, key=itemgetter(3))
        else:
            return

    def first_word(self, word):
        """Place the first word at a random position in the grid."""
        vertical = random.randrange(0, 2)
        if vertical:
            if self.rows <= len(word[0]):
                return 0
            row = random.randrange(0, self.rows - len(word[0]))
            col = random.randrange(0, self.cols)
        else:
            if self.cols <= len(word[0]):
                return 0
            row = random.randrange(0, self.rows)
            col = random.randrange(0, self.cols - len(word[0]))
        self.set_word(word, row, col, vertical)

    def add_words(self, word):
        """Add the rest of the words to the grid."""
        coordlist = self.get_coords(word)
        if not coordlist:
            return
        row, col, vertical = coordlist[0], coordlist[1], coordlist[2]
        self.set_word(word, row, col, vertical)

    def check_score_horiz(self, word, row, col, word_length, score=1):
        cell_occupied = self.cell_occupied
        if col and cell_occupied(row, col-1) or col + word_length != self.cols and cell_occupied(row, col + word_length):
            return 0
        for letter in word[0]:
            active_cell = self.grid[row][col]
            if active_cell == self.empty:
                if row + 1 != self.rows and cell_occupied(row+1, col) or row and cell_occupied(row-1, col):
                    return 0
            elif active_cell == letter:
                score += 1
            else:
                return 0
            col += 1
        return score

    def check_score_vert(self, word, row, col, word_length, score=1):
        cell_occupied = self.cell_occupied
        if row and cell_occupied(row-1, col) or row + word_length != self.rows and cell_occupied(row + word_length, col):
            return 0
        for letter in word[0]:
            active_cell = self.grid[row][col]
            if active_cell == self.empty:
                if col + 1 != self.cols and cell_occupied(row, col+1) or col and cell_occupied(row, col-1):
                    return 0
            elif active_cell == letter:
                score += 1
            else:
                return 0
            row += 1
        return score

    def set_word(self, word, row, col, vertical):
        """Put words on the grid and add them to the word list."""
        #在单词后添加附带信息
        word.extend([row, col, vertical])
        self.current_wordlist.append(word)
        horizontal = not vertical
        for letter in word[0]:
            self.grid[row][col] = letter
            #空位则添加进去，已经存在的为相交点，从let_coords中扣除
            if (row, col, horizontal) not in self.let_coords[letter]:
                self.let_coords[letter].append((row, col, vertical))
            else:
                self.let_coords[letter].remove((row, col, horizontal))
            if vertical:
                row += 1
            else:
                col += 1
        #key = [i for i in range(2) for j in range(3) for k in range(4)]
        #print key

    def cell_occupied(self, row, col):
        cell = self.grid[row][col]
        if cell == self.empty:
            return False
        else:
            return True

def sort_add_word(lists = []):
    # 冒泡排序
    count = len(lists)
    for i in range(0, count):
        for j in range(i + 1, count):
            if len(lists[i][0]) < len(lists[j][0]):
                lists[i], lists[j] = lists[j], lists[i]
    return lists

def most_letter_id(numlist = []):
    numlist_length = len(numlist)
    first_board = math.floor(numlist_length/2.0-1)
    second_board = math.ceil(numlist_length/2.0)
    firstpart_num = 0
    secondpart_num = 0
    thirdpart_num = 0
    fourthpart_num = 0
    for i in range(numlist_length):
        for j in range(numlist_length):
            if i <= first_board and j >=second_board:
                firstpart_num += numlist[i][j]
            if i <= first_board and j <= first_board:
                secondpart_num += numlist[i][j]
            if i >= second_board and j <= first_board:
                thirdpart_num += numlist[i][j]
            if i >= second_board and j >= second_board:
                fourthpart_num += numlist[i][j]

    maxnum = max(firstpart_num, secondpart_num, thirdpart_num, fourthpart_num)
    if maxnum == firstpart_num:
        return 1
    if maxnum == secondpart_num:
        return 2
    if maxnum == thirdpart_num:
        return 3
    if maxnum == fourthpart_num:
        return 4
    return 0



def judge_insert_direction(is_insert_vert, numlist = []):
    numlist_length = len(numlist)
    first_board = math.floor(numlist_length/2.0-1)
    second_board = math.ceil(numlist_length/2.0)
    firstpart_num = 0
    secondpart_num = 0
    if is_insert_vert:
        for i in range(numlist_length):
            for j in range(numlist_length):
                if j <= first_board:
                    firstpart_num += numlist[i][j]
                if j >= second_board:
                    secondpart_num += numlist[i][j]
    else:
        for i in range(numlist_length):
            for j in range(numlist_length):
                if i <= first_board:
                    firstpart_num += numlist[i][j]
                if i >= first_board:
                    secondpart_num += numlist[i][j]

    if firstpart_num > secondpart_num:
        return 1
    else:
        return 0

def judge_insert_type(numlist = []):
    numlist_length = len(numlist)
    max_row = 0
    max_col = 0
    row_count = 0
    col_count = 0
    for i in range(numlist_length):
        for j in range(numlist_length):
            row_count += numlist[i][j]
            col_count += numlist[j][i]
        if row_count > max_row:
            max_row = row_count
        if col_count > max_col:
            max_col = col_count
        row_count = 0
        col_count = 0

    if max_row > max_col:
        return 0
    else:
        return 1

#计算01矩阵，传入best_grid
def calculate_ilegalrect(middata = []):

    rectlen = len(middata)
    numlist = [[0 for col in range(rectlen)] for row in range(rectlen)]
        
    for i in range(len(middata)):
        for j in range(len(middata[i])):
            if middata[i][j] == '-':
                numlist[i][j] = 1
            else:
                numlist[i][j] = 0

    backlist = [[1 for col in range(len(numlist) + 2)] for row in range(len(numlist) + 2)]
    resultlist = [[0 for col in range(len(numlist))] for row in range(len(numlist))]

    for i in range(len(numlist)):
        for j in range(len(numlist[i])):
            backlist[i+1][j+1] = numlist[i][j]

    #print backlist
    for i in range(len(backlist)):
        for j in range(len(backlist[i])):
            if i == 0 or i == len(backlist)-1 or j == 0 or j == len(backlist)-1:
                continue
            else:
                if (backlist[i-1][j-1] == 0 or backlist[i-1][j] == 0 or backlist[i-1][j+1] == 0 or backlist[i][j-1] == 0 or backlist[i][j+1] == 0 or backlist[i+1][j-1] == 0 or backlist[i+1][j] == 0 or backlist[i+1][j+1] == 0) and backlist[i][j] == 1:
                    backlist[i][j] = 2

    for i in range(len(backlist)):
        for j in range(len(backlist[i])):
            if i == 0 or i == len(backlist)-1 or j == 0 or j == len(backlist)-1:
                continue
            else:
                resultlist[i-1][j-1] = backlist[i][j]
                if resultlist[i-1][j-1] == 2:
                    resultlist[i-1][j-1] = 0

    return resultlist

def split_words(word_list=[]):
    cross_word_list = []
    new_list = word_list.split('|')
    for word in new_list:
        if word != "":
            cross_word_list.append([word])
    return cross_word_list

def run_word(col=4, row=4, word_list=[], add_world_list=[], num=0):
    num += 1
    if num >= 20:
        return
#初始化crossword对象 初始化的时候只有answer单词列表
    cross_word = Crossword(row, col, '-', word_list, add_world_list)
#答案排布计算
    result_num = cross_word.compute_crossword(1.00, 0)
    if result_num == len(word_list):
        result = cross_word.result
        return result
    else:
        return run_word(col+1, row+1, word_list, add_world_list, num)

#as main load in xls & print out result
def load_word(sheet_name, col_num, write_col, sheet_index, alone_col, startrect):
    # load in xls by ggod
    #workbook = xlrd.open_workbook(r'C:/Users/Administrator/Desktop/wordxls/word4.xls') 
    workbook = xlrd.open_workbook(r'/Users/ggod/Desktop/wordxls/WORDTOEXL.xls')
    #workbook = xlrd.open_workbook(r'/Users/ggod/Desktop/wordxls/word4.xls')
    sheet = workbook.sheet_by_name(sheet_name)
    data = sheet.col_values(col_num) 
    add_data = sheet.col_values(alone_col)
    copyworkbook = copy(workbook)
    write_sheet = copyworkbook.get_sheet(sheet_index)
    row = 0
    for word_list in data:
        if word_list and word_list != 'Answers':
            cross_word_list = []
            new_list = word_list.split('|')
            for word in new_list:
                if word != "":
                    cross_word_list.append([word])
            result = run_word(startrect, startrect, cross_word_list, add_data[row])
            # result = run_word(3, 3, cross_word_list, [])
            json_result = json.dumps(result)
            write_sheet.write(row, write_col, json_result)
        row += 1
    #copyworkbook.save(r'C:/Users/Administrator/Desktop/wordxls/word4.xls')
    copyworkbook.save(r'/Users/ggod/Desktop/wordxls/WORDTOEXL.xls')
    #copyworkbook.save(r'/Users/ggod/Desktop/wordxls/word4.xls')

starttime = time.time()
load_word('Word_0', 4, 7, 0, 6, 3)
load_word('Word_1', 4, 7, 1, 6, 5)
load_word('Word_2', 4, 7, 2, 6, 5)
load_word('Word_3', 4, 7, 3, 6, 5)
load_word('Word_4', 4, 7, 4, 6, 5)
load_word('Word_5', 4, 7, 5, 6, 5)
#load_word('ErrorList', 4, 7, 0, 6, 5)

print "finish in " + str(time.time() - starttime)

