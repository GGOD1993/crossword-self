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

import sys
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
    def __init__(self, rows, cols, empty=' ', available_words=[]):
        self.rows = rows
        self.cols = cols
        self.empty = empty
        self.available_words = available_words
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
                break
        answer = '\n'.join([''.join([u'{} '.format(c) for c in self.best_grid[r]])
                            for r in range(self.rows)])
      	#print answer + '\n\n' + str(len(self.best_wordlist)) + ' out of ' + str(wordlist_length)
        self.result["size"] = self.cols + add_size
        return len(self.best_wordlist)

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
            #垂直放置，行数不够无法放置
            if self.rows <= len(word[0]):
                return 0
            #在合理范围内随机放置
            row = random.randrange(0, self.rows - len(word[0]))
            col = random.randrange(0, self.cols)
        else:
            #同上
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
        print cell_occupied
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
            if (row, col, horizontal) not in self.let_coords[letter]:
                self.let_coords[letter].append((row, col, vertical))
            else:
                self.let_coords[letter].remove((row, col, horizontal))
            if vertical:
                row += 1
            else:
                col += 1

    def cell_occupied(self, row, col):
        cell = self.grid[row][col]
        if cell == self.empty:
            return False
        else:
            return True

def split_words(word_list=[]):
    cross_word_list = []
    new_list = word_list.split('|')
    for word in new_list:
        if word != "":
            cross_word_list.append([word])
    return cross_word_list

def make_hor_alone_words(crossword, full_list, full_row, full_col, next_col):
    # full_row = full_col = 0
    if len(full_list) > 2:
        return
    for full in full_list:
        word_dict = {}
        word_dict["answer"] = str(full)
        word_dict["firstLetterRow"] = full_row
        word_dict["firstLetterCol"] = full_col
        word_dict["isHorizontal"] = 1
        full_row += (crossword.cols + 2)
        full_col += next_col
        if crossword.result["data"]:
            crossword.result["data"].append(word_dict)
    
def make_ver_alone_words(crossword, normal_list, normal_row, normal_col):
    # normal_row = 1
    # normal_col = crossword.cols + 2
    for normal in normal_list:
        word_dict = {}
        word_dict["firstLetterRow"] = normal_row
        word_dict["firstLetterCol"] = normal_col
        word_dict["isHorizontal"] = 0
        if crossword.result["data"]:
            crossword.result["data"].append(word_dict)
        normal_col = 0

def make_two_alone_words(crossword, word_list, full_row):
    # print crossword.min_row, crossword.max_row
    min_row = crossword.min_row - 2
    max_row = crossword.max_row + 2
    for word in word_list:
        word_length = len(word)
        tmp_pos = (crossword.cols + 2 - word_length) / 2
        if tmp_pos <= 1:
            first_pos = tmp_pos
        else:
            first_pos = tmp_pos - 1
        word_dict = {}
        word_dict["answer"] = str(word)
        word_dict["firstLetterRow"] = min_row
        word_dict["firstLetterCol"] = first_pos
        word_dict["isHorizontal"] = 1
        min_row = max_row
        # word_dict["firstLetterRow"] = full_row
        # word_dict["firstLetterCol"] = first_pos
        # word_dict["isHorizontal"] = 1
        # full_row += (crossword.cols + 2)
        if crossword.result["data"]:
            crossword.result["data"].append(word_dict)

def make_alone_words(crossword, word_list):
    add_list = split_words(word_list)
    aloneLength = len(add_list)
    if aloneLength > 0 and aloneLength <= 4:
        near_word = full_word = 0
        full_list = near_list = []
        normal_list = []
        all_list = []
        for word in add_list:
            for i in word:
                if len(i) == crossword.cols + 2:
                    full_list.append(i)
                    all_list.append(i)
                    full_word += 1
                elif len(i) == crossword.cols + 1:
                    near_list.append(i)
                    # normal_list.append(i)
                    all_list.append(i)
                    near_word += 1
                else:
                    all_list.append(i)
                    normal_list.append(i)
        if full_word >= 3 or near_word >= 3:
            return
        full_near_num = full_word + near_word
        if full_near_num >= 3:
            return

        if aloneLength <= 2:
            make_two_alone_words(crossword, all_list, 0)
        else:
            full_row = full_col = 0
            make_hor_alone_words(crossword, full_list, full_row, full_col, 0)

            if full_word > 0:
                normal_row = 1
                normal_col = crossword.cols + 2
                make_ver_alone_words(crossword, normal_list, normal_row, normal_col)
            else:
                if near_word > 0:
                    make_hor_alone_words(crossword, near_list, 0, 0, 1)
                    make_ver_alone_words(crossword, normal_list, 2, crossword.cols + 2)
                else:
                    if aloneLength > 2:
                        hor_list = normal_list[:2]
                        make_hor_alone_words(crossword, hor_list, 0, 1, 0)
                        ver_list = normal_list[2:aloneLength]
                        make_ver_alone_words(crossword, ver_list, 1, crossword.cols + 2)
                    else:
                        make_hor_alone_words(crossword, normal_list, 0, 1, 0)
            
def load_alone_words(sheet_name, col_num, sheet_index, alone_col):
    workbook = xlrd.open_workbook(r'D:\crossword\Words-WordCrumble(en)-2.xls') 
    sheet = workbook.sheet_by_name(sheet_name)
    data = sheet.col_values(col_num) 
    add_data = sheet.col_values(alone_col) 
    copyworkbook = copy(workbook)
    write_sheet = copyworkbook.get_sheet(sheet_index)
    row = 0
    size = 0
    for word_list in add_data:
        i = data[row]
        if i:
            excel_data = i.encode('unicode-escape').decode('string_escape') 
            json_data = json.loads(excel_data)
            size = json_data['size']
        if size:
            add_answer = only_make_alone_words(size, add_data[row])
            if json_data["data"]:
                for answer in json_data["data"]:
                    answer["firstLetterRow"] += 2
                    answer["firstLetterCol"] += 2  
            json_data["data"].append(add_answer)
            json_data["size"] += 2 
            json_result = json.dumps(json_data)
            write_sheet.write(row, col_num, json_result)
        row += 1
    copyworkbook.save(r'D:\crossword\Words-WordCrumble(en)-2.xls')

def only_make_alone_words(size, word_list):
    add_list = split_words(word_list)
    aloneLength = len(add_list)
    data = []
    if aloneLength > 0 and aloneLength <= 4:
        fisrt_pos = random.randint(1, 4)
        for word in add_list:
            for i in word:
                word_len = len(i)
                if word_len <= size + 2:
                    world_first = random.randint(0, size + 2 - word_len)
                    word_dict = {}
                    word_dict["answer"] = str(i)
                    if fisrt_pos == 1:
                        word_dict["firstLetterRow"] = 0
                        word_dict["firstLetterCol"] = world_first
                        word_dict["isHorizontal"] = 1
                    elif fisrt_pos == 2:
                        word_dict["firstLetterRow"] = world_first
                        word_dict["firstLetterCol"] = size + 2
                        word_dict["isHorizontal"] = 0
                    elif fisrt_pos == 3:
                        word_dict["firstLetterRow"] = size + 2
                        word_dict["firstLetterCol"] = world_first
                        word_dict["isHorizontal"] = 1
                    elif fisrt_pos == 4:
                        word_dict["firstLetterRow"] = world_first
                        word_dict["firstLetterCol"] = 0
                        word_dict["isHorizontal"] = 0
                    data.append(word_dict)
                    if fisrt_pos == 4:
                        fisrt_pos = 1
                    else:
                        fisrt_pos += 1
    return data
#entry of run the core alg
#word_list 答案列表 add_word_list aloneword列表

def run_word(col=4, row=4, word_list=[], add_world_list=[], num=0):
    num += 1
    if num >= 20:
        return
#check the alone wordlist
    if len(add_world_list) > 0:
        add_size = 2
    else:
        add_size = 0
#初始化crossword对象 初始化的时候只有answer单词列表
    cross_word = Crossword(row, col, '-', word_list)
#答案排布计算
    result_num = cross_word.compute_crossword(1.00, 0)
    if result_num == len(word_list):
        #填充aloneword
        #make_alone_words(cross_word, add_world_list)
        result = cross_word.result
        # print result
        return result
    else:
        # return run_word(col+1, row+1, word_list, num)
        #递归计算
        return run_word(col+1, row+1, word_list, add_world_list, num)


#as main load in xls & print out result
def load_word(sheet_name, col_num, write_col, sheet_index, alone_col):
    # load in xls by ggod
    workbook = xlrd.open_workbook(r'/Users/ggod/Desktop/wordxls/word4.xls') 
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
            result = run_word(7, 7, cross_word_list, add_data[row])
            # result = run_word(3, 3, cross_word_list, [])
            json_result = json.dumps(result)
            write_sheet.write(row, write_col, json_result)
        row += 1
    copyworkbook.save(r'/Users/ggod/Desktop/wordxls/word4.xls')

load_word('ErrorList', 4, 7, 0, 6)

