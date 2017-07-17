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
    calculatetimes = 0
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
      	print answer + '\n\n' + str(len(self.best_wordlist)) + ' out of ' + str(wordlist_length)
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
            row = random.randrange(0, self.rows - len(word[0]))
            col = random.randrange(0, self.cols)
        else:
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
    cross_word = Crossword(row, col, '-', word_list)
#答案排布计算
    result_num = cross_word.compute_crossword(1.00, 0)
    if result_num == len(word_list):
        result = cross_word.result
        return result
    else:
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

