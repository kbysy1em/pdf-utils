"""Merge, rotation and extraction of PDF files

内部的にはゼロオリジンで処理する
本のページはワンオリジンなので、入力、出力のみワンオリジンとする
入力完了後は即座に-1してゼロオリジンで内部処理する
出力時は直前で+1してワンオリジンとして表示する
"""

#import sys
from decimal import *
import os
from pypdf import PdfWriter, PdfReader
import re
import sys
from collections import OrderedDict
from file_finder import FileFinder

def pt2mm(pt):
    return pt / Decimal(72.0) * Decimal(25.4)

class Feature:
    def execute(self):
        pass

#==============================================================================
class MergeFeature(Feature):
    def __init__(self):
        self.merger = PdfWriter()
        self.data = []

    def execute(self):
        pass

class MergeOneByOneFeature(MergeFeature):
    def execute(self, data):
        self.data = data
        self.merge_recursively(self.merger)

        # 挿入の場合はこちら(挿入したいページを0始まりで入力)
        #merger.merge(2, '')

        self.merger.write('merged.pdf')
        self.merger.close()

        return 'Done.'

    def merge_recursively(self, m):
        if self.data == []:
            return
        
        print('merged file: ' + self.data[0])
        m.append(self.data.pop(0))
        self.merge_recursively(m)

class MergeListFeature(MergeFeature):
    def execute(self):
        with open('./merge_list.txt', 'rt', encoding='utf-8') as f:
            self.lines = f.read().splitlines()

        self.merge_recursively(self.merger)

        # 挿入の場合はこちら(挿入したいページを0始まりで入力)
        #merger.merge(2, '')

        self.merger.write('merged.pdf')
        self.merger.close()

        return 'Done.'
        
    def merge_recursively(self, m):
        if self.lines == []:
            return
        
        print('merged file: ' + self.lines[0])
        m.append(self.lines.pop(0))
        self.merge_recursively(m)

#==============================================================================
class RotateFeature(Feature):
    def execute(self):
        input_file_name = input('File name?')
        file_name = FileFinder.find(input_file_name)

        if file_name is None:
            sys.exit('ファイル名が不正です')

        file = PdfReader(open(file_name, 'rb'))

        # pageは1はじまりとする
        print('Input pages rotated (comma separated). If blank, all pages are rotated')
        s = input()
        if not s:
            rotate_pages = range(1, file.numPages + 1)
        else:
            rotate_pages = [int(x.strip()) for x in s.split(',')]

        direction = input('Clockwise(y/n)?')
        if direction == 'n':
            rotation = AntiClockwiseRotation()
        else:
            rotation = ClockwiseRotation()

        self.output = PdfWriter()
        for i in range(file.numPages):
            page = file.getPage(i)
            if i + 1 in rotate_pages:
                rotation.rotate(page)
            self.output.addPage(page)

        d_name = re.search(r'^.*\\', file_name)
        f_name = re.search(r'[^\\]*?$', file_name)
        d_name_str = d_name.group() if d_name is not None else ''

        with open(d_name_str + 'rot-' + f_name.group(), 'wb') as f:
            self.output.write(f)

class Rotation:
    def rotate(self, page):
        pass

class ClockwiseRotation(Rotation):
    def rotate(self, page):
        page.rotateClockwise(90)

class AntiClockwiseRotation(Rotation):
    def rotate(self, page):
        page.rotateClockwise(-90)

#==============================================================================
class ExtractFeature(Feature):
    def execute(self):
        input_file_name = input('File name?')
        file_name = FileFinder.find(input_file_name)

        if file_name is None:
            sys.exit('Input file name is wrong.')

        merger = PdfWriter()

        print('If you want pages of 4th ~ 6th, input 4 and 6.')
        start = int(input('Start page? '))
        end = int(input('End page? '))
        merger.append(file_name, pages=(start - 1, end))

        # 以下ではstart=2, stop=3としている。
        # この場合、3ページ以降で4ページより前の部分が取り出され,
        # 結果的に3ページ目のみのファイルが得られる
        # merger.append(filename, pages=(2,3))

        # n = merger.getNumPages()
        # merger.append(input_file_name, pages=(0,n))

        merger.write('extracted.pdf')

        merger.close()

#==============================================================================
class InsertBlankFeature(Feature):
    def execute(self):
        input_file_name = input('File name?')
        file_name = FileFinder.find(input_file_name)

        if file_name is None:
            sys.exit('ファイル名が不正です')

        reader = PdfReader(file_name)

        # len(reader.pages)でページ数が得られる。
        # これはワンオリジンで数えた場合の最終ページのページ番号と一致する
        p_1origin = int(input(
            'Input page number bhind of which blank page is inserted. '
            f'Num of pages = {len(reader.pages)}\n'
        ))
        p = p_1origin - 1
        page_nums = [i for i in range(p - 2, p + 4) if i >= 0 and i < len(reader.pages)]

        # print page size
        for i in page_nums:
            print('page: {}; height: {}pt, {:.6f}mm; width: {}pt, {:.6f}mm'.format(
                i + 1, 
                reader.pages[i].mediaBox.height, pt2mm(reader.pages[i].mediaBox.height), 
                reader.pages[i].mediaBox.width, pt2mm(reader.pages[i].mediaBox.width)
            ))

        # input blank page size
        n_1origin = int(input('Input page number which has the desired paper size\n'))
        n = n_1origin - 1

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        
        print(f'Blank page is inserted after page {p_1origin}. Page height {reader.pages[n].mediaBox.height}pt,'
            f' width {reader.pages[n].mediaBox.width}pt')

        writer.insert_blank_page(
            width=reader.pages[n].mediaBox.width, 
            height=reader.pages[n].mediaBox.height,
            index=p + 1)

        d_name = re.search(r'^.*\\', file_name)
        f_name = re.search(r'[^\\]*?$', file_name)
        d_name_str = d_name.group() if d_name is not None else ''

        with open(d_name_str + 'bla-' + f_name.group(), 'wb') as f:
            writer.write(f)

#==============================================================================
class Option:
    def __init__(self, name, command, prep_call=None):
        self.name = name  # <1>
        self.command = command  # <2>
        self.prep_call = prep_call  # <3>

    def _handle_message(self, message):
        print(message)

    def choose(self):  # <4>
        data = self.prep_call() if self.prep_call else None  # <5>
        message = self.command.execute(data) if data else self.command.execute()  # <6>
        self._handle_message(message)

    def __str__(self):  # <7>
        return self.name

#==============================================================================

def option_choice_is_valid(choice, options):
    return choice in options or choice.upper() in options  # <1>

def get_option_choice(options):
    choice = input('操作を選択してください: ')  # <2>
    if not option_choice_is_valid(choice, options):  # <3>
        raise ValueError('Incorrect characters were entered during the function selection.')
    return options[choice.upper()]  # <4>

def print_options(options):
    for shortcut, option in options.items():
        print(f'({shortcut}) {option}')

def get_merge_files():
    files = []
    while True:
        file_name = input('Merge file: ')
        # ここでfilenameのバリデーションを行う
        if not file_name:
            break
        files.append(file_name)
    return files

def main():
    options = OrderedDict({
        '1': Option('Merge', MergeOneByOneFeature(), prep_call=get_merge_files),
        '2': Option('Rotate', RotateFeature()),
        '3': Option('Extract', ExtractFeature()),
        '4': Option('Insert Blank Page', InsertBlankFeature()),
        '5': Option('Merge by list', MergeListFeature()),
    })

    print_options(options)
    chosen_option = get_option_choice(options)
    chosen_option.choose()




    # if feature_num == '1':
    #     if os.path.exists('./merge_list.txt'):
    #         feature = MergeListFeature()
    #     else:
    #         feature = MergeOneByOneFeature()
    # elif feature_num == '2':
    #     feature = RotateFeature()
    # elif feature_num == '3':
    #     feature = ExtractFeature()
    # elif feature_num == '4':
    #     feature = InsertBlankFeature()
    # else:
    #     print('No feature selected')
    #     return

    # feature.execute()

if __name__ == '__main__':
    print('PDF file edit utility')
    try:
        main()
    except Exception as e:
        print(e)