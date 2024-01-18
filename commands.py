"""Commands
"""

import re
import sys
from decimal import *
from file_finder import FileFinder
from pypdf import PdfWriter, PdfReader

def pt2mm(pt):
    return pt / Decimal(72.0) * Decimal(25.4)

class Command:
    def execute(self):
        pass

#==============================================================================
class MergeCommand(Command):
    def __init__(self):
        self.merger = PdfWriter()
        self.data = []

    def execute(self):
        pass

class MergeOneByOneCommand(MergeCommand):
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

class MergeListCommand(MergeCommand):
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
        
        print('merging: ' + self.lines[0])
        m.append(self.lines.pop(0))
        self.merge_recursively(m)

#==============================================================================
class RotateCommand(Command):
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

class ClockwiseRotation():
    def rotate(self, page):
        page.rotateClockwise(90)

class AntiClockwiseRotation():
    def rotate(self, page):
        page.rotateClockwise(-90)

#==============================================================================
class ExtractCommand(Command):
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
class InsertBlankCommand(Command):
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
