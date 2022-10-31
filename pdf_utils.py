#TODO: 空白のページを追加できるようにする
# ページサイズは手入力とする
# その際、前後のページのページサイズを参考として表示する

# 基本、内部的にはゼロオリジンで処理する
# 本のページはワンオリジンなので、入力、出力のみワンオリジンとする
# 入力完了後は即座に-1してゼロオリジンで内部処理する
# 出力時は直前で+1してワンオリジンとして表示する

#import sys
from decimal import *
import os
import PyPDF2
import re
import sys
from file_finder import FileFinder

def pt2mm(pt):
    return pt / Decimal(72.0) * Decimal(25.4)

class Feature:
    def execute(self):
        pass

#==============================================================================
class MergeFeature(Feature):
    def __init__(self):
        self.merger = PyPDF2.PdfFileMerger()

    def execute(self):
        pass

class MergeOneByOneFeature(MergeFeature):
    def execute(self):
        print('Input filename, enter blank line to finish')
        self.merge_recursively(self.merger)

        # 挿入の場合はこちら(挿入したいページを0始まりで入力)
        #merger.merge(2, '')

        self.merger.write('merged.pdf')
        self.merger.close()

    def merge_recursively(self, m):
        file_name = input('File name? ')
        if file_name == '':
            return

        m.append(file_name)
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

        file = PyPDF2.PdfFileReader(open(file_name, 'rb'))

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

        self.output = PyPDF2.PdfFileWriter()
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
            sys.exit('ファイル名が不正です')

        merger = PyPDF2.PdfFileMerger()

        print('If you want pages of 4th ~ 6th, input 4 and 6')
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

        reader = PyPDF2.PdfReader(file_name)

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

        writer = PyPDF2.PdfWriter()
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
def main():
    feature_num = input('[1]: Merge, [2]: Rotate, [3]: Extract, [4]: InsertBlankPage? ')

    if feature_num == '1':
        if os.path.exists('./merge_list.txt'):
            feature = MergeListFeature()
        else:
            feature = MergeOneByOneFeature()
    elif feature_num == '2':
        feature = RotateFeature()
    elif feature_num == '3':
        feature = ExtractFeature()
    elif feature_num == '4':
        feature = InsertBlankFeature()
    else:
        print('No feature selected')
        return

    feature.execute()

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(e)