#import sys
import os
import PyPDF2
import re
import sys
from file_finder import FileFinder

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
def main():
    feature_num = input('[1]: Merge, [2]: Rotate, [3]: Extract? ')

    if feature_num == '1':
        if os.path.exists('./merge_list.txt'):
            feature = MergeListFeature()
        else:
            feature = MergeOneByOneFeature()
    elif feature_num == '2':
        feature = RotateFeature()
    elif feature_num == '3':
        feature = ExtractFeature()
    else:
        print('No feature selected')
        return

    feature.execute()

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(e)