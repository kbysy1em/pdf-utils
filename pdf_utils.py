"""Merge, rotation and extraction of PDF files

内部的にはゼロオリジンで処理する
本のページはワンオリジンなので、入力、出力のみワンオリジンとする
入力完了後は即座に-1してゼロオリジンで内部処理する
出力時は直前で+1してワンオリジンとして表示する
"""

import commands
from collections import OrderedDict
from file_finder import FileFinder

class Option:
    def __init__(self, name, command, prep_call=None):
        self.name = name
        self.command = command
        self.prep_call = prep_call

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
    """Get file list for merge
    
    Raise:
        FileNotFoundError
    """
    files = []
    while True:
        input_file_name = input('Merge file: ')

        # 入力が空文字列ならループから抜ける
        if not input_file_name:
            break
        
        try:
            file_name = FileFinder.find(input_file_name)
        except FileNotFoundError:
            raise
        
        files.append(file_name)
    return files

def get_rotate_directions():
    """Get rotate directions
    """
    pass

def main():
    options = OrderedDict({
        '1': Option('Merge', commands.MergeOneByOneCommand(), prep_call=get_merge_files),
        '2': Option('Rotate', commands.RotateCommand()),
        '3': Option('Extract', commands.ExtractCommand()),
        '4': Option('Insert Blank Page', commands.InsertBlankCommand()),
        '5': Option('Merge by list', commands.MergeListCommand()),
    })

    print_options(options)
    chosen_option = get_option_choice(options)
    chosen_option.choose()

if __name__ == '__main__':
    print('PDF file edit utility')
    try:
        main()
    except Exception as e:
        print(e)