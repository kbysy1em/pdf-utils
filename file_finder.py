import glob

class FileFinder():
    @staticmethod
    def find(file_name):
        if file_name == '':
            files = glob.glob('*.pdf')
#            files.extend(glob.glob('*.PDF'))
        else:
            files = glob.glob(file_name)

        if len(files) == 1:
            return files[0]
        else:
            return None
