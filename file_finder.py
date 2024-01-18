import glob

class FileFinder():
    """
    """
    @staticmethod
    def find(file_name):
        """
        Exception:
            FileNotFoundError
        """
        if file_name == '':
            files = glob.glob('*.pdf')
#            files.extend(glob.glob('*.PDF'))
        else:
            files = glob.glob(file_name)

        if len(files) == 1:
            return files[0]
        elif len(files) == 0:
            raise FileNotFoundError()
        else:
            raise FileNotFoundError('Multiple files detected. Unable to identify file.')
