import os

def search_file (filename, dirpath=""):
    try:
        if dirpath == "":
            dirpath = os.path.dirname(__file__)
        for (path, dir, files) in os.walk(dirpath):
            for f in files:
                if f == filename:
                    return os.path.join(path, f)
    except PermissionError:
        pass

def get_numeric_pos (str):
    for i, s in enumerate (str):
        if s.isnumeric() == True :
            return i

def exportcsv (wbname):
    try:
        if dirpath == "":
            dirpath = os.path.dirname(__file__)
        for (path, dir, files) in os.walk(dirpath):
            for f in files:
                if f == filename:
                    return os.path.join(path, f)
    except PermissionError:
        pass
   
