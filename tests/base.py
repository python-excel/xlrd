import os
import shutil


def from_this_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

def from_home_dir(filename, fn):
    abs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    home_path = os.path.expanduser( os.path.join('~', filename))

    shutil.copyfile(abs_path, home_path)
    fn(home_path)
    shutil.os.remove(home_path)
