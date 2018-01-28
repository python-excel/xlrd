import os


def from_this_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
