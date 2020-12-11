import os


def from_sample(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'samples', filename)
