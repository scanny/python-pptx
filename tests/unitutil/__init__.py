# encoding: utf-8

"""
Helper objects for unit testing.
"""


def count(start=0, step=1):
    """
    Local implementation of `itertools.count()` to allow v2.6 compatibility.
    """
    n = start
    while True:
        yield n
        n += step
