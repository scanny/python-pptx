# encoding: utf-8

"""Utility functions for loading files for unit testing."""

import os
import sys


_thisdir = os.path.split(__file__)[0]
test_file_dir = os.path.abspath(os.path.join(_thisdir, "..", "test_files"))


def absjoin(*paths):
    return os.path.abspath(os.path.join(*paths))


def snippet_bytes(snippet_file_name):
    """Return bytes read from snippet file having `snippet_file_name`."""
    snippet_file_path = os.path.join(
        test_file_dir, "snippets", "%s.txt" % snippet_file_name
    )
    with open(snippet_file_path, "rb") as f:
        return f.read().strip()


def snippet_seq(name, offset=0, count=sys.maxsize):
    """
    Return a tuple containing the unicode text snippets read from the snippet
    file having *name*. Snippets are delimited by a blank line. If specified,
    *count* snippets starting at *offset* are returned.
    """
    path = os.path.join(test_file_dir, "snippets", "%s.txt" % name)
    with open(path, "rb") as f:
        text = f.read().decode("utf-8")
    snippets = text.split("\n\n")
    start, end = offset, offset + count
    return tuple(snippets[start:end])


def snippet_text(snippet_file_name):
    """
    Return the unicode text read from the test snippet file having
    *snippet_file_name*.
    """
    snippet_file_path = os.path.join(
        test_file_dir, "snippets", "%s.txt" % snippet_file_name
    )
    with open(snippet_file_path, "rb") as f:
        snippet_bytes = f.read()
    return snippet_bytes.decode("utf-8")


def testfile(name):
    """
    Return the absolute path to test file having *name*.
    """
    return absjoin(test_file_dir, name)


def testfile_bytes(*segments):
    """Return bytes of file at path formed by adding `segments` to test file dir."""
    path = os.path.join(test_file_dir, *segments)
    with open(path, "rb") as f:
        return f.read()
