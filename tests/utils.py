import os.path


def docx_path(filename):
    return os.path.join(os.path.dirname(__file__), 'docs', filename)
