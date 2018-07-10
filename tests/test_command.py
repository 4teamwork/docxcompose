from docxcompose import command
from utils import docx_path
import pytest


def test_command_creates_composed_docx_file_at_output_path(tmpdir):
    output_path = tmpdir.join('outfile.docx')
    assert not output_path.exists()

    arguments = [docx_path('master.docx'),
                 docx_path('table.docx'),
                 '--output-document', output_path.strpath]
    with pytest.raises(SystemExit) as exc_info:
        command.main(arguments)

    assert exc_info.value.code == 0
    assert output_path.exists()
    assert output_path.isfile()
    assert output_path.size() > 0
