from pptx.presentation import Presentation
from pypyt import open_ppt


def test_open():
    assert isinstance(open_ppt('tests\__template__.pptx'), Presentation)
