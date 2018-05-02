from collections import namedtuple
from pytest import fixture

from pptx.presentation import Presentation
from pypyt import open_ppt, get_shapes_by_name


FakePresentation = namedtuple('Presentation', 'slides')
FakeSlide = namedtuple('Slide', 'shapes')
FakeShape = namedtuple('Shape', 'name')

@fixture
def fake_presentation():
    return FakePresentation([
        FakeSlide([FakeShape('sh1'), FakeShape('sh2')]),
        FakeSlide([FakeShape('sh2'), FakeShape('sh3')]),
    ])



def test_open():
    assert isinstance(open_ppt('tests/__template__.pptx'), Presentation)


def test_get_one_shape(fake_presentation):
    assert [FakeShape('sh1')] == get_shapes_by_name(fake_presentation, 'sh1')
    assert [FakeShape('sh3')] == get_shapes_by_name(fake_presentation, 'sh3')


def test_get_no_shape(fake_presentation):
    assert [] == get_shapes_by_name(fake_presentation, 'sh4')


def test_get_multiple_shapes(fake_presentation):
    assert [FakeShape('sh2'), FakeShape('sh2')] == get_shapes_by_name(fake_presentation, 'sh2')

