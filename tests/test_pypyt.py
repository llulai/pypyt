from collections import namedtuple
from pytest import fixture, raises

from pptx.presentation import Presentation
from pypyt import open_ppt, get_shapes_by_name, render_chart
import pandas as pd


FakePresentation = namedtuple('Presentation', 'slides')
FakeSlide = namedtuple('Slide', 'shapes')
FakeShape = namedtuple('Shape', 'name')

class Dummy:
    text_frame = None
    categories = None
    series = None



class FakeChart:
    def __init__(self):
        self.chart_data = Dummy()
        self.chart_title = Dummy()
        self.chart_title.text_frame = Dummy()
        self.chart_title.text_frame.text = None

    def __eq__(self, other):
        return (self.chart_data.categories == other.chart_data.categories
                and self.chart_data.series == other.chart_data.series
                and self.chart_title.text_frame.text == other.chart_title.text_frame.text)

    def __repr__(self):
        return f"{{categories: {self.chart_data.categories}," \
               f" series: {self.chart_data.series}," \
               f" title: {self.chart_title.text_frame.text}}}"

    def replace_data(self, chart_data):
        self.chart_data.categories = [cat.label for cat in chart_data.categories]
        self.chart_data.series = {s.name: s.values for s in chart_data._series}


@fixture
def fake_presentation():
    return FakePresentation([
        FakeSlide([FakeShape('sh1'), FakeShape('sh2')]),
        FakeSlide([FakeShape('sh2'), FakeShape('sh3')]),
    ])


@fixture
def chart_values_dict():
    return {
        'categories': ['d1', 'd2', 'd3', 'd4'],
        'data': {
            'clicks': [125, 300, 250, 200],
            'displays': [500, 450, 600, 400],
        }
    }


@fixture
def chart_values_df():
    data = [
        {'clicks': 125, 'displays': 500},
        {'clicks': 300, 'displays': 450},
        {'clicks': 250, 'displays': 600},
        {'clicks': 200, 'displays': 400},
    ]

    return pd.DataFrame(data, index=['d1', 'd2', 'd3', 'd4'])


def test_open():
    assert isinstance(open_ppt('tests/__template__.pptx'), Presentation)


def test_get_one_shape(fake_presentation):
    assert [FakeShape('sh1')] == get_shapes_by_name(fake_presentation, 'sh1')
    assert [FakeShape('sh3')] == get_shapes_by_name(fake_presentation, 'sh3')


def test_get_no_shape(fake_presentation):
    assert [] == get_shapes_by_name(fake_presentation, 'sh4')


def test_get_multiple_shapes(fake_presentation):
    assert [FakeShape('sh2'), FakeShape('sh2')] == get_shapes_by_name(fake_presentation, 'sh2')


def test_chart_same_output_no_title(chart_values_dict, chart_values_df):

    fake_chart_dict = FakeChart()
    fake_chart_df = FakeChart()

    render_chart(chart_values_df, fake_chart_dict)
    render_chart(chart_values_dict, fake_chart_df)

    assert fake_chart_dict == fake_chart_df


def test_chart_same_output_title(chart_values_dict, chart_values_df):

    chart_values_dict['title'] = 'titulo'
    chart_values_df.title = 'titulo'

    fake_chart_dict = FakeChart()
    fake_chart_df = FakeChart()

    render_chart(chart_values_df, fake_chart_dict)
    render_chart(chart_values_dict, fake_chart_df)

    assert fake_chart_dict == fake_chart_df


def test_chart_invalid_values():
    chart_invalid_data  = [1, 2, 3]

    with raises(NotImplementedError):
        render_chart(chart_invalid_data, FakeChart())
