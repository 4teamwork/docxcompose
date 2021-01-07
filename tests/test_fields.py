from datetime import datetime
from docxcompose.properties import FieldBase
from docxcompose.utils import word_to_python_date_format


class FieldForTesting(FieldBase):

    def _get_fieldname_string(self):
        return self.node


class TestFieldNameParsing(object):

    def test_can_parse_quoted_property_names(self):
        node = ' DOCPROPERTY "Propertyname"  \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name

    def test_can_parse_unquoted_property_names(self):
        node = ' DOCPROPERTY Propertyname  \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name

    def test_can_parse_quoted_property_names_with_spaces(self):
        node = ' DOCPROPERTY "Text Property"  \\* MERGEFORMAT '
        assert "Text Property" == FieldForTesting(node).name

    def test_can_parse_unquoted_property_names_with_spaces(self):
        node = ' DOCPROPERTY Text Property  \\* MERGEFORMAT '
        assert "Text Property" == FieldForTesting(node).name

    def test_can_parse_quoted_property_names_with_extra_spaces(self):
        node = ' DOCPROPERTY  "Text Property"  \\* MERGEFORMAT '
        assert "Text Property" == FieldForTesting(node).name

    def test_can_parse_unquoted_property_names_with_extra_spaces(self):
        node = ' DOCPROPERTY  Text Property  \\* MERGEFORMAT '
        assert "Text Property" == FieldForTesting(node).name


class TestFieldDateFormatParsing(object):

    def test_can_parse_quoted_date_format(self):
        node = ' DOCPROPERTY "Propertyname" \\@ "dd-yy-mm" \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name
        assert "%d-%y-%M" == FieldForTesting(node).date_format

    def test_can_parse_quoted_date_format_with_spaces(self):
        node = ' DOCPROPERTY "Propertyname" \\@ "dd yy mm" \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name
        assert "%d %y %M" == FieldForTesting(node).date_format

    def test_can_parse_quoted_date_format_with_extra_spaces(self):
        node = ' DOCPROPERTY "Propertyname"   \\@ "dd yy mm"   \\* MERGEFORMAT'
        assert "Propertyname" == FieldForTesting(node).name
        assert "%d %y %M" == FieldForTesting(node).date_format

    def test_can_parse_unquoted_date_format(self):
        node = ' DOCPROPERTY "Propertyname" \\@ dd-yy-mm \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name
        assert "%d-%y-%M" == FieldForTesting(node).date_format

    def test_can_parse_unquoted_date_format_with_spaces(self):
        node = ' DOCPROPERTY "Propertyname" \\@ dd yy mm \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name
        assert "%d %y %M" == FieldForTesting(node).date_format

    def test_can_parse_unquoted_date_format_with_extra_spaces(self):
        node = ' DOCPROPERTY "Propertyname"   \\@ dd yy mm   \\* MERGEFORMAT'
        assert "Propertyname" == FieldForTesting(node).name
        assert "%d %y %M" == FieldForTesting(node).date_format


class TestFieldDateFormatMapping(object):

    def test_correctly_maps_simple_date(self):
        date = datetime(2020, 11, 19)

        assert '%y/%m/%d' == word_to_python_date_format('yy/MM/dd')
        assert '%y/%m/%d' == word_to_python_date_format('YY/MM/DD')
        assert '20/11/19' == date.strftime(
            word_to_python_date_format('yy/MM/dd'))

        assert '%Y/%m/%d' == word_to_python_date_format('yyyy/MM/dd')
        assert '%Y/%m/%d' == word_to_python_date_format('YYYY/MM/DD')
        assert '2020/11/19' == date.strftime(
            word_to_python_date_format('YYYY/MM/DD'))

    def test_correctly_maps_date_padding(self):
        date = datetime(2001, 2, 4)

        assert '%y/%m/%d' == word_to_python_date_format('yy/MM/dd')
        assert '01/02/04' == date.strftime(
            word_to_python_date_format('yy/MM/dd'))

        assert '%y/%-m/%-d' == word_to_python_date_format('yy/M/d')
        assert '01/2/4' == date.strftime(
            word_to_python_date_format('yy/M/d'))

    def test_correctly_maps_date_with_weekday_and_month_name(self):
        date = datetime(2020, 11, 19, 23, 59, 43)

        assert '%a %d %b %Y' == word_to_python_date_format('ddd dd MMM yyyy')
        assert 'Thu 19 Nov 2020' == date.strftime(
            word_to_python_date_format('ddd dd MMM yyyy'))

        assert '%A %d %B %Y' == word_to_python_date_format('dddd dd MMMM yyyy')
        assert 'Thursday 19 November 2020' == date.strftime(
            word_to_python_date_format('dddd dd MMMM yyyy'))

    def test_correctly_maps_date_with_time(self):
        date = datetime(2020, 11, 19, 1, 9, 8)

        assert '%a %d %b %Y %H:%M:%S' == word_to_python_date_format(
            'ddd DD MMM YYYY HH:mm:ss')
        assert 'Thu 19 Nov 2020 01:09:08' == date.strftime(
            word_to_python_date_format('ddd DD MMM YYYY HH:mm:ss'))

        assert '%a %d %b %Y %-H:%-M:%-S' == word_to_python_date_format(
            'ddd DD MMM YYYY H:m:s')
        assert 'Thu 19 Nov 2020 1:9:8' == date.strftime(
            word_to_python_date_format('ddd DD MMM YYYY H:m:s'))
