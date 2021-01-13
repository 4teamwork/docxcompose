from datetime import datetime
from docxcompose.properties import FieldBase
from docxcompose.utils import word_to_python_date_format


class FieldForTesting(FieldBase):

    def _get_fieldname_string(self):
        return self.node


class TestFieldNameParsing(object):

    def test_can_parse_quoted_property_names(self):
        node = ' DOCPROPERTY "Propertyname"  \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"

    def test_can_parse_unquoted_property_names(self):
        node = ' DOCPROPERTY Propertyname  \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"

    def test_can_parse_quoted_property_names_with_spaces(self):
        node = ' DOCPROPERTY "Text Property"  \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Text Property"

    def test_can_parse_unquoted_property_names_with_spaces(self):
        node = ' DOCPROPERTY Text Property  \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Text Property"

    def test_can_parse_quoted_property_names_with_extra_spaces(self):
        node = ' DOCPROPERTY  "Text Property"  \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Text Property"

    def test_can_parse_unquoted_property_names_with_extra_spaces(self):
        node = ' DOCPROPERTY  Text Property  \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Text Property"


class TestFieldDateFormatParsing(object):

    def test_can_parse_quoted_date_format(self):
        node = ' DOCPROPERTY "Propertyname" \\@ "dd-yy-mm" \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "%d-%y-%M"

    def test_can_parse_quoted_date_format_with_spaces(self):
        node = ' DOCPROPERTY "Propertyname" \\@ "dd yy mm" \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "%d %y %M"

    def test_can_parse_quoted_date_format_with_extra_spaces(self):
        node = ' DOCPROPERTY "Propertyname"   \\@ "dd yy mm"   \\* MERGEFORMAT'
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "%d %y %M"

    def test_can_parse_unquoted_date_format(self):
        node = ' DOCPROPERTY "Propertyname" \\@ dd-yy-mm \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "%d-%y-%M"

    def test_can_parse_unquoted_date_format_with_spaces(self):
        node = ' DOCPROPERTY "Propertyname" \\@ dd yy mm \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "%d %y %M"

    def test_can_parse_unquoted_date_format_with_extra_spaces(self):
        node = ' DOCPROPERTY "Propertyname"   \\@ dd yy mm   \\* MERGEFORMAT'
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "%d %y %M"


class TestFieldDateFormatMapping(object):

    def test_correctly_maps_simple_date(self):
        date = datetime(2020, 11, 19)

        assert word_to_python_date_format('yy/MM/dd') == '%y/%m/%d'
        assert word_to_python_date_format('YY/MM/DD') == '%y/%m/%d'
        assert date.strftime(
            word_to_python_date_format('yy/MM/dd')) == '20/11/19'

        assert word_to_python_date_format('yyyy/MM/dd') == '%Y/%m/%d'
        assert word_to_python_date_format('YYYY/MM/DD') == '%Y/%m/%d'
        assert date.strftime(
            word_to_python_date_format('YYYY/MM/DD')) == '2020/11/19'

    def test_correctly_maps_date_padding(self):
        date = datetime(2001, 2, 4)

        assert word_to_python_date_format('yy/MM/dd') == '%y/%m/%d'
        assert date.strftime(
            word_to_python_date_format('yy/MM/dd')) == '01/02/04'

        assert word_to_python_date_format('yy/M/d') == '%y/%-m/%-d'
        assert date.strftime(
            word_to_python_date_format('yy/M/d')) == '01/2/4'

    def test_correctly_maps_date_with_weekday_and_month_name(self):
        date = datetime(2020, 11, 19, 23, 59, 43)

        assert word_to_python_date_format('ddd dd MMM yyyy') == '%a %d %b %Y'
        assert date.strftime(
            word_to_python_date_format('ddd dd MMM yyyy')) == 'Thu 19 Nov 2020'

        assert word_to_python_date_format('dddd dd MMMM yyyy') == '%A %d %B %Y'
        assert date.strftime(word_to_python_date_format(
            'dddd dd MMMM yyyy')) == 'Thursday 19 November 2020'

    def test_correctly_maps_date_with_time(self):
        date = datetime(2020, 11, 19, 1, 9, 8)

        assert word_to_python_date_format(
            'ddd DD MMM YYYY HH:mm:ss') == '%a %d %b %Y %H:%M:%S'
        assert date.strftime(word_to_python_date_format(
            'ddd DD MMM YYYY HH:mm:ss')) == 'Thu 19 Nov 2020 01:09:08'

        assert word_to_python_date_format(
            'ddd DD MMM YYYY H:m:s') == '%a %d %b %Y %-H:%-M:%-S'
        assert date.strftime(word_to_python_date_format(
            'ddd DD MMM YYYY H:m:s')) == 'Thu 19 Nov 2020 1:9:8'
