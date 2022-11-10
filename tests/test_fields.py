from babel.dates import format_datetime
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
        node = ' DOCPROPERTY "Propertyname" \\@ "ddd-yy-MM" \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "E-yy-MM"

    def test_can_parse_quoted_date_format_with_spaces(self):
        node = ' DOCPROPERTY "Propertyname" \\@ "ddd yy MM" \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "E yy MM"

    def test_can_parse_quoted_date_format_with_extra_spaces(self):
        node = ' DOCPROPERTY "Propertyname"   \\@ "ddd yy MM"   \\* MERGEFORMAT'
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "E yy MM"

    def test_can_parse_unquoted_date_format(self):
        node = ' DOCPROPERTY "Propertyname" \\@ DDD-yy-MM \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "E-yy-MM"

    def test_can_parse_unquoted_date_format_with_spaces(self):
        node = ' DOCPROPERTY "Propertyname" \\@ dddd yy MMMM \\* MERGEFORMAT '
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "EEEE yy MMMM"

    def test_can_parse_unquoted_date_format_with_extra_spaces(self):
        node = ' DOCPROPERTY "Propertyname"   \\@ dd yy MM   \\* MERGEFORMAT'
        assert FieldForTesting(node).name == "Propertyname"
        assert FieldForTesting(node).date_format == "dd yy MM"


class TestFieldDateFormatMapping(object):

    def test_correctly_maps_simple_date(self):
        date = datetime(2020, 11, 19)

        assert word_to_python_date_format('yy/MM/dd') == 'yy/MM/dd'
        assert word_to_python_date_format('YY/MM/DD') == 'yy/MM/dd'
        assert format_datetime(
            date, word_to_python_date_format('yy/MM/dd')) == '20/11/19'

        assert word_to_python_date_format('yyyy/MM/dd') == 'yyyy/MM/dd'
        assert word_to_python_date_format('YYYY/MM/DD') == 'yyyy/MM/dd'
        assert format_datetime(
            date, word_to_python_date_format('YYYY/MM/DD')) == '2020/11/19'

    def test_correctly_maps_date_padding(self):
        date = datetime(2001, 2, 4)

        assert word_to_python_date_format('yy/MM/dd') == 'yy/MM/dd'
        assert format_datetime(
            date, word_to_python_date_format('yy/MM/dd')) == '01/02/04'

        assert word_to_python_date_format('yy/M/d') == 'yy/M/d'
        assert format_datetime(
            date, word_to_python_date_format('yy/M/d')) == '01/2/4'

    def test_correctly_maps_date_with_weekday_and_month_name(self):
        date = datetime(2020, 11, 19, 23, 59, 43)

        assert word_to_python_date_format('ddd dd MMM yyyy') == 'E dd MMM yyyy'
        assert format_datetime(date, word_to_python_date_format(
            'ddd dd MMM yyyy')) == 'Thu 19 Nov 2020'

        assert word_to_python_date_format('dddd dd MMMM yyyy') == 'EEEE dd MMMM yyyy'
        assert format_datetime(date, word_to_python_date_format(
            'dddd dd MMMM yyyy')) == 'Thursday 19 November 2020'

    def test_correctly_maps_date_with_time(self):
        date = datetime(2020, 11, 19, 1, 9, 8)

        assert word_to_python_date_format(
            'ddd DD MMM YYYY HH:mm:ss') == 'E dd MMM yyyy HH:mm:ss'
        assert format_datetime(date, word_to_python_date_format(
            'ddd DD MMM YYYY HH:mm:ss')) == 'Thu 19 Nov 2020 01:09:08'

        assert word_to_python_date_format(
            'ddd DD MMM YYYY H:m:s') == 'E dd MMM yyyy H:m:s'
        assert format_datetime(date, word_to_python_date_format(
            'ddd DD MMM YYYY H:m:s')) == 'Thu 19 Nov 2020 1:9:8'
