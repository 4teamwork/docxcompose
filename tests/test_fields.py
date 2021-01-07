from docxcompose.properties import FieldBase


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

    def test_can_parse_quoted_date_format(self):
        node = ' DOCPROPERTY "Propertyname" \\@ "dd-yy-mm" \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name
        assert "dd-yy-mm" == FieldForTesting(node).date_format

    def test_can_parse_quoted_date_format_with_spaces(self):
        node = ' DOCPROPERTY "Propertyname" \\@ "dd yy mm" \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name
        assert "dd yy mm" == FieldForTesting(node).date_format

    def test_can_parse_quoted_date_format_with_extra_spaces(self):
        node = ' DOCPROPERTY "Propertyname"   \\@ "dd yy mm"   \\* MERGEFORMAT'
        assert "Propertyname" == FieldForTesting(node).name
        assert "dd yy mm" == FieldForTesting(node).date_format

    def test_can_parse_unquoted_date_format(self):
        node = ' DOCPROPERTY "Propertyname" \\@ dd-yy-mm \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name
        assert "dd-yy-mm" == FieldForTesting(node).date_format

    def test_can_parse_unquoted_date_format_with_spaces(self):
        node = ' DOCPROPERTY "Propertyname" \\@ dd yy mm \\* MERGEFORMAT '
        assert "Propertyname" == FieldForTesting(node).name
        assert "dd yy mm" == FieldForTesting(node).date_format

    def test_can_parse_unquoted_date_format_with_extra_spaces(self):
        node = ' DOCPROPERTY "Propertyname"   \\@ dd yy mm   \\* MERGEFORMAT'
        assert "Propertyname" == FieldForTesting(node).name
        assert "dd yy mm" == FieldForTesting(node).date_format
