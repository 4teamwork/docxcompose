from difflib import unified_diff
from lxml import etree
from utils import ComparableDocument


def pytest_assertrepr_compare(config, op, left, right):
    if (isinstance(left, ComparableDocument)
            and isinstance(right, ComparableDocument) and op == "=="):

        extra_left = [
            item for item in right.partnames if item not in left.partnames]
        extra_right = [
            item for item in left.partnames if item not in right.partnames]
        if left.has_neq_partnames:
            explanation = ['documents contain same parts']
            if extra_left:
                explanation.append('Left contains extra parts {}'.format(
                    ', '.join(extra_left)))
            if extra_right:
                explanation.append('Right contains extra parts {}'.format(
                    ', '.join(extra_right)))
            return explanation

        diffs = []
        for lpart, rpart in left.neq_parts:
            doc = etree.fromstring(lpart.blob)
            left_xml = etree.tounicode(doc, pretty_print=True)
            doc = etree.fromstring(rpart.blob)
            right_xml = etree.tounicode(doc, pretty_print=True)

            diffs.extend(unified_diff(
                left_xml.splitlines(),
                right_xml.splitlines(),
                fromfile=lpart.partname,
                tofile=lpart.partname))

        if diffs:
            filenames = [p[0].partname for p in left.neq_parts]
            diffs.insert(
                0, 'document parts are equal. Not equal parts: {}'.format(
                    ', '.join(filenames)))
            return diffs
