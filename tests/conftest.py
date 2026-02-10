import random
from difflib import unified_diff

import pytest
from lxml import etree
from utils import ComparableDocument


@pytest.fixture
def static_reseed():
    """Static random seed fixture to test content that generates random ids."""
    random.seed(42)


def pytest_assertrepr_compare(config, op, left, right):
    if (
        isinstance(left, ComparableDocument)
        and isinstance(right, ComparableDocument)
        and op == "=="
    ):

        left.post_compare_failed(right)
        right.post_compare_failed(left)

        if left.has_neq_partnames:
            extra_right = [
                item for item in right.partnames if item not in left.partnames
            ]
            extra_left = [
                item for item in left.partnames if item not in right.partnames
            ]

            explanation = ["documents contain same parts"]
            if right.doc is None:
                explanation.append("Right document is None")
            if left.doc is None:
                explanation.append("Left document is None")
            if extra_left:
                explanation.append(
                    "Left contains extra parts {}".format(", ".join(extra_left))
                )
            if extra_right:
                explanation.append(
                    "Right contains extra parts {}".format(", ".join(extra_right))
                )
            return explanation

        diffs = []
        for lpart, rpart in left.neq_parts:

            if not lpart.partname.endswith(".xml"):
                diffs.append("Binary parts differ {}".format(lpart.partname))
                diffs.append("")
                continue

            doc = etree.fromstring(lpart.blob)
            left_xml = etree.tounicode(doc, pretty_print=True)
            doc = etree.fromstring(rpart.blob)
            right_xml = etree.tounicode(doc, pretty_print=True)

            diffs.extend(
                unified_diff(
                    left_xml.splitlines(),
                    right_xml.splitlines(),
                    fromfile=lpart.partname,
                    tofile=lpart.partname,
                )
            )
            diffs.append("")

        if diffs:
            filenames = [p[0].partname for p in left.neq_parts]
            diffs.insert(
                0,
                "document parts are equal. Not equal parts: {}".format(
                    ", ".join(filenames)
                ),
            )
            return diffs
