"""Exceptions used with python-docx.

The base exception class is PythonDocxError.
"""


class PythonDocxError(Exception):
    """Generic error class."""


class InvalidSpanError(PythonDocxError):
    """Raised when an invalid merge region is specified in a request to merge table
    cells."""


class InvalidXmlError(PythonDocxError):
    """Raised when invalid XML is encountered, such as on attempt to access a missing
    required child element."""


class UnsupportedXPathError(PythonDocxError):
    """Raised when an XPath expression uses constructs outside the supported subset."""


class BindingRefreshError(PythonDocxError):
    """Raised when a content-control refresh plan cannot be applied.

    The .problems attribute is a tuple of |BindingProblem| objects, each identifying
    the content control that could not be refreshed and the reason.
    """

    def __init__(self, problems):
        self.problems = tuple(problems)
        super(BindingRefreshError, self).__init__(str(self))

    def __str__(self):
        lines = ["content-control refresh plan has %d problem(s):" % len(self.problems)]
        lines.extend(f"  - {problem}" for problem in self.problems)
        return chr(10).join(lines)
