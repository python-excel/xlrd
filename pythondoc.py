#
#!/usr/bin/env python
#
# $Id: pythondoc.py 3271 2007-09-09 09:45:14Z fredrik $
# pythondoc documentation generator
#
# history:
# 2003-10-19 fl   first preview release (2.0a1)
# 2003-10-19 fl   fix HTML in descriptor tags, 1.5.2 tweaks, etc (2.0a2)
# 2003-10-20 fl   added encoding support, default HTML generator, etc (2.0a3)
# 2003-10-21 fl   fixed some 1.5.2 issues, etc (2.0b1)
# 2003-10-22 fl   HTML tweaks, pluggable output generators, etc (2.0b2)
# 2003-10-23 fl   fixed encoding, added @author, @version, @since etc
# 2003-10-24 fl   disable XML output by default
# 2003-10-25 fl   moved info properties into an 'info' element
# 2003-10-26 fl   expand wildcards on windows (2.0b3)
# 2003-10-30 fl   added support for RISC OS
# 2003-10-31 fl   (experimental) support module-level comments
# 2003-11-01 fl   minor HTML tweaks (2.0b4)
# 2003-11-03 fl   pythondoc 2.0 final
# 2003-11-15 fl   added support for inline @link/@linkplain tags (2.1b1)
# 2003-11-20 fl   fixed class attribute parsing bug
# 2004-03-27 fl   handle multiple single-line methods
# 2004-09-01 fl   support Python 2.4 decorators (2.1b2)
# 2004-09-21 fl   fixed output filename for "pythondoc ."
# 2005-03-25 fl   added docstring extraction for classes and methods (2.1b3)
# 2005-06-18 fl   fixed correct HTML output when using ElementTree 1.3 (2.1b4)
# 2005-12-23 fl   use xml.etree where available
# 2006-04-04 fl   refactored comment parser code; added -s support (2.1b5)
# 2006-04-06 fl   handle multiple params in docstrings correctly (2.1b6)
# 2007-09-09 fl   moved HTML parser into pythondoc module itself
#
# Copyright (c) 2002-2007 by Fredrik Lundh.
#

##
# This is the PythonDoc tool.  This tool parses Python source files
# and generates API descriptions in XML and HTML.
# <p>
# For more information on the PythonDoc tool and the markup format, see
# <a href="http://effbot.org/zone/pythondoc.htm">the PythonDoc page</a>
# at <a href="http://effbot.org/">effbot.org</a>.
##

# --------------------------------------------------------------------
# Software License
# --------------------------------------------------------------------
#
# Copyright (c) 2002-2007 by Fredrik Lundh
#
# By obtaining, using, and/or copying this software and/or its
# associated documentation, you agree that you have read, understood,
# and will comply with the following terms and conditions:
#
# Permission to use, copy, modify, and distribute this software and
# its associated documentation for any purpose and without fee is
# hereby granted, provided that the above copyright notice appears in
# all copies, and that both that copyright notice and this permission
# notice appear in supporting documentation, and that the name of
# Secret Labs AB or the author not be used in advertising or publicity
# pertaining to distribution of the software without specific, written
# prior permission.
#
# SECRET LABS AB AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH REGARD
# TO THIS SOFTWARE, INCLUDING ALL IMPLIED WARRANTIES OF MERCHANT-
# ABILITY AND FITNESS.  IN NO EVENT SHALL SECRET LABS AB OR THE AUTHOR
# BE LIABLE FOR ANY SPECIAL, INDIRECT OR CONSEQUENTIAL DAMAGES OR ANY
# DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
# WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS
# ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE
# OF THIS SOFTWARE.
#
# --------------------------------------------------------------------

# to do in later releases:
#
# TODO: test this release under 1.5.2 !
# TODO: better rendering of constructors/package modules
# TODO: check @param names against @def/define tags
# TODO: support recursive parsing (-R)
# TODO: warn for tags that doesn't make sense for a given target type
# TODO: HTML output localization (the %s module, returns, raises, etc)
# TODO: make compactHTML generate an element tree instead of raw HTML
#
# nice to have, maybe:
#
# IDEA: support multiple output handlers (multiple -O statements);
#       make -x an alias for -Oxml
# IDEA: make pythondoc self-contained (include stub element implementation)

VERSION_DATE = "2.1b7-20070909"
VERSION = VERSION_DATE.split("-")[0]

COPYRIGHT = "(c) 2002-2007 by Fredrik Lundh"

# explicitly import site (for exemaker etc)
import site

# stuff we use in this module
import glob, os, re, string, sys, tokenize

# make sure elementtree is available
try:
    try:
        import xml.etree.ElementTree as ET
    except ImportError:
        import elementtree.ElementTree as ET
except ImportError:
    raise RuntimeError(
        "PythonDoc %s requires ElementTree 1.1 or later "
        "(available from http://effbot.org/downloads)." % VERSION
        )

# extension separator (not all systems use a period)
try:
    EXTSEP = os.extsep
except AttributeError:
    EXTSEP = "."

##
# Debug level.  The higher the value, the more junk you'll see on
# standard output.
# <p>
# You can use the <b>-V</b> option to <b>pythondoc</b> to increase
# the debug level.

DEBUG = 0

##
# Whitespace tokens.  These are ignored when the parser is scanning
# for a subject.

WHITESPACE_TOKEN = (
    tokenize.NL, tokenize.NEWLINE, tokenize.DEDENT, tokenize.INDENT
    )

##
# Default encoding.  To override this for a module, put a "coding"
# directive in your Python module (see PEP 263 for details).

ENCODING = "iso-8859-1"

##
# Known tags.  The parser generates warnings for tags that are not in
# this list, but it still copies them to the XML infoset.

TAGS = (
    "def", "defreturn",
    "param", "keyparam",
    "return",
    "throws", "exception",
    # javadoc tags not used by the standard generator
    "author", "deprecated", "see", "since", "version"
    )

##
# (Helper) Combines filename prefix with extension part.
#
# @param prefix Filename prefix.
# @param ext Extension string, including a leading period.  The
#    period is replaced with a platform-specific separator, if
#    necessary.
# @return The combined name.

def joinext(prefix, ext):
    assert ext[0] == "." # require leading separator, to match os.path.splitext
    return prefix + EXTSEP + ext[1:]

##
# (Helper) Extracts block tags from a PythonDoc comment.
#
# @param comment Comment text.
# @return A list of (lineno, tag, text) tuples, where the tag is None
#     for the initial description.
# @defreturn List of tuples.

def gettags(comment):

    tags = []

    tag = None
    tag_lineno = lineno = 0
    tag_text = []

    for line in comment:
        if line[:1] == "@":
            tags.append((tag_lineno, tag, string.join(tag_text, "\n")))
            line = string.split(line, " ", 1)
            tag = line[0][1:]
            if len(line) > 1:
                tag_text = [line[1]]
            else:
                tag_text = []
            tag_lineno = lineno
        else:
            tag_text.append(line)
        lineno = lineno + 1

    tags.append((tag_lineno, tag, string.join(tag_text, "\n")))

    return tags

##
# (Helper) Flattens an element tree, returning only the text contents.
#
# @param elem An element tree.
# @return A text string.
# @defreturn String.

def flatten(elem):
    text = elem.text or ""
    for e in elem:
        text += flatten(e)
        if e.tail:
            text += e.tail
    return text

##
# (Helper) Extracts summary from a PythonDoc comment.  This function
# gets the first complete sentence from the description string.
#
# @param description An element containing the description.
# @return A summary string.
# @defreturn String.

def getsummary(description):

    description = flatten(description)

    # extract the first sentence from the description
    m = re.search("(?s)(.+?\.)\s", description + " ")
    if m:
        return m.group(1)

    return description # sorry

##
# (Helper) Parses HTML descriptor text into an XHTML structure.
#
# @param parser Parser instance (provides a warning method).
# @param text Text fragment.
# @return An element tree containing XHTML data.
# @defreturn Element.

def parsehtml(parser, tag, text, lineno):

    # transcode
    if parser.encoding != "ascii":
        try:
            text = unicode(text, parser.encoding)
        except NameError:
            pass # 1.5.2

    # process inline links (@link, @linkplain)
    # note that links are replaced with <a href='link:...> elements;
    # the href's are resolved in a later step
    if "{" in text:
        def fixlink(m, parser=parser, lineno=lineno):
            linkdef = string.split(m.group(1), None, 2)
            if len(linkdef) == 2:
                # default text is same as last part of target name
                href = linkdef[1]
                if href[:1] == "#":
                    href = href[1:]
                linkdef.append(string.split(href, ".")[-1])
            if len(linkdef) < 3 or linkdef[0] not in ("@link", "@linkplain"):
                parser.warning(
                    (lineno, 0),
                    "malformed @link near this line",
                    )
                return m.group(0)
            type, href, text = linkdef
            href = "link:" + html_encode(href)
            if type == "@link":
                return "<a href='%s' class='link'><b>%s</b></a>" % (href, text)
            else:
                return "<a href='%s' class='linkplain'>%s</a>" % (href, text)
        text = re.sub("\{(@link[^}]+)\}", fixlink, text)

    if "<" not in text and "&" not in text:
        # plain text
        elem = ET.Element(tag)
        elem.text = string.strip(text)
        return elem

    p = HTMLTreeBuilder()
    ix = 0
    try:
        p.feed("<%s>" % tag)
        p.feed("<p>") # make sure everything's wrapped in a paragraph tag
        # feed line by line
        for line in string.split(text, "\n"):
            p.feed(line + "\n")
            ix = ix + 1
        p.feed("</%s>" % tag)
        tree = p.close()
    except:
        parser.warning(
            (lineno+ix, 0),
            "HTML parser error near this line (%s)",
            sys.exc_value
            )
        return ET.Element("p")

    return tree

# --------------------------------------------------------------------
# copied from ElementTree/HTMLTreeBuilder.py

import htmlentitydefs

AUTOCLOSE = "p", "li", "tr", "th", "td", "head", "body"
IGNOREEND = "img", "hr", "meta", "link", "br"

if sys.version[:3] == "1.5":
    is_not_ascii = re.compile(r"[\x80-\xff]").search # 1.5.2
else:
    is_not_ascii = re.compile(eval(r'u"[\u0080-\uffff]"')).search

try:
    from HTMLParser import HTMLParser
except ImportError:
    from sgmllib import SGMLParser
    # hack to use sgmllib's SGMLParser to emulate 2.2's HTMLParser
    class HTMLParser(SGMLParser):
        # the following only works as long as this class doesn't
        # provide any do, start, or end handlers
        def unknown_starttag(self, tag, attrs):
            self.handle_starttag(tag, attrs)
        def unknown_endtag(self, tag):
            self.handle_endtag(tag)

##
# ElementTree builder for HTML source code.  This builder converts an
# HTML document or fragment to an ElementTree.
# <p>
# The parser is relatively picky, and requires balanced tags for most
# elements.  However, elements belonging to the following group are
# automatically closed: P, LI, TR, TH, and TD.  In addition, the
# parser automatically inserts end tags immediately after the start
# tag, and ignores any end tags for the following group: IMG, HR,
# META, and LINK.
#
# @keyparam builder Optional builder object.  If omitted, the parser
#     uses the standard <b>elementtree</b> builder.
# @keyparam encoding Optional character encoding, if known.  If omitted,
#     the parser looks for META tags inside the document.  If no tags
#     are found, the parser defaults to ISO-8859-1.  Note that if your
#     document uses a non-ASCII compatible encoding, you must decode
#     the document before parsing.

class HTMLTreeBuilder(HTMLParser):

    def __init__(self, encoding=None):
        self.__stack = []
        self.__builder = ET.TreeBuilder()
        self.encoding = encoding or "iso-8859-1"
        HTMLParser.__init__(self)

    ##
    # Flushes parser buffers, and return the root element.
    #
    # @return An Element instance.

    def close(self):
        HTMLParser.close(self)
        return self.__builder.close()

    ##
    # (Internal) Handles start tags.

    def handle_starttag(self, tag, attrs):
        if tag == "meta":
            # look for encoding directives
            http_equiv = content = None
            for k, v in attrs:
                if k == "http-equiv":
                    http_equiv = string.lower(v)
                elif k == "content":
                    content = v
            if http_equiv == "content-type" and content:
                # use mimetools to parse the http header
                import mimetools, StringIO
                header = mimetools.Message(
                    StringIO.StringIO("%s: %s\n\n" % (http_equiv, content))
                    )
                encoding = header.getparam("charset")
                if encoding:
                    self.encoding = encoding
        if tag in AUTOCLOSE:
            if self.__stack and self.__stack[-1] == tag:
                self.handle_endtag(tag)
        self.__stack.append(tag)
        attrib = {}
        if attrs:
            for k, v in attrs:
                attrib[string.lower(k)] = v
        self.__builder.start(tag, attrib)
        if tag in IGNOREEND:
            self.__stack.pop()
            self.__builder.end(tag)

    ##
    # (Internal) Handles end tags.

    def handle_endtag(self, tag):
        if tag in IGNOREEND:
            return
        lasttag = self.__stack.pop()
        if tag != lasttag and lasttag in AUTOCLOSE:
            self.handle_endtag(lasttag)
        self.__builder.end(tag)

    ##
    # (Internal) Handles character references.

    def handle_charref(self, char):
        if char[:1] == "x":
            char = int(char[1:], 16)
        else:
            char = int(char)
        if 0 <= char < 128:
            self.__builder.data(chr(char))
        else:
            self.__builder.data(unichr(char))

    ##
    # (Internal) Handles entity references.

    def handle_entityref(self, name):
        entity = htmlentitydefs.entitydefs.get(name)
        if entity:
            if len(entity) == 1:
                entity = ord(entity)
            else:
                entity = int(entity[2:-1])
            if 0 <= entity < 128:
                self.__builder.data(chr(entity))
            else:
                self.__builder.data(unichr(entity))
        else:
            self.unknown_entityref(name)

    ##
    # (Internal) Handles character data.

    def handle_data(self, data):
        if isinstance(data, type('')) and is_not_ascii(data):
            # convert to unicode, but only if necessary
            data = unicode(data, self.encoding, "ignore")
        self.__builder.data(data)

    ##
    # (Hook) Handles unknown entity references.  The default action
    # is to ignore unknown entities.

    def unknown_entityref(self, name):
        pass # ignore by default; override if necessary

# --------------------------------------------------------------------

##
# (Helper) Parses a PythonDoc comment into an PythonDoc info structure.
#
# @param parser Parser instance (provides a warning method).
# @param lineno Line number where this comment starts.
# @param comment A list of text line making up the comment.
# @param dedent If true, strip leading whitespace from all comment
#     lines except the first one.
# @return An element tree containing XHTML data.
# @defreturn Element.

def parsecomment(parser, lineno, comment, dedent=0):

    subject_info = ET.Element("info")

    # untabify
    for ix in range(len(comment)):
        comment[ix] = string.expandtabs(comment[ix])

    if dedent:
        margin = None
        for ix in range(1, len(comment)):
            s = string.lstrip(comment[ix])
            if not s:
                continue
            m = len(comment[ix]) - len(s)
            if margin is None:
                margin = m
            else:
                margin = min(m, margin)
        if margin:
            for ix in range(1, len(comment)):
                comment[ix] = comment[ix][margin:]

    for ix, tag, text in gettags(comment):

        pos = lineno + ix + 1, 0

        # check tag name
        if tag is None:
            tag = "description"
        else:
            if tag not in TAGS:
                parser.warning(
                    pos,
                    "unknown tag in description: @%s", tag
                    )
            if tag in ("throws", "exception"):
                tag = "exception" # PythonDoc extension

        # deal with "named" tags
        if tag in ("param", "keyparam", "exception"):
            text = string.split(text, " ", 1)
            name = text[0]
            if len(text) > 1:
                text = string.lstrip(text[1])
            else:
                text = ""
        else:
            name = None

        tag_elem = parsehtml(parser, tag, text, pos[0])

        # generate summaries
        if tag == "description":
            summary = getsummary(tag_elem)
            if summary:
                elem = ET.SubElement(subject_info, "summary")
                elem.text = summary

        subject_info.append(tag_elem)

        if name:
            tag_elem.set("name", name)

    return subject_info

##
# Module parser.
# <p>
# This class implements the PythonDoc source code scanner.  It reads
# source code from a file or a file-like object, and builds an element
# tree with information about the module.
# <p>
# Note that the constructor only sets things up for parsing.  Use the
# {@link ModuleParser.parse} method to parse the file.  Or for
# convenience, use the {@link parse} function to create a parser
# object and parse a given file.
#
# @param file Name of the module source file, or a file object.  If a
#    file object is used, it must provide a <b>name</b> attribute and
#    a <b>readline</b> method.
# @param prefix Optional name prefix.  If given, this is prepended to
#    the module name.  For example, if the prefix is set to "prefix"
#    and the module filename is "name.py", the module is assumed to
#    contain the "prefix.name" namespace.

class ModuleParser:

    ##
    # Module name.

    name = None

    def __init__(self, file, prefix=None):
        if hasattr(file, "readline"):
            self.file = file
            self.filename = file.name
        else:
            self.file = None
            self.filename = file
        name = os.path.splitext(os.path.basename(self.filename))[0]
        if prefix and prefix != ".":
            name = prefix + "." + name
        self.name = name
        self.stack = [
            ET.Element(
                "module",
                name=name, filename=self.filename
                )
            ]
        self.indent = 0
        self.scope = [] # list of (indent, tag, name, ...) tuples
        self.handler = self.look_for_encoding
        self.encoding = ENCODING

    ##
    # Parses the file.
    #
    # @keyparam docstring If true, look for markup in docstrings.
    # @return An element tree containing information about the module.
    # @defreturn Element.
    # @exception IOError If the file could not be opened.

    def parse(self, docstring=0):
        if self.file is None:
            file = open(self.filename)
        else:
            file = self.file
        try:
            tokenize.tokenize(file.readline, self.handle_token)
        except tokenize.TokenError, v:
            message, lineno = v
            self.warning(lineno, "exception in tokenizer: %s", message)
        if len(self.stack) != 1:
            pass # FIXME: print warning?
        tree = self.stack[0] # may be incomplete
        # fixup internal links
        # 1) find all named elements
        elems = {}
        for elem in tree.getiterator():
            name = elem.get("name")
            if name:
                elems[name] = elem
        # 2) find all link anchors
        for elem in tree.getiterator("a"):
            href = elem.get("href")
            if href[:5] == "link:":
                # FIXME: add support for external links
                href = href[5:]
                if href[:1] == "#":
                    href = href[1:]
                target = elems.get(self.name + "." + href)
                if target:
                    href = "#" + target.get("name") + "-" + target.tag
                    elem.set("href", href)
        if docstring:
            # look for markup in docstrings
            for info in tree.getiterator("info"):
                docstring = info.findtext("docstring")
                if not docstring:
                    continue
                comment = docstring.split("\n")
                newinfo = parsecomment(self, 0, comment, dedent=1)
                for elem in info:
                    if newinfo.find(elem.tag) is None:
                        newinfo.append(elem)
                info[:] = newinfo
        return tree

    ##
    # Prints a warning message to standard output.
    #
    # @param position A (line, column) tuple.  The column can be set
    #     to None if not known (or not relevant).
    # @param format Message or format string.
    # @param *args Optional arguments.

    def warning(self, position, format, *args):
        line, column = position
        message = "%s:%d: WARNING: %s" % (self.filename, line, format % args)
        sys.stderr.write(message)
        sys.stderr.write("\n")

    ##
    # Dispatches tokens to the current handler.  Each handler should
    # return the handler to call for the next token.
    # <p>
    # This method also handles indentation and dedentation tokens,
    # and manages the scope stack.

    def handle_token(self, *args):
        # dispatch incoming tokens to the current handler
        if DEBUG > 1:
            print self.handler.im_func.func_name, self.indent,
            print tokenize.tok_name[args[0]], repr(args[1])
        if args[0] == tokenize.DEDENT:
            self.indent = self.indent - 1
            while self.scope and self.scope[-1][0] >= self.indent:
                del self.scope[-1]
                del self.stack[-1]
        self.handler = apply(self.handler, args)
        if args[0] == tokenize.INDENT:
            self.indent = self.indent + 1

    ##
    # (Token handler) Scans for encoding directive.

    def look_for_encoding(self, type, token, start, end, line):
        if type == tokenize.COMMENT:
            if string.rstrip(token) == "##":
                return self.look_for_pythondoc(type, token, start, end, line)
            m = re.search("coding[:=]\s*([-_.\w]+)", token)
            if m:
                self.encoding = m.group(1)
                return self.look_for_pythondoc
        if start[0] > 2:
            return self.look_for_pythondoc
        return self.look_for_encoding

    ##
    # (Token handler) Scans for PythonDoc comments.

    def look_for_pythondoc(self, type, token, start, end, line):
        if type == tokenize.COMMENT and string.rstrip(token) == "##":
            # found a comment: set things up for comment processing
            self.comment_start = start
            self.comment = []
            return self.process_comment_body
        else:
            # deal with "bare" subjects
            if token == "def" or token == "class":
                self.subject_indent = self.indent
                self.subject_parens = 0
                self.subject_start = self.comment_start = None
                self.subject = []
                return self.process_subject(type, token, start, end, line)
            return self.look_for_pythondoc

    ##
    # (Token handler) Processes a comment body.  This handler adds
    # comment lines to the current comment.

    def process_comment_body(self, type, token, start, end, line):
        if type == tokenize.COMMENT:
            if start[1] != self.comment_start[1]:
                self.warning(
                    start,
                    "comment line should be aligned with marker"
                    )
            line = string.rstrip(token)
            if line == "##":
                # handle module comments (experimental)
                # FIXME: add more consistency checks?
                if self.stack[0].find("info") is not None:
                    self.warning(
                        self.comment_start,
                        "multiple module comments are not allowed"
                        )
                    # FIXME: ignore additional comments?
                self.process_subject_info(None, self.stack[0])
                return self.look_for_pythondoc
            elif line[:2] == "# ":
                line = line[2:]
            elif line[:1] == "#":
                line = line[1:]
            self.comment.append(line)
        else:
            if not self.comment:
                self.warning(
                    self.comment_start,
                    "found pythondoc marker but no comment body"
                    )
                return self.look_for_pythondoc
            self.subject_start = None
            self.subject = []
            if type != tokenize.NL:
                return self.process_subject(type, token, start, end, line)
            return self.process_subject # end of comment
        return self.process_comment_body

    ##
    # (Token handler) Processes the comment subject.  The subject can
    # be either a plain variable, or a function/method or class
    # definition.
    # <p>
    # This method is also used to process "bare" subjects; that is,
    # functions, methods, and classes that don't have PythonDoc
    # markup.  In that case, the comment_start variable is set to
    # None.

    def process_subject(self, type, token, start, end, line):
        # got an item; deal with it
        if self.subject:
            # method/function/class definition
            definition = self.subject[0] in ("def", "class")
            if definition:
                if type not in WHITESPACE_TOKEN:
                    if token == "(":
                        self.subject_parens = self.subject_parens + 1
                    elif token == ")":
                        self.subject_parens = self.subject_parens - 1
                if self.subject_parens or token != ":":
                    self.subject.append(token)
                    return self.process_subject
            else:
                # simple assignment
                if token != "=":
                    self.warning(
                        self.subject_start,
                        "bad subject %s; ignoring description",
                        repr(self.subject[0])
                        )
                    # might be a pythondoc marker; pass it to the scanner
                    return self.look_for_pythondoc(
                        type, token, start, end, line
                        )
                # FIXME: keep adding stuff until end of expression
        else:
            if type in WHITESPACE_TOKEN:
                return self.process_subject
            if type == tokenize.COMMENT:
                self.warning(
                    start,
                    "comment between description and subject; " +
                    "ignoring description"
                    )
                # might be a pythondoc marker; pass it to the scanner
                return self.look_for_pythondoc(
                    type, token, start, end, line
                    )
            # FIXME: check token type!
            # the @ token type is currently tokenize.ERRORTOKEN; hopefully
            # this will change before 2.4 final
            if token == "@":
                self.decorator_parens = 0
                return self.skip_decorator
            self.subject_start = start
            self.subject.append(token)
            if token in ("def", "class"):
                # handle single-line subjects
                while self.scope and self.scope[-1][0] >= self.indent:
                    self.scope.pop()
                    self.stack.pop()
                self.subject_indent = self.indent
                self.subject_parens = 0
            return self.process_subject

        # check if this is a method or a function
        method = self.scope and self.scope[-1][1] == "class"

        # calculate fully qualified subject name
        name = [self.name]
        for s in self.scope:
            name.append(s[2])
        if definition:
            name.append(self.subject[1])
        else:
            name.append(self.subject[0])

        # calculate subject definition statement
        statement = []
        for part in self.subject:
            if part in ("class", "def"):
                continue
            statement.append(part)
            if part == ",":
                statement.append(" ")
        if self.subject[0] == "def" and method:
            # ignore the first argument for methods
            # 'name', '(', 'self', ',', ' ', ...)
            del statement[2:min(5, len(statement)-1)]
        statement = string.join(statement, "")

        # create subject element
        if self.subject[0] == "class":
            subject_elem = ET.Element("class")
        elif self.subject[0] == "def":
            if method:
                subject_elem = ET.Element("method")
            else:
                subject_elem = ET.Element("function")
        else:
            subject_elem = ET.Element("variable")

        self.stack[-1].append(subject_elem)

        # add new subject to the scope and element stacks
        if definition:
            self.scope.append((self.subject_indent,) + tuple(self.subject))
            self.stack.append(subject_elem)

        subject_info = self.process_subject_info(name, subject_elem)

        # add local name to info
        elem = ET.Element("name")
        elem.text = name[-1]

        subject_info.insert(0, elem)

        name = string.join(name, ".")

        subject_elem.set("name", name)
        subject_elem.set("lineno", str(self.subject_start[0]))

        if subject_info.find("def") is None and statement:
            # add subject definition (unless specified in comment)
            elem = ET.Element("def")
            elem.text = statement
            # add to front, to make the XML easier to read
            subject_info.insert(0, elem)

        if definition:
            return self.look_for_docstring(type, token, start, end, line)
        else:
            return self.look_for_pythondoc(type, token, start, end, line)

    ##
    # (Token handler) Skips a decorator.

    def skip_decorator(self, type, token, start, end, line):
        if token == "(":
            self.decorator_parens = self.decorator_parens + 1
        elif token == ")":
            self.decorator_parens = self.decorator_parens - 1
        if self.decorator_parens or type != tokenize.NEWLINE:
            return self.skip_decorator
        return self.process_subject

    ##
    # (Token handler helper) Processes a PythonDoc comment.  This
    # method creates an "info" element based on the current comment,
    # and attaches it to the current subject element.
    #
    # @param subject_name Subject name (or None if the name is not known).
    # @param subject_elem The current subject element.
    # @return The info element.  Note that this element has already
    #     been attached to the subject element.
    # @defreturn Element

    def process_subject_info(self, subject_name, subject_elem):

        # process pythondoc comment (if any)
        if self.comment_start:
            subject_info = parsecomment(
                self, self.comment_start[0], self.comment
                )
        else:
            subject_info = ET.Element("info")

        subject_elem.append(subject_info)

        if DEBUG:
            if subject_name:
                subject_name = string.join(subject_name, ".")
            else:
                subject_name = "<module>"
            print "---", subject_name
            prefix = "   " * len(self.scope)
            for line in self.comment:
                print prefix + line

        return subject_info

    ##
    # (Token handler) Look for docstring inside a definition.

    def look_for_docstring(self, type, token, start, end, line):
        if type in WHITESPACE_TOKEN or token == ":":
            return self.look_for_docstring
        if type == tokenize.STRING:
            subject_elem = self.stack[-1]
            subject_info = subject_elem.find("info")
            if subject_info is not None:
                elem = ET.SubElement(subject_info, "docstring")
                # FIXME: add string sanity check here, before doing eval
                elem.text = eval(token)
        return self.skip_subject_body(type, token, start, end, line)

    ##
    # (Token handler) Skips over the subject body.

    def skip_subject_body(self, type, token, start, end, line):
        # for now, just hand control back to the pythondoc scanner,
        # and let it skip over the subject body while looking for the
        # next marker.
        return self.look_for_pythondoc(type, token, start, end, line)

##
# Parses a module.
# <p>
# This function creates a {@link #ModuleParser} instance, and uses it
# to parse the given file.  For details, see {@linkplain #ModuleParser
# the <b>ModuleParser</b> documentation}.
#
# @param file Name of the module source file, or a file object.
# @param prefix Optional name prefix.
# @keyparam docstring If true, look for markup in docstrings.
# @return An element tree containing the module description.
# @defreturn Element.
# @exception IOError If the file could not be found, or could not
#    be opened for reading.

def parse(file, prefix=None, docstring=0):
    m = ModuleParser(file, prefix)
    return m.parse(docstring=docstring)

# --------------------------------------------------------------------
# default formatter

if sys.version[:3] == "1.5":
    _escape = re.compile(r"[&<>\"\x80-\xff]") # 1.5.2
else:
    _escape = re.compile(eval(r'u"[&<>\"\u0080-\uffff]"'))

_escape_map = {
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
}

##
# Encodes reserved HTML characters and non-ASCII characters as HTML
# character references.
#
# @def html_encode(text)
# @param text Source text.
# @return An encoded string.

def html_encode(text, pattern=_escape):
    if not text:
        return ""
    def escape_entities(m, map=_escape_map):
        char = m.group()
        text = map.get(char)
        if text is None:
            text = "&#%d;" % ord(char)
        return text
    text = pattern.sub(escape_entities, text)
    try:
        return text.encode("ascii")
    except AttributeError:
        return text # 1.5.2

##
# Compact HTML formatter.  This formatter turns a module XML
# description into a minimal HTML document.
# <p>
# This formatter supports the following options:
# </p>
# <dl>
# <dt><b>-Dstyle</b>=URL</dt>
# <dd>Stylesheet URL.  If this option is present, the formatter adds a
# stylesheet &lt;link&gt; element to the HTML output.</dd>
# <dt><b>-Dzone</b></dt>
# <dd>Generate effbot.org zone documents.</dd>
# </dl>
#
# @param options Options dictionary.

class CompactHTML:

    def __init__(self, options=None):
        self.options = options or {}

    ##
    # Writes an element containing some text (plain or formatted).
    #
    # @param elem Element.
    # @param compact If true, try to minimize the amount of vertical
    #     padding.

    def writetext(self, elem, compact=0):
        if len(elem):
            if compact and len(elem) == 1 and elem[0].tag == "p":
                elem = elem[0]
                self.file.write(html_encode(elem.text))
                for e in elem:
                    ET.ElementTree(e).write(self.file)
                self.file.write(html_encode(elem.tail))
            else:
                for e in elem:
                    ET.ElementTree(e).write(self.file)
        elif elem is not None and elem.text:
            if compact:
                self.file.write(html_encode(elem.text))
            else:
                self.file.write("<p>%s</p>\n" % html_encode(elem.text))

    ##
    # Writes an object description (the description text, parameters,
    # return values) etc.
    #
    # @param object The object element.
    # @param summary If true, use summary instead of full description.

    def writeobject(self, object, summary=0):
        name = object.get("name")
        info = object.find("info")
        localname = string.split(name, ".")[-1]
        define = info.findtext("def")
        anchor = html_encode(name + "-" + object.tag)
        if object.tag == "class":
            # look for the constructor
            for obj in object:
                if obj.get("name") == name + ".__init__":
                    inf = obj.find("info")
                    define = string.split(inf.findtext("def"), "(", 1)
                    define = localname + "(" + define[1]
                    self.file.write(
                        "<dt><b>%s</b> (class)" % html_encode(define)
                        )
                    break
            else:
                self.file.write(
                    "<dt><b>%s</b> (class) " % html_encode(localname)
                    )
        elif object.tag == "variable":
            self.file.write(
                "<dt><a id='%s' name='%s'><b>%s</b></a> (variable)" % (
                    anchor, anchor, html_encode(localname)
                    )
                )
        else:
            self.file.write(
                "<dt><a id='%s' name='%s'><b>%s</b></a>" % (
                    anchor, anchor, html_encode(define)
                    )
                )
            if object.tag == "function" or object.tag == "method":
                defreturn = info.find("defreturn")
                if defreturn is not None:
                    defreturn = flatten(defreturn)
                    self.file.write(" &rArr; %s" % html_encode(defreturn))
        self.file.write(" [<a href='#%s'>#</a>]</dt>\n" % anchor)
        self.file.write("<dd>\n")
        if summary:
            text = info.findtext("summary")
            if text:
                self.file.write("<p>%s</p>\n" % html_encode(text))
        else:
            self.writetext(info.find("description"))
        param = info.findall("param")
        keyparam = info.findall("keyparam")
        exception = info.findall("exception")
        return_elem = info.find("return")
        if param or keyparam or exception or return_elem != None:
            self.file.write("<dl>\n")
            for p in param + keyparam:
                name = p.get("name")
                if p.tag == "keyparam":
                    name = name + "="
                self.file.write("<dt><i>%s</i></dt>\n" % html_encode(name))
                self.file.write("<dd>\n")
                self.writetext(p, compact=1)
                self.file.write("</dd>\n")
            if return_elem is not None:
                self.file.write("<dt>Returns:</dt>\n")
                self.file.write("<dd>\n")
                self.writetext(return_elem, compact=1)
                self.file.write("</dd>\n")
            for e in exception:
                name = html_encode(e.get("name"))
                self.file.write("<dt>Raises <b>%s</b>:</dt>" % name)
                self.file.write("<dd>\n")
                self.writetext(e, compact=1)
                self.file.write("</dd>\n")
            self.file.write("</dl><br />\n")
        if object.tag == "class" and summary:
            self.file.write(
                "<p>For more information about this class, see "
                "<a href='#%s'><i>The %s Class</i></a>.</p>\n" % (
                    anchor, localname
                    )
                )
        self.file.write("</dd>\n")

    ##
    # Writes a module description to file.
    #
    # @param module A module element tree, as returned by {@link
    # ModuleParser.parse}.
    # @param file Output file name (minus extension).
    # @return If successful, the output filename used to store the
    #     module.
    # @defreturn String or None.

    def save(self, module, file):

        title = "The %s Module" % module.get("name")

        zone = self.options.has_key("zone")

        if zone:
            filename = joinext(file, ".txt")
        else:
            filename = joinext(file, ".html")
        self.file = open(filename, "w")

        if zone:
            # generate zone document
            self.file.write(title + "\n\n")
        else:
            self.file.write(
                "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' "
                "'http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd'>\n"
                )
            self.file.write("<html>\n<head>\n")
            self.file.write(
                "<meta http-equiv='Content-Type' "
                "content='text/html; charset=us-ascii' />\n"
                )
            self.file.write("<title>%s</title>\n" % html_encode(title))
            try:
                style = self.options["style"]
            except KeyError:
                pass
            else:
                self.file.write(
                    "<link rel='stylesheet' href='%s' type='text/css' />\n" %
                    html_encode(style)
                    )
            self.file.write("</head>\n<body>\n")
            self.file.write("<h1>%s</h1>\n" % title)

        # 0) module comments
        info = module.find("info")
        if info is not None:
            self.writetext(info.find("description"))
            self.file.write("<h2>Module Contents</h2>\n")

        # 1) toplevel subjects (including class overviews)
        objects = []
        for object in module:
            info = object.find("info")
            if info is None or info.find("description") is None:
                continue
            if object.tag in ("variable", "function", "class"):
                objects.append(object)
        objects.sort(lambda a, b: cmp(
            string.lower(string.split(a.get("name"), ".")[-1]),
            string.lower(string.split(b.get("name"), ".")[-1])
            ))
        self.file.write("<dl>\n")
        for object in objects:
            self.writeobject(object, object.tag == "class")
        self.file.write("</dl>\n")
        # 2) class descriptions
        for object in objects:
            if object.tag != "class":
                continue
            name = object.get("name")
            localname = string.split(name, ".")[-1]
            anchor = name + "-class"
            self.file.write(
                "<h2><a id='%s' name='%s'>The %s Class</a></h2>\n" % (
                    anchor, anchor, localname
                    )
                )
            self.file.write("<dl>\n")
            self.writeobject(object)
            objects = []
            for object in object:
                info = object.find("info")
                if info is None or info.find("description") is None:
                    continue
                if object.tag not in ("method", "variable"):
                    continue
                objects.append(object)
            objects.sort(lambda a, b: cmp(
            string.lower(string.split(a.get("name"), ".")[-1]),
            string.lower(string.split(b.get("name"), ".")[-1])
                ))
            for object in objects:
                if object.tag == "variable":
                    object.tag = "attribute"
                self.writeobject(object)
                if object.tag == "attribute":
                    object.tag = "variable"
            self.file.write("</dl>\n")

        if not zone:
            self.file.write("</body></html>\n")

        self.file.close()
        self.file = None

        return filename

##
# Prints a usage message and exits.

def usage():
    print "PythonDoc", VERSION, COPYRIGHT
    print
    print "Usage:"
    print
    print "  pythondoc [options] files..."
    print
    print "where the files can be either python modules or package"
    print "directories."
    print
    print "Options:"
    print
    print "  -p prefix    Prepend given prefix to symbol names."
    print "  -f           Generate output also for files without descriptions."
    print "  -x           Generate XML output (pythondoc infosets)."
    print
    print "  -s           Look for markup in docstrings (experimental)."
    print
    print "Output options:"
    print
    print "  -O format    Use given output format handler."
    print "  -D name      Define output variable."
    print "  -D name=text Set output variable to given text."
    print
    print "For more information on PythonDoc and the PythonDoc comment syntax,"
    print "see http://effbot.org/zone/pythondoc.htm"
    sys.exit(1)

if __name__ == "__main__":

    import getopt

    try:
        opts, args = getopt.getopt(sys.argv[1:], "D:fO:p:Vsx")
    except getopt.error:
        usage()

    force = 0
    prefix = None
    docstring = 0
    output_xml = 0
    output_handler = CompactHTML
    output_options = {}

    for k, v in opts:
        if k == "-f":
            force = 1
        elif k == "-p":
            prefix = v
        elif k == "-s":
            docstring = 1
        elif k == "-x":
            output_xml = 1
        elif k == "-O":
            try:
                m = __import__(v)
                for k in string.split(v, ".")[1:]:
                    m = getattr(m, k)
                output_handler = getattr(m, "PythonDocGenerator")
            except (ImportError, AttributeError):
                print "cannot find/load", repr(v), "generator"
                sys.exit(1)
        elif k == "-D":
            try:
                k, v = string.split(v, "=", 1)
            except ValueError:
                k = v; v = None
            output_options[k] = v
        elif k == "-V":
            DEBUG = DEBUG + 1

    if not args:
        usage()

    # instantiate output handler
    output_handler = output_handler(output_options)

    # check if handler supports custom tags
    try:
        TAGS = TAGS + output_handler.tags
    except AttributeError:
        pass

    import time
    t0 = time.time()

    input = output = 0

    for filename in args:

        this_prefix = prefix

        if os.path.isdir(filename):
            # FIXME: explicitly check if this is a package?
            files = glob.glob(os.path.join(filename, joinext("*", ".py")))
            if not this_prefix:
                this_prefix = os.path.basename(filename)
        else:
            if sys.platform == "win32" and glob.has_magic(filename):
                files = glob.glob(filename)
            else:
                files = [filename]

        files.sort()

        for file in files:

            try:
                module = parse(file, this_prefix, docstring=docstring)
            except IOError, v:
                sys.stderr.write("%s error: %s\n" % (file, v[1]))
                continue

            input = input + 1

            # check if any toplevel object has a description
            if not force:
                for n in module:
                    i = n.find("info")
                    if i and i.find("description") is not None:
                        break
                else:
                    continue # no documented subjects

            f = "pythondoc-" + string.replace(module.get("name"), ".", EXTSEP)

            if output_xml:
                # generate XML
                filename = joinext(f, ".xml")
                try:
                    out = open(filename, "w")
                    ET.ElementTree(module).write(out)
                    out.close()
                except IOError, v:
                    sys.stderr.write("%s error: %s\n" % (filename, v[1]))
                else:
                    sys.stderr.write("%s ok\n" % filename)

            # generate output
            try:
                out = output_handler.save(module, f)
            except IOError, v:
                sys.stderr.write("%s error: %s\n" % (file, v[1]))
            else:
                if out:
                    sys.stderr.write("%s ok\n" % out)

            output = output + 1

    # flush output handler
    try:
        done = output_handler.done
    except AttributeError:
        pass
    else:
        out = output_handler.done()
        if out:
            sys.stderr.write("%s ok\n" % out)

    if DEBUG:
        sys.stderr.write(
            "%d files parsed, %d descriptions generated, in %.2f seconds\n" % (
                input, output, time.time() - t0
                ))
