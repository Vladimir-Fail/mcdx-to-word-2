"""
Mathcad 14/15 (.xmcd / .xml) → Markdown with LaTeX → DOCX converter.

WHY THIS FILE EXISTS
--------------------
Mathcad 14/15 stores worksheets as XML with several namespaces
(`ml:`, `ws:`, `u:`, `p:`). The XML is verbose and quirky; there is
no formal spec available. This module reverse-engineers the relevant
parts and renders them as LaTeX so that Pandoc can turn the Markdown
into a Word document with editable equations.

NOTES FOR AI / FUTURE MAINTAINERS
---------------------------------
The Mathcad XML dialect has many "gotchas". This file documents them
inline as `NOTE:` blocks. Whenever you touch a parser branch, please
add a short NOTE explaining the specific Mathcad construct it handles.
Every quirk encoded below was observed in real .xmcd files (14.1).

Big-picture pipeline:

    .xmcd (XML)  --[MathcadParser.process_file]-->  Markdown+LaTeX
    Markdown+LaTeX  --[WordConverter.convert_to_word]-->  .docx

Two numeric-formatting knobs matter downstream:
  * `sig_figs_small` / `sig_figs_large` — significand digits.
  * `sci_threshold` — switch small numbers to scientific notation.

The parser uses a Russian-locale convention: decimal comma ",".
This matches how Mathcad regional output looks and how the resulting
Word document is expected to read.
"""

import math
import os
import re
import sys
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pypandoc


# ---------------------------------------------------------------------------
# Optional drag-and-drop support. Only enabled if tkinterdnd2 is installed.
# ---------------------------------------------------------------------------
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_SUPPORTED = True
except ImportError:
    DND_SUPPORTED = False


# ===========================================================================
# MathcadParser — the core XML → LaTeX translator
# ===========================================================================
class MathcadParser:
    """
    Converts Mathcad XML/XMCD worksheet regions into LaTeX fragments.

    NOTE (namespace handling):
        Mathcad emits tags such as `<ml:apply>` and `<ws:region>`. The
        namespace prefix changes nothing semantically for us; we always
        strip it via `strip_ns()` before comparing tag names. This lets
        us match on plain "apply", "region", "id", etc.

    NOTE (region ordering):
        Regions are laid out on a 2D canvas. Each `<region>` has a `top`
        attribute (float, in points). We sort by `top` to reconstruct the
        visual reading order. Column (`left`) is ignored because Mathcad
        users usually write top-to-bottom.
    """

    # Numeric formatting knobs (overridable from the GUI).
    sig_figs_small = 4
    sig_figs_large = 8

    # Number of decimal places for angles in polar form.
    # Mathcad itself prints angles with 15+ digits; we round for readability.
    ANGLE_DECIMALS = 1

    # Scientific-notation threshold: numbers with |x| < 10^sci_threshold
    # are rendered as "m·10^k". `None` disables the feature.
    sci_threshold = None

    # -------------------------------------------------------------------
    # Low-level helpers
    # -------------------------------------------------------------------

    @staticmethod
    def strip_ns(tag):
        """
        Remove an XML namespace prefix from a tag or attribute key.

        NOTE: Mathcad uses several namespaces (ml:, ws:, u:, p:). The
        namespace URIs differ between Mathcad versions and locales, so
        we cannot match on the full "{uri}tag" form. Stripping the URI
        is the only robust cross-version approach.
        """
        if isinstance(tag, str) and "}" in tag:
            return tag.split("}", 1)[1]

        return tag

    @classmethod
    def find_first_by_tag(cls, node, tag):
        """
        Depth-first search for the first descendant with a given tag.

        NOTE: Mathcad nests elements deeply. For example `<eval>` may
        contain `<result>` only after several levels of `<apply>` /
        `<sequence>` wrappers. A recursive search is simpler and safer
        than hand-coding the exact paths.
        """
        if node is None:
            return None

        for child in node.iter():
            if cls.strip_ns(child.tag) == tag:
                return child

        return None

    # -------------------------------------------------------------------
    # Numeric formatting
    # -------------------------------------------------------------------

    @classmethod
    def _compute_target_sig_figs(cls, abs_value):
        """
        Choose how many significant digits to show for a real number.

        Rules:
          * For |x| >= 1: clamp between sig_figs_small and sig_figs_large
            based on the integer digit count.
          * For |x| <  1: use sig_figs_small.
          * Bonus digit when the leading digit is 1 or 2 (heuristic that
            matches common engineering convention).

        NOTE: Mathcad's own display precision is set in the worksheet
        settings (often 4-15). We deliberately choose a smaller, more
        readable set because the Word document is meant for humans.
        """
        if abs_value >= 1:
            int_digits = len(str(int(abs_value)))
            base_sig_figs = max(
                cls.sig_figs_small,
                min(cls.sig_figs_large, int_digits)
            )
        else:
            base_sig_figs = cls.sig_figs_small

        sci_str = f"{abs_value:.15e}"
        first_digit = sci_str[0]

        if first_digit in ("1", "2"):
            return base_sig_figs + 1

        return base_sig_figs

    @classmethod
    def format_scientific(cls, value, sig_figs):
        r"""
        Render a number as LaTeX scientific notation: m \cdot 10^{n}.

        NOTES:
          * Decimal separator is a comma (Russian locale).
          * Trailing zeros in the mantissa are trimmed.
          * Exponent 0 is elided (returns just the mantissa).
        """
        if value == 0:
            return "0"

        sign = "-" if value < 0 else ""
        abs_value = abs(value)

        decimals = max(sig_figs - 1, 0)

        sci_str = f"{abs_value:.{decimals}e}"
        mantissa_part, exp_part = sci_str.split("e")
        exponent = int(exp_part)
        mantissa_str = mantissa_part

        # Trim insignificant zeros in the mantissa.
        if "." in mantissa_str:
            mantissa_str = mantissa_str.rstrip("0").rstrip(".")

        if not mantissa_str:
            mantissa_str = "0"

        mantissa_str = mantissa_str.replace(".", ",")

        if exponent == 0:
            return sign + mantissa_str

        return (
            rf"{sign}{mantissa_str}"
            rf" \cdot 10^{{{exponent}}}"
        )

    @classmethod
    def format_num(cls, num_str):
        """
        Format a raw numeric string from XML into a display-ready string.

        Handles:
          * Zero → "0".
          * Optional scientific notation (if sci_threshold is set).
          * Fallback to the input string if parsing fails.

        NOTE: `<real>` elements contain the raw decimal in C locale, e.g.
        "-0.49999999999999978" or "1.7347234759768071E-18". We parse with
        Python's float() (which understands "E-18") and re-emit with a
        comma decimal separator.
        """
        if not num_str:
            return ""

        try:
            value = float(num_str)

            if value == 0:
                return "0"

            abs_value = abs(value)

            # Route very small / very large magnitudes through the
            # scientific formatter when the user enabled it.
            if cls.sci_threshold is not None:
                threshold = 10.0 ** cls.sci_threshold

                if abs_value < threshold:
                    target_sig_figs = cls._compute_target_sig_figs(
                        abs_value
                    )
                    return cls.format_scientific(
                        value, target_sig_figs
                    )

            target_sig_figs = cls._compute_target_sig_figs(
                abs_value
            )

            formatted = f"{value:.{target_sig_figs}g}"

            # Python's %g may fall back to exponent form (e.g. 1.7e-18).
            # That form is not LaTeX-friendly, so expand it.
            if "e" in formatted.lower():
                rounded_value = float(formatted)

                result = (
                    f"{rounded_value:.20f}"
                    .rstrip("0")
                    .rstrip(".")
                )

                if not result:
                    result = "0"
            else:
                result = formatted

            return result.replace(".", ",")

        except (ValueError, TypeError):
            return str(num_str).replace(".", ",")

    @classmethod
    def format_angle_value(cls, angle):
        """
        Render a polar-form angle with a fixed number of decimals.

        NOTE: Angles in Mathcad results often carry 15+ digits of
        floating-point noise (e.g. 81.6078831602174). Rounding to one
        decimal keeps the document readable. "Negative zero" is coerced
        to exactly zero.
        """
        try:
            value = float(angle)
        except (ValueError, TypeError):
            return str(angle).replace(".", ",")

        if math.isclose(value, 0.0, abs_tol=1e-12):
            value = 0.0

        text = f"{value:.{cls.ANGLE_DECIMALS}f}"

        return text.replace(".", ",")

    # -------------------------------------------------------------------
    # LaTeX escaping
    # -------------------------------------------------------------------

    @staticmethod
    def escape_latex(text):
        """
        Escape LaTeX special characters in plain text.

        NOTE: Used only for identifiers and string results that end up
        inside \\text{...}. Mathcad identifiers frequently contain
        characters like "_", "%", "&", "#" that would otherwise break
        LaTeX compilation.
        """
        replacements = {
            "\\": r"\textbackslash{}",
            "{": r"\{",
            "}": r"\}",
            "_": r"\_",
            "%": r"\%",
            "&": r"\&",
            "#": r"\#",
            "$": r"\$",
            "^": r"\textasciicircum{}",
            "~": r"\textasciitilde{}",
        }

        return "".join(
            replacements.get(char, char)
            for char in str(text)
        )

    # -------------------------------------------------------------------
    # Identifier / string parsing
    # -------------------------------------------------------------------

    @classmethod
    def parse_identifier(cls, node, function=False):
        """
        Render a Mathcad identifier (variable or function name) as LaTeX.

        NOTE (subscripts):
            Mathcad 14/15 encodes subscripts in TWO ways:
              1. As an XML attribute `subscript="0"` on the <id> node.
                 Example: <ml:id subscript="0">Xw1</ml:id> → Xw1_0.
              2. Inside the text itself, e.g. "R.max".
            Both forms appear in real worksheets. We support both.

        NOTE (function flag):
            When a bare <id> is used as a function (e.g. `pp(z)`), we
            wrap it in \\mathrm{...} so it renders upright instead of
            italic — matching the visual look of Mathcad user functions.
        """
        if node is None:
            return "?"

        base = (node.text or "").strip()
        subscript = None

        # Form 1: attribute-based subscript.
        for key, value in node.attrib.items():
            if cls.strip_ns(key) == "subscript":
                subscript = value
                break

        # Form 2: child-element subscript or fallback id.
        for child in node:
            tag = cls.strip_ns(child.tag)

            if tag in ("subscript", "sub"):
                subscript = "".join(child.itertext()).strip()

            elif tag in ("name", "id", "sym") and not base:
                base = "".join(child.itertext()).strip()

        # Form 3: dotted notation in the text itself.
        if subscript is None and "." in base:
            base, subscript = base.split(".", 1)

        base_latex = cls.escape_latex(base)

        if function:
            base_latex = rf"\mathrm{{{base_latex}}}"

        if subscript is not None and subscript != "":
            subscript_latex = cls.escape_latex(subscript)

            return (
                base_latex
                + rf"_{{\mathrm{{{subscript_latex}}}}}"
            )

        return base_latex

    @classmethod
    def parse_string(cls, node):
        """
        Render a Mathcad string literal as LaTeX.

        NOTE: Mathcad strings (from concat, num2str, etc.) can span
        multiple lines. Single-line strings become \\text{"..."}; multi-
        line strings are wrapped in a `gathered` environment with quotes
        only on the first/last line — this mirrors how Mathcad renders
        the string across lines in the worksheet.
        """
        text = "".join(node.itertext())
        text = text.replace("\r\n", "\n").replace("\r", "\n")

        lines = text.split("\n")

        if len(lines) == 1:
            escaped = cls.escape_latex(text)
            return rf'\text{{"{escaped}"}}'

        parts = [
            rf"\text{{{cls.escape_latex(line or ' ')}}}"
            for line in lines
        ]

        parts[0] = (
            r'\text{"'
            + cls.escape_latex(lines[0])
            + "}"
        )

        parts[-1] = (
            r"\text{"
            + cls.escape_latex(lines[-1])
            + '"}'
        )

        return (
            r"\begin{gathered}"
            + r" \\ ".join(parts)
            + r"\end{gathered}"
        )

    # -------------------------------------------------------------------
    # Complex number handling
    # -------------------------------------------------------------------

    @staticmethod
    def normalize_angle_degrees(angle):
        """
        Normalize an angle into the half-open interval (-180, 180].

        NOTE: atan2 returns (-180, 180]. Mathcad applies the same
        convention, so we simply clean up the boundary cases and remove
        "-0.0". Small floating-point residue near ±180 is snapped to
        exactly ±180, and near 0 to exactly 0.
        """
        if math.isclose(angle, 0.0, abs_tol=1e-12):
            return 0.0

        angle = math.fmod(angle, 360.0)

        if angle > 180.0:
            angle -= 360.0
        elif angle <= -180.0:
            angle += 360.0

        if math.isclose(angle, 0.0, abs_tol=1e-12):
            return 0.0

        if math.isclose(abs(angle), 180.0, abs_tol=1e-12):
            return 180.0

        return angle

    @classmethod
    def calculate_complex_polar(cls, real_value, imag_value):
        """
        Convert (real, imag) to (modulus, angle_in_degrees).

        NOTE: Mathcad displays complex numbers in polar form as
        `|z|∠φ°` (modulus and angle). We convert rectangular values
        to that display form.
        """
        real_value = float(real_value)
        imag_value = float(imag_value)

        modulus = math.hypot(real_value, imag_value)

        angle = math.degrees(
            math.atan2(imag_value, real_value)
        )

        angle = cls.normalize_angle_degrees(angle)

        return modulus, angle

    @staticmethod
    def get_imag_symbol(node):
        """
        Get the imaginary-unit symbol ("i" or "j") from an <imag> node.

        NOTE: Mathcad lets users choose between "i" and "j" for the
        imaginary unit. The symbol is stored as an XML attribute, so a
        worksheet can technically mix both notations. We respect the
        attribute per-node rather than forcing one global choice.
        """
        if node is None:
            return "i"

        return (node.attrib.get("symbol") or "i").strip() or "i"

    @classmethod
    def parse_complex(cls, real_value, imag_value, imag_symbol="i"):
        r"""
        Render a rectangular complex number as polar: |z|\angle\phi^\circ.

        NOTE: This is Mathcad's "polar display" convention for complex
        results. If the rectangular form fails to parse (e.g. symbolic
        results contain variables), we fall back to a + bj display.
        """
        try:
            modulus, angle = cls.calculate_complex_polar(
                real_value, imag_value
            )
        except (ValueError, TypeError):
            real_text = cls.format_num(str(real_value))
            imag_text = cls.format_num(str(imag_value))

            sign = "" if imag_text.startswith("-") else "+"

            return f"{real_text}{sign}{imag_text}{imag_symbol}"

        modulus_text = cls.format_num(str(modulus))
        angle_text = cls.format_angle_value(angle)

        return (
            rf"{modulus_text}"
            rf"\angle"
            rf"{angle_text}^\circ"
        )

    # -------------------------------------------------------------------
    # Detection of complex-exponential notation (∠ in exponent form)
    # -------------------------------------------------------------------

    @staticmethod
    def is_degree_unit(node):
        """
        Detect <ml:id>deg</ml:id> — the degree "unit" pseudo-variable.

        NOTE: Mathcad treats `deg` as a unit-like constant whose value
        is π/180. In complex exponential notation Mathcad writes
        `e^(1j·deg·φ)` to mean `e^(jφ)`. We detect this pattern and
        convert to the angle form ∠φ°.
        """
        if node is None:
            return False

        if MathcadParser.strip_ns(node.tag) != "id":
            return False

        return (node.text or "").strip() == "deg"

    @classmethod
    def is_euler_base(cls, node):
        """
        Detect the Euler constant `e` used as the base of a power.

        NOTE: `e` is encoded either as <ml:e/> (older format) or as
        <ml:id>e</ml:id> (14/15). The check for a subscript attribute
        is important: `e` with a subscript is a *user variable*, not
        Euler's number.
        """
        if node is None:
            return False

        tag = cls.strip_ns(node.tag)

        if tag == "e":
            return True

        if tag in ("id", "sym"):
            base = (node.text or "").strip()

            for key in node.attrib:
                if cls.strip_ns(key) == "subscript":
                    return False

            return base == "e"

        return False

    @classmethod
    def extract_polar_angle_degrees(cls, node):
        """
        Recognize the exponent form `1j·deg·φ` and return the LaTeX for φ.

        WHY THIS EXISTS
        ---------------
        Mathcad writes complex exponentials as:

            <apply><pow/>
                <id>e</id>
                <apply><mult/>
                    <apply><mult/>
                        <imag symbol="j">1</imag>
                        <id>deg</id>
                    </apply>
                    <real>84</real>          <-- the actual angle
                </apply>
            </apply>

        We detect this tree-shape and return the LaTeX for the angle
        operand ("84"). If the shape is not exactly what we expect
        (e.g. the user wrote `e^(x+y)` symbolically), we return None
        so the generic `pow` handler takes over.

        The check for `i_deg_found` ensures we only accept the pattern
        once; extra factors (e.g. a numeric coefficient) are allowed
        and appended.
        """
        if node is None:
            return None

        children = list(node)

        if not children:
            return None

        op = cls.strip_ns(children[0].tag)

        if op != "mult":
            return None

        operands = children[1:]

        if len(operands) < 2:
            return None

        def operand_is_i_times_deg(operand):
            """
            Helper: is the operand `1j·deg` (imag × deg)?
            """
            parts = list(operand)

            if not parts:
                return False

            if cls.strip_ns(parts[0].tag) != "mult":
                return False

            factors = parts[1:]

            has_imag = any(
                cls.strip_ns(p.tag) == "imag" for p in factors
            )

            has_deg = any(
                cls.is_degree_unit(p) for p in factors
            )

            return has_imag and has_deg

        i_deg_found = False
        angle_parts = []

        for operand in operands:
            if operand_is_i_times_deg(operand):
                if i_deg_found:
                    # More than one `j·deg` factor — unusual, bail out.
                    return None
                i_deg_found = True
                continue

            latex = cls.parse_node_to_latex(operand)

            if latex:
                angle_parts.append(latex)

        if not i_deg_found or not angle_parts:
            return None

        if len(angle_parts) == 1:
            return angle_parts[0]

        return " \\cdot ".join(angle_parts)

    @classmethod
    def format_angle_symbol(cls, angle_latex):
        r"""
        Wrap an angle in the LaTeX angle marker: \angle φ°.

        NOTE: An empty angle part (shouldn't happen in practice)
        becomes \\angle^\\circ — the bare angle symbol.
        """
        if not angle_latex:
            return r"\angle^{\circ}"

        return rf"\angle{angle_latex}^{{\circ}}"

    @classmethod
    def parse_complex_node(cls, node):
        """
        Parse a <complex> element: <real>r</real> <imag>m</imag>.

        NOTE: This is the *result* form Mathcad uses when the numeric
        evaluator produces a complex number. It's different from the
        input form `a + bj` (see the `plus`/`minus` handling in
        parse_node_to_latex).
        """
        children = list(node)

        real_node = next(
            (
                child for child in children
                if cls.strip_ns(child.tag) == "real"
            ),
            None
        )

        imag_node = next(
            (
                child for child in children
                if cls.strip_ns(child.tag) == "imag"
            ),
            None
        )

        real_value = (
            real_node.text.strip()
            if real_node is not None and real_node.text
            else "0"
        )

        imag_value = (
            imag_node.text.strip()
            if imag_node is not None and imag_node.text
            else "0"
        )

        imag_symbol = cls.get_imag_symbol(imag_node)

        return cls.parse_complex(
            real_value,
            imag_value,
            imag_symbol
        )

    # -------------------------------------------------------------------
    # Angle-mark machinery (deferred ∠ substitution)
    # -------------------------------------------------------------------

    # Placeholder regex: matches "@@ANGLE0@@", "@@ANGLE1@@", etc.
    ANGLE_MARK_RE = re.compile(r"@@ANGLE\d+@@")

    # Stack of currently-active marks (a list of one-key dicts).
    _angle_mark_stack = []

    # Monotonic counter so nested angles never collide.
    _angle_mark_counter = 0

    @classmethod
    def push_angle_mark(cls, latex):
        r"""
        Hide a ready-made LaTeX angle (e.g. \angle 84^\circ) inside a
        placeholder token "@@ANGLEn@@".

        WHY THIS IS NEEDED
        ------------------
        When we encounter an exponent `e^(j·deg·84)`, we want to emit
        `∠84°`. But the parent `mult` handler still needs to see the
        result as an opaque *operand* — it does things like "insert 1
        before an angle if no coefficient is present". If we emitted
        raw LaTeX with an angle symbol, the mult handler would try to
        parse and reformat it.

        The placeholder approach:
          1. Emit a token "@@ANGLEn@@" that contains no LaTeX special
             characters. It survives any string manipulation safely.
          2. When the outermost parser frame completes, walk the final
             string and substitute every placeholder for the real
             angle LaTeX — inserting an implicit "1" modulus when no
             numeric coefficient precedes.
        """
        mark = f"@@ANGLE{cls._angle_mark_counter}@@"
        cls._angle_mark_counter += 1
        cls._angle_mark_stack.append({mark: latex})
        return mark

    @classmethod
    def pop_all_angle_marks(cls):
        """
        Merge and clear every active angle placeholder.

        NOTE: Returns a flat dict {mark: latex} so the caller can
        perform one pass of substitutions over the final string.
        """
        merged = {}

        for marks in cls._angle_mark_stack:
            merged.update(marks)

        cls._angle_mark_stack.clear()

        return merged

    @staticmethod
    def _needs_unit_prefix(before_text):
        r"""
        Decide whether an implicit "1" modulus must be inserted before
        an angle symbol, based on the LaTeX text that precedes it.

        Heuristic:
          * If the preceding text ends with a digit, letter, '}', ')'
            or ']' — a coefficient (or a complex expression) is
            already present. Do NOT add "1".
          * Otherwise (empty context, or last char is an operator like
            '+', '-', '=', '(', '{', '[', or '·') — there is no
            coefficient. Add "1" so the reader sees `1∠φ°`.

        Example cases:
          "0,9987 + "   → last char is space → add "1"
          "\\frac{1}{"  → last char is '{' → add "1"
          "2 · "        → last char is space → add "1"
          "1,5"         → last char is digit → no prefix
          "x}"          → last char is '}' → no prefix
        """
        stripped = before_text.rstrip()

        if not stripped:
            return True

        last = stripped[-1]

        if last.isdigit() or last.isalpha() or last in "})]":
            return False

        return True

    @classmethod
    def restore_angle_symbols(cls, text, marks):
        r"""
        Substitute every "@@ANGLEn@@" placeholder for its real
        \angle... LaTeX, inserting a leading "1" modulus when the
        surrounding context lacks a coefficient.

        NOTE: This is THE single place where the implicit-unity rule
        is enforced. All earlier stages just emit opaque placeholders.
        """
        for mark, angle_latex in marks.items():
            # Special case: the string "\angle@@ANGLEn@@" means the
            # angle symbol was already emitted alongside the mark.
            # Collapse it to "1\angle..." regardless of context.
            bare = rf"\angle{mark}"
            while bare in text:
                text = text.replace(bare, rf"1{angle_latex}", 1)

            # Normal placeholder substitution.
            while mark in text:
                idx = text.find(mark)
                before = text[:idx]

                if cls._needs_unit_prefix(before):
                    replacement = "1" + angle_latex
                else:
                    replacement = angle_latex

                text = (
                    text[:idx]
                    + replacement
                    + text[idx + len(mark):]
                )

        return text

    # -------------------------------------------------------------------
    # Detection of the "unit factor" `1`
    # -------------------------------------------------------------------

    @staticmethod
    def is_unit_factor(node):
        """
        Detect whether a node represents the literal number 1.

        NOTE: Mathcad ALWAYS writes `1·e^(jφ)` — it inserts an explicit
        `1` as the modulus. In the XML this shows up either as
        <real>1</real> or as <complex><real>1</real></complex>.

        The `_process_mult_angle_args` function uses this information to
        merge that explicit 1 into the angle output, producing
        `1∠φ°` without an extra `· 1`.
        """
        if node is None:
            return False

        tag = MathcadParser.strip_ns(node.tag)

        if tag == "real":
            return (node.text or "").strip() == "1"

        if tag == "complex":
            real_node = next(
                (
                    child for child in node
                    if MathcadParser.strip_ns(child.tag) == "real"
                ),
                None
            )

            imag_node = next(
                (
                    child for child in node
                    if MathcadParser.strip_ns(child.tag) == "imag"
                ),
                None
            )

            real_ok = (
                real_node is not None
                and (real_node.text or "").strip() == "1"
            )

            imag_empty = (
                imag_node is None
                or not (imag_node.text or "").strip()
                or (imag_node.text or "").strip() == "0"
            )

            return real_ok and imag_empty

        return False

    @staticmethod
    def is_complex_latex(latex):
        """
        Rough check: does this LaTeX fragment represent a complex number?

        NOTE: Used only to decide whether to shrink matrix columns so
        that polar-form entries stay readable (they are much wider than
        rectangular entries).
        """
        return r"\angle" in latex or latex.rstrip().endswith("i")

    # -------------------------------------------------------------------
    # mult-operator post-processing
    # -------------------------------------------------------------------

    @staticmethod
    def _is_bare_angle_mark(text):
        """
        True if `text` is an angle mark without surrounding parens.

        NOTE: This is called on already-processed fragments, so the
        placeholder token has already been replaced by real LaTeX.
        """
        stripped = text.strip()
        if MathcadParser.ANGLE_MARK_RE.fullmatch(stripped):
            return True
        if re.fullmatch(r"1?\\angle[^()]*", stripped):
            return True
        return False

    @staticmethod
    def _is_angle_mark_wrapped(text):
        """
        True if `text` is an angle fragment, optionally wrapped in parens.
        """
        stripped = text.strip()
        if MathcadParser.ANGLE_MARK_RE.fullmatch(stripped):
            return True
        if re.fullmatch(
            r"(?:\\left\()?1?\\angle[^)]*(?:\\right\))?",
            stripped
        ):
            return True
        return False

    @staticmethod
    def _is_pure_number_latex(text):
        """True if text is purely a number (possibly signed, with comma)."""
        return bool(re.match(r"^[+-]?[\d,.]+$", text.strip()))

    @staticmethod
    def _is_simple_coefficient(text):
        """True for a plain number or a plain identifier (with optional subscript)."""
        stripped = text.strip()
        if not stripped:
            return False
        if re.fullmatch(r"[+-]?[\d,.]+", stripped):
            return True
        if re.fullmatch(r"[A-Za-z][A-Za-z0-9]*(_\{[^}]*\})?", stripped):
            return True
        return False

    @staticmethod
    def _ends_with_digit(text):
        """True if `text` ends with a digit (looking past trailing braces/spaces)."""
        stripped = text.strip().rstrip("} \t")
        return bool(stripped) and stripped[-1].isdigit()

    @classmethod
    def _process_mult_angle_args(cls, args, children):
        r"""
        Post-process a chain of `mult` operands to produce clean LaTeX.

        MATHCAD QUIRK EXPLAINED
        -----------------------
        Mathcad's XML for multiplication is fully explicit. For example
        `3·a·∠120°·b` becomes a flat <apply><mult/> ... </apply> with
        one <apply> per factor. The XML does not tell us whether the
        `1` in `1∠120°` is a *modulus* or just a redundant unit.

        This function:
          1. Merges the explicit `1` in `1 · ∠120°` so we don't emit
             "1 · 1∠120°".
          2. Decides where an implicit `1` modulus is needed (e.g. when
             an angle is preceded by `+` or `(`).
          3. Joins everything with `\cdot` — except in the special case
             where a coefficient directly precedes an angle, where the
             implicit product is written juxtaposed.
        """
        n = len(args)

        # First pass: gather every operand's metadata.
        factors = []
        for i in range(n):
            arg = args[i]
            node = children[1 + i] if 1 + i < len(children) else None

            factors.append({
                "text": arg,
                "node": node,
                "is_angle": cls._is_angle_mark_wrapped(arg),
                "is_bare_angle": cls._is_bare_angle_mark(arg),
                "is_unit_one": cls.is_unit_factor(node),
                "is_pure_number": cls._is_pure_number_latex(arg),
                "has_one_angle": bool(
                    re.search(r"(?<!\d)1\\angle", arg)
                ),
                "keep": True,
                "prefix_one": False,
            })

        # Step 1: merge the explicit 1 that appears immediately before
        # an angle. Because Mathcad always writes `1·e^(jφ)` the XML
        # contains a literal <real>1</real> right before the angle
        # mark. We drop it and set prefix_one so the angle becomes
        # "1∠φ°" instead of "1 · 1∠φ°".
        for i, f in enumerate(factors):
            if not f["is_angle"] or f["has_one_angle"]:
                continue

            merge_idx = -1
            for j in range(i - 1, -1, -1):
                if not factors[j]["text"].strip():
                    continue
                if factors[j]["is_unit_one"]:
                    merge_idx = j
                break

            if merge_idx != -1:
                factors[merge_idx]["keep"] = False
                f["prefix_one"] = True

        # Step 2: decide where an implicit "1" modulus is required.
        #
        # Cases:
        #   * nothing before  → add 1
        #   * pure number before → no (number IS the modulus)
        #   * identifier before:
        #       - ends with digit → add 1 (e.g. "Xw1" is not a modulus)
        #       - otherwise → no
        #   * any other expression (parens, etc.) → add 1
        for i, f in enumerate(factors):
            if not f["is_angle"] or f["has_one_angle"]:
                continue
            if f["prefix_one"]:
                continue

            prev_idx = -1
            for j in range(i - 1, -1, -1):
                if not factors[j]["text"].strip():
                    continue
                if factors[j]["is_unit_one"]:
                    continue
                prev_idx = j
                break

            if prev_idx == -1:
                f["prefix_one"] = True
                continue

            prev_text = factors[prev_idx]["text"].strip()

            if cls._is_pure_number_latex(prev_text):
                pass
            elif cls._is_simple_coefficient(prev_text):
                if cls._ends_with_digit(prev_text):
                    f["prefix_one"] = True
            else:
                f["prefix_one"] = True

        # Step 3: assemble the final string.
        parts = []
        for f in factors:
            if not f["keep"]:
                continue
            text = f["text"]
            if f["is_angle"] and f["prefix_one"]:
                text = "1" + text
            parts.append((text, f))

        result = ""
        for text, f in parts:
            if not result:
                result = text
                continue

            # If the current piece is a bare angle mark and the previous
            # piece is a simple coefficient, juxtapose them (a·∠x →
            # a∠x). Otherwise insert an explicit \cdot.
            if cls._is_bare_angle_mark(text):
                prev_clean = result.strip()
                if cls._is_simple_coefficient(prev_clean):
                    result += text
                else:
                    result += f" \\cdot {text}"
            else:
                result += f" \\cdot {text}"

        return result if result else "1"

    # -------------------------------------------------------------------
    # Matrix rendering
    # -------------------------------------------------------------------

    MATRIX_MAX_WIDTH = 1000

    @classmethod
    def build_matrix_latex(cls, rows, cols, delimiter="bmatrix"):
        r"""
        Build a LaTeX matrix from already-parsed cell strings.

        NOTE: For very wide matrices (mostly those containing complex
        polar entries) we split into multiple `pmatrix` blocks joined by
        `\quad`. Word's equation renderer cannot handle arbitrarily wide
        matrices and produces broken layout otherwise.
        """
        if not rows or not cols:
            return ""

        total_width = sum(
            max(len(cell) for cell in column)
            for column in zip(*rows)
        ) + 2 * (len(cols) - 1)

        if total_width <= cls.MATRIX_MAX_WIDTH:
            body = r" \\ ".join(
                " & ".join(row)
                for row in rows
            )
            return rf"\begin{{{delimiter}}}{body}\end{{{delimiter}}}"

        # Otherwise: split into column groups that each fit.
        groups = []
        current_group = []
        current_width = 0

        for index, width in enumerate(cols):
            addition = width + (2 if current_group else 0)

            if (
                current_group
                and current_width + addition > cls.MATRIX_MAX_WIDTH
            ):
                groups.append(current_group)
                current_group = [index]
                current_width = width
            else:
                current_group.append(index)
                current_width += addition

        if current_group:
            groups.append(current_group)

        parts = []

        for group in groups:
            body = r" \\ ".join(
                " & ".join(row[index] for index in group)
                for row in rows
            )
            parts.append(
                rf"\begin{{pmatrix}}{body}\end{{pmatrix}}"
            )

        return r" \quad ".join(parts)

    @classmethod
    def parse_matrix_node(cls, node, children=None):
        """
        Convert a <matrix> element into a LaTeX matrix.

        MATHCAD QUIRK
        -------------
        Mathcad's matrix XML stores cells as a flat list of child
        elements (row-major order) and the shape separately in the
        `rows`/`cols` attributes. Older files may omit those attributes
        — in that case we guess a single-column matrix.
        """
        if children is None:
            children = list(node)

        if not children:
            return ""

        try:
            rows_count = int(node.attrib.get("rows", 0))
            cols_count = int(node.attrib.get("cols", 0))
        except (TypeError, ValueError):
            rows_count = 0
            cols_count = 0

        cell_latex_list = [
            cls.parse_node_to_latex(child) or "0"
            for child in children
        ]

        # Fallback for malformed shape metadata.
        if rows_count <= 0 or cols_count <= 0:
            rows_count = len(cell_latex_list)
            cols_count = 1

        needed = rows_count * cols_count

        if len(cell_latex_list) < needed:
            cell_latex_list.extend(
                ["0"] * (needed - len(cell_latex_list))
            )
        elif len(cell_latex_list) > needed:
            cols_count = (
                len(cell_latex_list) + rows_count - 1
            ) // rows_count
            needed = rows_count * cols_count

            if len(cell_latex_list) < needed:
                cell_latex_list.extend(
                    ["0"] * (needed - len(cell_latex_list))
                )

        rows = [
            cell_latex_list[
                i * cols_count:(i + 1) * cols_count
            ]
            for i in range(rows_count)
        ]

        col_widths = [
            max(len(cell[i]) for cell in rows)
            for i in range(cols_count)
        ]

        # Complex entries (with angle markers) tend to be very wide, so
        # we shrink the effective max width to force column-splitting.
        has_complex = any(
            cls.is_complex_latex(cell)
            for row in rows
            for cell in row
        )

        if has_complex:
            saved_max_width = cls.MATRIX_MAX_WIDTH
            cls.MATRIX_MAX_WIDTH = min(col_widths)
            try:
                return cls.build_matrix_latex(
                    rows, col_widths, delimiter="bmatrix"
                )
            finally:
                cls.MATRIX_MAX_WIDTH = saved_max_width

        return cls.build_matrix_latex(
            rows, col_widths, delimiter="bmatrix"
        )

    # -------------------------------------------------------------------
    # Top-level node dispatcher
    # -------------------------------------------------------------------

    @classmethod
    def parse_node_to_latex_root(cls, node):
        """
        Entry point for parsing a single expression node.

        NOTE: This is the outermost wrapper. Its only job beyond
        delegating to `parse_node_to_latex` is to flush any lingering
        angle placeholders. Inside recursive parsing we deliberately
        keep placeholders unresolved so that the parent `mult` /
        `plus` handlers can see them as opaque tokens.
        """
        try:
            latex = cls.parse_node_to_latex(node)
        finally:
            marks = cls.pop_all_angle_marks()

        if marks:
            latex = cls.restore_angle_symbols(latex, marks)

        # Safety net: if some nested call leaked a placeholder we clean
        # it up here.
        if "@@ANGLE" in latex:
            extra_marks = cls.pop_all_angle_marks()
            if extra_marks:
                latex = cls.restore_angle_symbols(
                    latex, extra_marks
                )

        return latex

    @classmethod
    def parse_node_to_latex(cls, node):
        """
        Recursive XML → LaTeX dispatcher for a single Mathcad node.

        The tag name identifies the construct. The list below maps
        Mathcad XML tags to their LaTeX equivalents:

            math       → whitespace-joined children
            real       → number literal
            id / sym   → identifier (variable or function name)
            str/string → string literal
            complex    → <real>, <imag> complex-number result
            result / symResult → wrapper elements around numeric or
                                 symbolic results
            imag       → imaginary unit with coefficient
            matrix     → matrix
            parens     → explicit parentheses
            apply      → operator application (first child is the op)
            function   → function definition with bound variables
            sequence   → comma-separated list
            placeholder→ empty square

        NOTE: `apply` is the workhorse. Mathcad encodes EVERY operator
        (+, -, ·, ÷, ^, √, |·|, conjugate, user-function-call, ...) as
        <apply> with the operator as the first child. See the branches
        inside `if tag == "apply":`.
        """
        if node is None:
            return "?"

        tag = cls.strip_ns(node.tag)
        children = list(node)

        # --- math: top-level container inside a <region> --------------
        if tag == "math":
            return " ".join(
                cls.parse_node_to_latex(child)
                for child in children
            )

        # --- real number literal --------------------------------------
        if tag == "real":
            return cls.format_num(
                node.text.strip() if node.text else ""
            )

        # --- identifier / symbol --------------------------------------
        if tag in ("id", "sym"):
            return cls.parse_identifier(node)

        # --- string literal -------------------------------------------
        if tag in ("str", "string"):
            return cls.parse_string(node)

        # --- complex number result ------------------------------------
        if tag == "complex":
            return cls.parse_complex_node(node)

        # --- result / symResult wrappers ------------------------------
        # These appear inside <eval> / <symEval> to hold numeric or
        # symbolic output. We just render their inner content.
        if tag in ("result", "symResult"):
            parts = [
                cls.parse_node_to_latex(child)
                for child in children
            ]
            parts = [part for part in parts if part and part != "?"]
            if parts:
                return " ".join(parts)
            if node.text and node.text.strip():
                return cls.parse_string(node)
            return ""

        # --- imaginary unit -------------------------------------------
        # MATHCAD QUIRK: Mathcad writes the imaginary unit as
        #   <imag symbol="j">1</imag>
        # where the text content is the coefficient (usually 1) and the
        # attribute carries the letter (i or j). We DROP a coefficient
        # of exactly 1 because `1j` reads awkwardly in LaTeX — the bare
        # symbol alone is the correct notation.
        if tag == "imag":
            symbol = node.attrib.get("symbol", "i")
            raw = node.text.strip() if node.text else ""
            number = cls.format_num(raw)

            if not raw or number == "1":
                # Coeff is 1 (or omitted) → bare imaginary unit.
                return symbol
            if number == "-1":
                return f"-{symbol}"

            return f"{number}{symbol}"

        # --- matrix ---------------------------------------------------
        if tag == "matrix":
            return cls.parse_matrix_node(node, children)

        # --- explicit parentheses -------------------------------------
        if tag == "parens":
            child_latex = (
                cls.parse_node_to_latex(children[0])
                if children else ""
            )
            return rf"\left({child_latex}\right)"

        # --- operator application -------------------------------------
        # This is the BIG branch. The first child is the operator
        # (a self-closing element like <mult/>), the remaining children
        # are the operands.
        if tag == "apply":
            if not children:
                return ""

            op_node = children[0]
            op = cls.strip_ns(op_node.tag)

            # ---- Special-case: a + bj → polar form ----
            #
            # Mathcad evaluates `a + b·j` with numeric a, b into a
            # complex result. We intercept this shape and produce the
            # polar display `|z|∠φ°` instead of `a + bj`.
            if op == "plus" and len(children) >= 3:
                first_node = children[1]
                second_node = children[2]
                first_tag = cls.strip_ns(first_node.tag)
                second_tag = cls.strip_ns(second_node.tag)

                if first_tag == "real" and second_tag == "imag":
                    real_value = (
                        first_node.text.strip()
                        if first_node.text else "0"
                    )
                    imag_value = (
                        second_node.text.strip()
                        if second_node.text else "0"
                    )
                    imag_symbol = cls.get_imag_symbol(second_node)
                    return cls.parse_complex(
                        real_value, imag_value, imag_symbol
                    )

            # ---- Special-case: a - bj → polar form ----
            if op == "minus" and len(children) >= 3:
                first_node = children[1]
                second_node = children[2]
                first_tag = cls.strip_ns(first_node.tag)
                second_tag = cls.strip_ns(second_node.tag)

                if first_tag == "real" and second_tag == "imag":
                    real_value = (
                        first_node.text.strip()
                        if first_node.text else "0"
                    )
                    imag_value_raw = (
                        second_node.text.strip()
                        if second_node.text else "0"
                    )
                    try:
                        imag_value = str(-float(imag_value_raw))
                    except (ValueError, TypeError):
                        imag_value = f"-({imag_value_raw})"

                    imag_symbol = cls.get_imag_symbol(second_node)
                    return cls.parse_complex(
                        real_value, imag_value, imag_symbol
                    )

            # Recursively parse every operand (everything after the op).
            args = [
                cls.parse_node_to_latex(child)
                for child in children[1:]
            ]

            # ---- multiplication ----
            if op == "mult":
                result = cls._process_mult_angle_args(
                    args, children
                )
                # Flush angle placeholders introduced by nested pow()
                # calls (which emit @@ANGLEn@@ tokens).
                marks_to_restore = cls.pop_all_angle_marks()
                return cls.restore_angle_symbols(
                    result, marks_to_restore
                )

            # ---- division ----
            if op == "div":
                if len(args) > 1:
                    return (
                        f"\\frac{{{args[0]}}}"
                        f"{{{args[1]}}}"
                    )
                return (
                    f"\\frac{{{args[0]}}}{{?}}"
                    if args else ""
                )

            # ---- addition ----
            if op == "plus":
                if len(args) > 1:
                    return f"{args[0]} + {args[1]}"
                return args[0] if args else ""

            # ---- subtraction ----
            if op == "minus":
                if len(args) > 1:
                    return f"{args[0]} - {args[1]}"
                return f"-{args[0]}" if args else "-"

            # ---- unary negation ----
            # MATHCAD QUIRK: Mathcad sometimes nests a long chain of
            # <neg/> around a placeholder (visible in empty-input
            # regions). We render the outermost minus and recurse.
            if op == "neg":
                return f"-{args[0]}" if args else "-"

            # ---- exponentiation ----
            # Intercepts e^(j·deg·φ) and produces an angle mark; other
            # powers fall through to plain `{base}^{exp}`.
            if op == "pow":
                base_node = (
                    children[1] if len(children) > 1 else None
                )
                exp_node = (
                    children[2] if len(children) > 2 else None
                )

                if cls.is_euler_base(base_node):
                    angle_latex = cls.extract_polar_angle_degrees(
                        exp_node
                    )
                    if angle_latex is not None:
                        return cls.push_angle_mark(
                            cls.format_angle_symbol(angle_latex)
                        )

                if len(args) > 1:
                    return f"{{{args[0]}}}^{{{args[1]}}}"
                return (
                    f"{{{args[0]}}}^{{?}}"
                    if args else ""
                )

            # ---- square root ----
            if op == "sqrt":
                return (
                    f"\\sqrt{{{args[0]}}}"
                    if args else "\\sqrt{?}"
                )

            # ---- absolute value ----
            if op == "absval":
                return (
                    f"\\left| {args[0]} \\right|"
                    if args else "\\left| ? \\right|"
                )

            # ---- complex conjugate ----
            if op == "conjugate":
                return (
                    f"\\overline{{{args[0]}}}"
                    if args else "\\overline{?}"
                )

            # ---- indexing / subscript access ----
            if op == "indexer":
                if not args:
                    return ""
                index_latex = ", ".join(args[1:])
                return rf"{args[0]}_{{\text[{index_latex}]}}"

            # ---- matrix transpose ----
            if op == "transpose":
                return (
                    f"{{{args[0]}}}^{{T}}"
                    if args else "^{T}"
                )

            # ---- equality (used inside symbolic output) ----
            if op == "equal":
                if len(args) > 1:
                    return f"{args[0]} = {args[1]}"
                return f"{args[0]} = ?" if args else "="

            # ---- user-defined function call ----
            # If the operator itself is an <id>, we're looking at a
            # function application: pp(z), arg(z), etc.
            if op in ("id", "sym"):
                function_name = cls.parse_identifier(
                    op_node, function=True
                )
                return (
                    rf"{function_name}\left("
                    + ", ".join(args)
                    + r"\right)"
                )

            # Unknown operator — emit a placeholder so the doc still
            # compiles.
            return "?"

        # --- function definition (f(x) := ...) ------------------------
        if tag == "function":
            bound_vars = next(
                (
                    child for child in children
                    if cls.strip_ns(child.tag) == "boundVars"
                ),
                None
            )
            function_node = next(
                (
                    child for child in children
                    if cls.strip_ns(child.tag) in ("id", "sym")
                ),
                None
            )
            function_latex = (
                cls.parse_identifier(
                    function_node, function=True
                )
                if function_node is not None else ""
            )
            variables_latex = (
                ", ".join(
                    cls.parse_node_to_latex(child)
                    for child in bound_vars
                )
                if bound_vars is not None else ""
            )
            return (
                rf"{function_latex}\left("
                + variables_latex
                + r"\right)"
            )

        # --- comma-separated sequence ---------------------------------
        # Appears as the argument list of concat() and similar.
        if tag == "sequence":
            return ", ".join(
                cls.parse_node_to_latex(child)
                for child in children
            )

        # --- empty placeholder ----------------------------------------
        # Renders as the LaTeX "square" symbol.
        if tag == "placeholder":
            return r"\square"

        # --- fallback: try each child in order ------------------------
        # Unknown wrapper — descend until we find something renderable.
        if children:
            for child in children:
                result = cls.parse_node_to_latex(child)
                if result and result != "?":
                    return result

        return ""

    # -------------------------------------------------------------------
    # eval / symEval handling
    # -------------------------------------------------------------------

    @classmethod
    def parse_eval_to_latex_root(cls, node):
        """
        Entry point for `<eval>` regions (numeric or symbolic evaluation).

        NOTE: Same placeholder-flushing logic as
        `parse_node_to_latex_root` — we need to resolve angle marks
        before returning to the caller.
        """
        try:
            latex = cls.parse_eval_to_latex(node)
        finally:
            marks = cls.pop_all_angle_marks()

        if marks:
            latex = cls.restore_angle_symbols(latex, marks)

        if "@@ANGLE" in latex:
            extra_marks = cls.pop_all_angle_marks()
            if extra_marks:
                latex = cls.restore_angle_symbols(
                    latex, extra_marks
                )

        return latex

    @classmethod
    def parse_eval_to_latex(cls, node):
        """
        Render an <eval> element as `expr = result` (or
        `expr = symResult = result` when symbolic evaluation is used).

        MATHCAD XML STRUCTURE
        ---------------------
        <eval>
            <expression>...</expression>          (the input)
            <command>...</command>                (e.g. explicit ALL)
            <symResult>...</symResult>            (symbolic expansion)
            <result>...</result>                  (final numeric value)
        </eval>

        Not every child is present. For purely numeric evaluations there
        is no <symEval> and no <symResult>. The parser tolerates all
        combinations.

        The output uses "=" between stages so the reader can follow the
        chain: input expression → symbolic intermediate → final value.
        """
        result_node = cls.find_first_by_tag(node, "result")
        sym_eval_node = cls.find_first_by_tag(node, "symEval")

        parts = []

        if sym_eval_node is not None:
            # Pick the first non-command, non-symResult child — that is
            # the original expression the user typed.
            expression_node = next(
                (
                    child for child in sym_eval_node
                    if cls.strip_ns(child.tag)
                    not in ("command", "symResult")
                ),
                None
            )
            sym_result_node = cls.find_first_by_tag(
                sym_eval_node, "symResult"
            )

            if expression_node is not None:
                expression_latex = cls.parse_node_to_latex(
                    expression_node
                )
                if expression_latex:
                    parts.append(expression_latex)

            if sym_result_node is not None:
                sym_result_latex = cls.parse_node_to_latex(
                    sym_result_node
                )
                if sym_result_latex:
                    parts.append(sym_result_latex)

        else:
            # Numeric-only evaluation: no symEval wrapper.
            expression_node = next(
                (
                    child for child in node
                    if cls.strip_ns(child.tag) != "result"
                ),
                None
            )
            if expression_node is not None:
                expression_latex = cls.parse_node_to_latex(
                    expression_node
                )
                if expression_latex:
                    parts.append(expression_latex)

        if result_node is not None:
            result_latex = cls.parse_node_to_latex(result_node)
            if result_latex:
                parts.append(result_latex)

        return " = ".join(parts)

    # -------------------------------------------------------------------
    # Whole-file processing
    # -------------------------------------------------------------------

    @classmethod
    def process_file(cls, filepath):
        """
        Read an .xmcd/.xml worksheet and produce Markdown + LaTeX.

        MATHCAD FILE STRUCTURE (simplified)
        -----------------------------------
        <worksheet>
            <settings>...</settings>              (fonts, precision...)
            <regions>
                <region top="..." left="...">
                    <math>...or...</math>
                    <text>...or...</text>
                    <rendering .../>
                </region>
                ...
            </regions>
            <binaryContent>...</binaryContent>    (embedded images)
        </worksheet>

        We iterate over all <region> elements in document order, sort
        them by `top` (visual reading order), and render their content.

        Text regions become plain Markdown paragraphs; math regions
        become display equations surrounded by `$$ ... $$`.
        """
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()
        except Exception as exc:
            raise ValueError(
                f"Ошибка чтения XML/XMCD файла: {exc}"
            ) from exc

        # Collect every <region> from the tree.
        regions = [
            element
            for element in root.iter()
            if cls.strip_ns(element.tag) == "region"
        ]

        # Reading-order sort. `top` is a float in points.
        def get_top(region):
            try:
                return float(region.attrib.get("top", 0))
            except (TypeError, ValueError):
                return 0.0

        regions.sort(key=get_top)

        content = []

        for region in regions:
            # --- text regions become paragraphs -----------------------
            text_node = cls.find_first_by_tag(region, "text")
            if text_node is not None:
                text_values = []
                for paragraph in text_node.iter():
                    if cls.strip_ns(paragraph.tag) == "p":
                        paragraph_text = (
                            "".join(paragraph.itertext()).strip()
                        )
                        if paragraph_text:
                            text_values.append(paragraph_text)

                if text_values:
                    content.append("\n".join(text_values) + "\n\n")
                else:
                    text_value = (
                        "".join(text_node.itertext()).strip()
                    )
                    if text_value:
                        content.append(f"{text_value}\n\n")

            # --- math regions become display equations ----------------
            math_node = cls.find_first_by_tag(region, "math")
            if math_node is None:
                continue

            for child in math_node:
                tag = cls.strip_ns(child.tag)
                latex = ""

                if tag == "define":
                    # MATHCAD SHAPE:
                    #   <define><lhs/><rhs/></define>
                    # lhs is usually an <id>; rhs may be wrapped in
                    # <eval> (numeric result attached) or be a bare
                    # expression.
                    if len(child) >= 2:
                        left_side = cls.parse_node_to_latex(child[0])
                        right_side = child[1]

                        if cls.strip_ns(right_side.tag) == "eval":
                            right_latex = (
                                cls.parse_eval_to_latex_root(right_side)
                            )
                        else:
                            right_latex = (
                                cls.parse_node_to_latex_root(right_side)
                            )

                        if right_latex:
                            latex = f"{left_side} = {right_latex}"
                        else:
                            latex = left_side

                elif tag == "eval":
                    # Standalone evaluation (no assignment).
                    latex = cls.parse_eval_to_latex_root(child)

                else:
                    # Bare expression region (rare but possible).
                    latex = cls.parse_node_to_latex_root(child)

                if latex:
                    content.append(f"$$ {latex} $$\n\n")

        return "".join(content)


# ===========================================================================
# WordConverter — Markdown + LaTeX → .docx via Pandoc
# ===========================================================================
class WordConverter:
    """
    Thin wrapper around pypandoc. Handles Pandoc bootstrap (auto-download
    if missing) and applies an optional reference DOCX for styling.
    """

    @staticmethod
    def check_pandoc(log_callback):
        """Verify Pandoc availability; download it if missing."""
        try:
            version = pypandoc.get_pandoc_version()
            log_callback(f"[OK] Обнаружен Pandoc версии: {version}")
            return True
        except OSError:
            log_callback(
                "[INFO] Pandoc не найден. "
                "Начинаю автоматическую загрузку..."
            )
            try:
                pypandoc.download_pandoc()
                log_callback(
                    "[OK] Pandoc успешно скачан и установлен!"
                )
                return True
            except Exception as exc:
                log_callback(
                    f"[ERROR] Не удалось скачать Pandoc: {exc}"
                )
                return False

    @staticmethod
    def convert_to_word(
        markdown_text, output_file, template_file, log_callback
    ):
        """
        Convert a Markdown+LaTeX string to a .docx file.

        NOTE: `--reference-doc=...` lets us reuse a Word file's styles
        (fonts, sizes, paragraph spacing) in the generated document.
        """
        try:
            log_callback(
                f"[INFO] Начинается конвертация в "
                f"'{output_file}'..."
            )

            extra_args = []

            if template_file and os.path.exists(template_file):
                extra_args.append(
                    f"--reference-doc={template_file}"
                )
                log_callback(
                    "[INFO] Применяется шаблон: "
                    f"{os.path.basename(template_file)}"
                )
            else:
                log_callback(
                    "[WARN] Шаблон стилей не найден. "
                    "Используется стандартный стиль."
                )

            pypandoc.convert_text(
                source=markdown_text,
                to="docx",
                format="markdown",
                outputfile=output_file,
                extra_args=extra_args
            )

            log_callback("[OK] Успех! Документ сохранен.")
            return True

        except Exception as exc:
            log_callback(f"[ERROR] Ошибка конвертации: {exc}")
            return False


# ===========================================================================
# ConverterApp — Tkinter GUI
# ===========================================================================
class ConverterApp:
    """
    Simple Tkinter GUI:

      * Pick an .xmcd/.xml file (or drag-and-drop it).
      * Configure rounding / scientific-notation thresholds.
      * Pick a reference .docx style template.
      * Convert to a .docx with Pandoc.
    """

    def __init__(self, root):
        self.root = root
        self.root.title("Mathcad to Word Converter (со стилями)")
        self.root.geometry("600x780")
        self.root.configure(padx=20, pady=20)

        self.input_file = None
        self.template_file = self.find_default_template()

        self.setup_ui()

        if DND_SUPPORTED:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self.on_drop)

    def find_default_template(self):
        """
        Look for a reference .docx in a "шаблон/" folder next to the
        script (or bundled executable).
        """
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        template_dir = os.path.join(base_dir, "шаблон")
        expected_template = os.path.join(
            template_dir, "Шаблон стилей.docx"
        )

        if os.path.exists(expected_template):
            return expected_template

        if os.path.isdir(template_dir):
            for filename in os.listdir(template_dir):
                if filename.lower().endswith(".docx"):
                    return os.path.join(template_dir, filename)

        # Dev fallback: an absolute path that was used during testing.
        fallback_dir = r"C:\VS code здесь\mcdx to word\шаблон"
        if os.path.isdir(fallback_dir):
            for filename in os.listdir(fallback_dir):
                if filename.lower().endswith(".docx"):
                    return os.path.join(fallback_dir, filename)

        return ""

    # -------------------------------------------------------------------
    # UI construction
    # -------------------------------------------------------------------

    def setup_ui(self):
        # ---- Input file picker ----
        frame_input = ttk.LabelFrame(
            self.root,
            text="Исходный файл Mathcad (.xmcd, .xml)",
            padding=10
        )
        frame_input.pack(fill=tk.X, pady=(0, 15))

        self.lbl_input = ttk.Label(
            frame_input,
            text="Файл не выбран",
            foreground="gray"
        )
        self.lbl_input.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(
            frame_input, text="Выбрать файл",
            command=self.browse_input
        ).pack(side=tk.RIGHT, padx=5)

        if DND_SUPPORTED:
            ttk.Label(
                frame_input,
                text="(Или перетащите файл сюда)",
                font=("Segoe UI", 8, "italic")
            ).pack(side=tk.BOTTOM, pady=5)

        # ---- Numeric formatting controls ----
        frame_settings = ttk.LabelFrame(
            self.root,
            text="Настройки округления (значащие цифры)",
            padding=10
        )
        frame_settings.pack(fill=tk.X, pady=(0, 15))

        self.sig_figs_small_var = tk.IntVar(value=4)
        self.sig_figs_large_var = tk.IntVar(value=8)

        ttk.Label(
            frame_settings,
            text="Для малых чисел и дробей (по умолчанию 4):"
        ).grid(row=0, column=0, sticky=tk.W, pady=2)

        ttk.Spinbox(
            frame_settings, from_=1, to=15,
            textvariable=self.sig_figs_small_var, width=5
        ).grid(row=0, column=1, sticky=tk.W, padx=10, pady=2)

        ttk.Label(
            frame_settings,
            text="Для больших чисел (по умолчанию 8):"
        ).grid(row=1, column=0, sticky=tk.W, pady=2)

        ttk.Spinbox(
            frame_settings, from_=1, to=20,
            textvariable=self.sig_figs_large_var, width=5
        ).grid(row=1, column=1, sticky=tk.W, padx=10, pady=2)

        # ---- Scientific-notation section ----
        ttk.Separator(
            frame_settings, orient="horizontal"
        ).grid(
            row=2, column=0, columnspan=2,
            sticky="ew", pady=(8, 6)
        )

        self.use_sci_var = tk.BooleanVar(value=False)

        ttk.Checkbutton(
            frame_settings,
            text="Научная форма для малых чисел",
            variable=self.use_sci_var,
            command=self.toggle_sci
        ).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=2)

        ttk.Label(
            frame_settings,
            text="Порог 10^N (например -4):"
        ).grid(row=4, column=0, sticky=tk.W, pady=2)

        self.sci_threshold_var = tk.IntVar(value=-4)
        self.sci_spinbox = ttk.Spinbox(
            frame_settings, from_=-15, to=0,
            textvariable=self.sci_threshold_var, width=5
        )
        self.sci_spinbox.grid(
            row=4, column=1, sticky=tk.W, padx=10, pady=2
        )
        self.sci_spinbox.config(state="disabled")

        ttk.Label(
            frame_settings,
            text=(
                "Числа с модулем меньше 10^N выводятся в виде "
                "m·10^k, например 6,344·10^-6."
            ),
            font=("Segoe UI", 8, "italic"),
            foreground="gray",
            wraplength=520,
            justify=tk.LEFT
        ).grid(
            row=5, column=0, columnspan=2,
            sticky=tk.W, pady=(4, 0)
        )

        # ---- Template picker ----
        frame_template = ttk.LabelFrame(
            self.root,
            text="Шаблон стилей Word (.docx)",
            padding=10
        )
        frame_template.pack(fill=tk.X, pady=(0, 15))

        self.lbl_template = ttk.Label(
            frame_template,
            text=(self.template_file or "Шаблон не найден"),
            foreground="black" if self.template_file else "red"
        )
        self.lbl_template.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(
            frame_template, text="Изменить",
            command=self.browse_template
        ).pack(side=tk.RIGHT, padx=5)

        # ---- Convert button ----
        self.btn_convert = ttk.Button(
            self.root,
            text="Конвертировать и сохранить как...",
            command=self.process_and_save,
            state=tk.DISABLED
        )
        self.btn_convert.pack(fill=tk.X, pady=10, ipady=5)

        # ---- Log pane ----
        frame_log = ttk.LabelFrame(
            self.root, text="Статус и Логи", padding=5
        )
        frame_log.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(
            frame_log, height=10, state=tk.DISABLED,
            bg="#f4f4f9", font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Kick off Pandoc check asynchronously so the UI stays responsive.
        self.log("Проверка окружения...")
        self.root.after(
            500,
            lambda: WordConverter.check_pandoc(self.log)
        )

    def toggle_sci(self):
        """Enable/disable the sci-threshold spinbox."""
        self.sci_spinbox.config(
            state="normal" if self.use_sci_var.get() else "disabled"
        )

    def log(self, message):
        """Append a line to the log pane and scroll to the bottom."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()

    # -------------------------------------------------------------------
    # File selection
    # -------------------------------------------------------------------

    def browse_input(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл Mathcad",
            filetypes=[
                ("Mathcad XML", "*.xmcd *.xml"),
                ("All files", "*.*")
            ]
        )
        if filepath:
            self.set_input_file(filepath)

    def on_drop(self, event):
        """Handle a drag-and-drop file."""
        filepath = event.data
        # tkinterdnd2 wraps paths with spaces in {curly braces}.
        if filepath.startswith("{") and filepath.endswith("}"):
            filepath = filepath[1:-1]
        self.set_input_file(filepath)

    def set_input_file(self, filepath):
        if filepath.lower().endswith((".xmcd", ".xml")):
            self.input_file = filepath
            self.lbl_input.config(
                text=os.path.basename(filepath),
                foreground="black"
            )
            self.btn_convert.config(state=tk.NORMAL)
            self.log(f"[INFO] Выбран файл: {filepath}")
        else:
            messagebox.showwarning(
                "Неверный формат",
                "Выберите файл .xmcd или .xml"
            )

    def browse_template(self):
        filepath = filedialog.askopenfilename(
            title="Выберите шаблон стилей Word",
            filetypes=[("Word Documents", "*.docx")]
        )
        if filepath:
            self.template_file = filepath
            self.lbl_template.config(
                text=os.path.basename(filepath),
                foreground="black"
            )
            self.log(f"[INFO] Шаблон изменен на: {filepath}")

    # -------------------------------------------------------------------
    # Conversion driver
    # -------------------------------------------------------------------

    def process_and_save(self):
        if not self.input_file:
            return

        output_file = filedialog.asksaveasfilename(
            title="Сохранить результат как...",
            defaultextension=".docx",
            initialfile="Результат_Mathcad.docx",
            filetypes=[("Word Document", "*.docx")]
        )
        if not output_file:
            return

        # Check that the output file is writable BEFORE doing all the
        # work. Word locks open documents, and we want to warn early.
        if os.path.exists(output_file):
            try:
                with open(output_file, "a"):
                    pass
            except PermissionError:
                messagebox.showerror(
                    "Ошибка доступа",
                    f"Не удалось перезаписать файл:\n"
                    f"{output_file}\n\n"
                    "Скорее всего, он открыт в Word. "
                    "Закройте документ и повторите попытку."
                )
                self.log("[ERROR] Файл заблокирован.")
                return
            except Exception as exc:
                messagebox.showerror(
                    "Ошибка",
                    f"Не удалось получить доступ к файлу:\n{exc}"
                )
                self.log(f"[ERROR] Ошибка доступа: {exc}")
                return

        self.log("\n--- Запуск обработки ---")

        # Push GUI settings into the parser's class-level knobs.
        try:
            MathcadParser.sig_figs_small = (
                self.sig_figs_small_var.get()
            )
            MathcadParser.sig_figs_large = (
                self.sig_figs_large_var.get()
            )
            self.log(
                "[INFO] Округление: "
                f"малые числа = {MathcadParser.sig_figs_small}, "
                f"большие числа = {MathcadParser.sig_figs_large}"
            )
        except Exception:
            self.log(
                "[WARN] Ошибка чтения настроек округления. "
                "Используются значения 4 и 8."
            )
            MathcadParser.sig_figs_small = 4
            MathcadParser.sig_figs_large = 8

        if self.use_sci_var.get():
            try:
                threshold = int(self.sci_threshold_var.get())
            except Exception:
                threshold = -4

            MathcadParser.sci_threshold = threshold
            self.log(
                "[INFO] Научная форма включена: "
                f"числа с модулем < 10^{threshold} "
                "будут выводиться в виде m·10^k"
            )
        else:
            MathcadParser.sci_threshold = None
            self.log("[INFO] Научная форма отключена")

        self.log("[INFO] Парсинг XML файла Mathcad...")

        try:
            markdown_content = MathcadParser.process_file(
                self.input_file
            )
            if not markdown_content.strip():
                self.log(
                    "[WARN] Файл не содержит текста "
                    "или распознанных формул."
                )
        except Exception as exc:
            self.log(f"[ERROR] Ошибка парсинга: {exc}")
            return

        success = WordConverter.convert_to_word(
            markdown_text=markdown_content,
            output_file=output_file,
            template_file=self.template_file,
            log_callback=self.log
        )

        if success:
            open_now = messagebox.askyesno(
                "Успех",
                f"Документ сохранен:\n{output_file}\n\n"
                "Открыть файл сейчас?"
            )
            if open_now:
                self.open_file(output_file)

    def open_file(self, filepath):
        """Open the generated file with the OS-default application."""
        try:
            if sys.platform == "win32":
                os.startfile(filepath)
            elif sys.platform == "darwin":
                os.system(f"open '{filepath}'")
            else:
                os.system(f"xdg-open '{filepath}'")
        except Exception as exc:
            self.log(f"[ERROR] Не удалось открыть файл: {exc}")


# ===========================================================================
# Entry point
# ===========================================================================
if __name__ == "__main__":
    if DND_SUPPORTED:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    style = ttk.Style(root)
    if sys.platform == "win32":
        style.theme_use("vista")

    app = ConverterApp(root)
    root.mainloop()