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

Pipeline:

    .xmcd (XML)  --[MathcadParser.process_file]-->  Markdown+LaTeX
    Markdown+LaTeX  --[WordConverter.convert_to_word]-->  .docx

Two numeric-formatting knobs matter downstream:
  * `sig_figs_small` / `sig_figs_large` — significand digits.
  * `sci_threshold` — switch small numbers to scientific notation.

The parser uses a Russian-locale convention: decimal comma ",".
This matches how Mathcad regional output looks and how the resulting
Word document is expected to read.

IMPORTANT — IMAGINARY UNIT
--------------------------
The imaginary unit is FORCED to `j` everywhere via `IMAGINARY_SYMBOL`.
Rationale:

    * In the *input* part of `<eval>`/`<symEval>` Mathcad records
      `<ml:imag symbol="j">1</ml:imag>` if the user chose `j`.
    * But in `<symResult>` the symbolic engine ALWAYS re-emits the
      imaginary unit with `symbol="i"`, regardless of what the user
      typed. So after `explicit ALL` you get `i` even in a worksheet
      that was authored with `j`.

This mismatch used to leak into the output (`j` in inputs, `i` in
symbolic results). To fix it we IGNORE the per-node `symbol`
attribute entirely and always render `j`. If you ever need to
support worksheets genuinely authored with `i`, reintroduce a
class-level toggle and thread it through the parser — but do NOT
trust `symbol` in `symResult`; it lies.

IMPORTANT — TRAILING-DOT SUBSCRIPTS
-----------------------------------
Mathcad sometimes writes a subscript with a trailing dot, e.g.
`<ml:id subscript="1.">X</ml:id>` (the on-screen `X.1.`). The dot
is a Mathcad placeholder that terminates the index but is not part
of the variable name. If we render it verbatim we get `X_{1.}`
which reads as if a digit is missing after the dot. We therefore
replace ONLY THE LAST character of the subscript — and only when it
is a dot — with a single `*`, producing `X_{1*}`.

Crucially, only the FINAL dot is replaced. Earlier dots in the
subscript are left intact. For example:

    subscript="1."       →  "1*"
    subscript="2'.1.."   →  "2'.1.*"   (the second-to-last dot stays)
    subscript="2'.1."    →  "2'.1*"

An earlier version of this file used `rstrip(".")` which stripped
EVERY trailing dot — turning `2'.1..` into `2'.1*` and silently
eating one of the user's dots. That is wrong: Mathcad only ever
treats the very last dot as a terminator.

IMPORTANT — ON-AXIS COMPLEX NUMBERS
-----------------------------------
Mathcad normally displays complex numbers in polar form
`|z|∠φ°`. But when the point lies exactly on one of the axes
(angle 0°, ±90°, 180° — i.e. one of the two parts is zero), the
polar form is silly (`5∠0°`) or confusing (`5∠180°` for `-5`).
For those cases we emit the rectangular form (`5`, `-5`, `5j`,
`-5j`). Only genuinely off-axis numbers get `r∠φ°`.

IMPORTANT — EXPLICIT ALL AND ON-AXIS ANGLES
-------------------------------------------
After `explicit ALL` Mathcad may leave complex exponentials such as
`e^(j·deg·180)` or `e^(j·deg·-90)` in the symbolic result instead of
collapsing them to `-1` or `-j`. Rendering those as `1∠180°` /
`1∠-90°` is technically correct but reads badly — on-axis angles are
always real or purely imaginary, so the rectangular form is clearer.

We therefore detect the `explicit ALL` command inside `<symEval>`
and, while parsing the corresponding `<symResult>`, render any
on-axis angle (0°, ±90°, 180°) as a plain number (real or imaginary
part only). Angles at other values keep the polar `r∠φ°` form.
Angles outside `<symResult>` — i.e. in the user's original input
expression — are NOT affected by this rule.

IMPORTANT — TRAILING-ZERO TRIMMING IN ANGLES
--------------------------------------------
Angles in polar-form complex numbers (`r∠φ°`) get one extra
cosmetic pass that regular numbers do NOT get: after rounding to
`ANGLE_DECIMALS` places, if EVERY fractional digit turned out to be
a zero, the decimal separator and those zeros are dropped.

    "147,0"   →  "147"
    "-90,0"   →  "-90"
    "0,0"     →  "0"
    "81,6"    →  "81,6"     (non-zero fractional digit: untouched)

WHY THIS IS ANGLES-ONLY
-----------------------
This rule is applied inside `format_angle_value`, which is called
from exactly one place — `parse_complex`, when it assembles the
polar-form output of a complex number. Regular real results are
formatted by `format_num`, which does NOT trim trailing zeros;
that preserves the worksheet's own precision conventions for
plain numeric results.

WHY POST-ROUNDING, NOT PRE-ROUNDING
-----------------------------------
The trim must run on the *already-rounded* string. Example: an
angle of `146.9999999999` rounds to `147,0` at one decimal place
and should then collapse to `147`. A pre-rounding check would have
seen a non-zero fractional digit and never trimmed.

WHY IT MATTERS
--------------
Mathcad itself shows on-axis angles (like 180° or -90°) as plain
integers. Without this trim a polar-form result would show up as
`1∠180,0°`, which reads as noise. This rule is the polar-form
counterpart of the "on-axis → rectangular" rule: they both keep
axis-aligned results legible.

IMPORTANT — "1∠" BEFORE AN ANGLE
--------------------------------
Mathcad writes `1·e^(j·deg·φ)` with an explicit modulus of 1, but it
also writes `Sb·e^(j·deg·φ)` (without any modulus) when the
coefficient is a variable. The exponential itself always has modulus
1, so the correct display of such a product is `Sb·1∠φ°` — the "1"
must be made explicit, otherwise `Sb∠φ°` reads as if `Sb` were the
modulus.

Rule enforced in `_process_mult_angle_args`:
  * pure number before the angle  →  the number IS the modulus,
    no extra "1" (e.g. `2∠120°`);
  * anything else (letter variable, subscripted identifier,
    parenthesized expression, empty context) → prepend "1"
    (e.g. `Sb·1∠120°`).

The check is on the *textual* predecessor of the angle mark, since
that is what the reader will actually see.

IMPORTANT — SIGN CLEANUP AROUND OPERATORS
-----------------------------------------
Three rendering artifacts are avoided at assembly time:

    1. `a + -b`          →  `a - b`
    2. `a - -b`          →  `a + b`
    3. `a \cdot -b`      →  `a \cdot \left(-b\right)`

The first two are handled by `_strip_leading_sign` (used inside the
`plus` / `minus` / `neg` handlers). The third is handled by
`_needs_parens_after_cdot` (used inside `_process_mult_angle_args`
when it joins the assembled operands with `\cdot`).

WHY THESE ARISE
---------------
`explicit ALL` output frequently contains on-axis exponentials that
have been collapsed to `-1` or `-j`. Feeding those into a generic
`+` / `-` / `\cdot` assembly produces strings like

    ... + -100
    502,04 \cdot -j

which are correct but ugly (and, in the `\cdot` case, technically
misleading: LaTeX renders the leading `-` of the right factor as a
binary minus, so `502,04 \cdot -j` reads as `502,04 \cdot (-j)` only
if the reader squints).

The fixes are purely cosmetic — the underlying expression is
unchanged — but they dramatically improve readability, especially
for the complex-arithmetic sections of the worksheet.

Note: we do NOT try to fold `+ (-1)·x` into `- x`, nor do we sort
factors to put numbers first. Both would require reaching inside
`\left(...\right)` blocks and offer only marginal cosmetic gain.
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
        attribute (float, in points). We sort by `top` to reconstruct
        the visual reading order. Column (`left`) is ignored because
        Mathcad users usually write top-to-bottom.
    """

    # Numeric formatting knobs (overridable from the GUI).
    sig_figs_small = 4
    sig_figs_large = 8

    # Number of decimal places for angles in polar form.
    #
    # NOTE: This applies ONLY to the angle in `r∠φ°` polar output
    # (see `format_angle_value`). Regular real numbers go through
    # `format_num` / `format_scientific` and are unaffected.
    #
    # NOTE: After rounding to this many decimals, the trailing
    # fractional zeros are dropped if ALL of them are zero — see
    # the module-level note "IMPORTANT — TRAILING-ZERO TRIMMING IN
    # ANGLES".
    ANGLE_DECIMALS = 1

    # Scientific-notation threshold: numbers with |x| < 10^sci_threshold
    # are rendered as "m·10^k". `None` disables the feature.
    sci_threshold = None

    # -------------------------------------------------------------------
    # Imaginary-unit symbol
    # -------------------------------------------------------------------
    #
    # The single source of truth for how the imaginary unit is rendered.
    # We FORCE "j" here on purpose; see the module-level note
    # "IMPORTANT — IMAGINARY UNIT" for why the per-node `symbol`
    # attribute cannot be trusted.
    #
    # If a future maintainer needs to support worksheets authored with
    # `i`, change this constant — but be aware this will NOT fix the
    # `symResult` case by itself, because Mathcad always emits `i`
    # there. You would have to scan the input side for the first
    # user-chosen symbol and remember it across the whole parse.
    IMAGINARY_SYMBOL = "j"

    # Tolerance for "this value is zero" checks (real/imag parts).
    # Chosen to be much larger than the smallest denormal double but
    # far below any physically meaningful number in a worksheet.
    AXIS_EPS = 1e-12

    # -------------------------------------------------------------------
    # explicit-ALL flag
    # -------------------------------------------------------------------
    #
    # True only while we are parsing a `<symResult>` whose sibling
    # `<command>` contains `explicit ALL`. In that mode, on-axis
    # angles (`e^(j·deg·0°)`, `e^(j·deg·±90°)`, `e^(j·deg·180°)`) are
    # rendered as plain rectangular numbers instead of polar form.
    #
    # The flag is set/reset in `parse_eval_to_latex` around the call
    # that parses the `<symResult>` child. It is a class attribute
    # (not an instance attribute) because `MathcadParser` is used as
    # a namespace of classmethods — see `process_file`.
    _explicit_all_mode = False

    # -------------------------------------------------------------------
    # Angle-mark machinery (deferred ∠ substitution)
    # -------------------------------------------------------------------
    #
    # WHY THIS EXISTS
    # ---------------
    # When we encounter an exponent `e^(j·deg·φ)`, we want to emit
    # `∠φ°`. But the parent `mult` handler still needs to see the
    # result as an opaque *operand* — it inserts an implicit "1"
    # modulus depending on what precedes the angle. If we emitted raw
    # LaTeX with `\angle` immediately, the mult handler would try to
    # reformat it and the "is this a coefficient?" logic would misfire.
    #
    # The placeholder approach:
    #   1. `push_angle_mark()` returns a token "@@ANGLEn@@" with no
    #      LaTeX special characters. It survives any string concat
    #      and comparison safely.
    #   2. At the outermost parse frame (parse_node_to_latex_root /
    #      parse_eval_to_latex_root) we call `pop_all_angle_marks()`
    #      and substitute every token for its real `\angle...` LaTeX,
    #      inserting an implicit "1" when no coefficient precedes.
    #
    # IMPORTANT INVARIANT
    # -------------------
    # Only the `mult` branch may pop marks, and it must pop ONLY the
    # marks pushed during ITS OWN argument parsing. Marks pushed by
    # sibling branches (e.g. the numerator of a fraction) must stay on
    # the stack until the outer frame resolves them.
    #
    # Bug history: an earlier version called `pop_all_angle_marks()`
    # inside `mult`, which stole sibling marks — e.g. `pow(e, i·deg·84)`
    # (numerator) had its mark swallowed by the `mult` inside
    # `pow(e, i·deg·78)` (denominator of a sibling fraction). Result:
    # one angle rendered as `1∠84°`, the other as raw `@@ANGLE3@@`.
    # `pop_angle_marks_since(depth)` fixes this.

    # Regex matching "@@ANGLE0@@", "@@ANGLE1@@", ...
    ANGLE_MARK_RE = re.compile(r"@@ANGLE\d+@@")

    # Stack of active marks. Each entry is a single-key dict.
    _angle_mark_stack = []

    # Monotonic counter so nested marks never collide.
    _angle_mark_counter = 0

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

    @classmethod
    def _angle_stack_depth(cls):
        """
        Current number of stacked angle marks.

        NOTE: We measure this BEFORE parsing any arguments so that a
        `mult` handler can distinguish "marks pushed by my operands"
        (depth >= this value) from "marks pushed by sibling branches"
        (depth <  this value, must not be touched).
        """
        return len(cls._angle_mark_stack)

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

        NOTE: This function does NOT trim trailing fractional zeros.
        Its output is used for regular numbers, matrix entries, etc.
        Angle trimming lives in `format_angle_value` (see the
        module-level note "IMPORTANT — TRAILING-ZERO TRIMMING IN
        ANGLES").
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
        r"""
        Render a polar-form angle with a fixed number of decimals.

        NOTE (rounding): Angles in Mathcad results often carry 15+
        digits of floating-point noise (e.g. 81.6078831602174).
        Rounding to `ANGLE_DECIMALS` places keeps the document
        readable. "Negative zero" is coerced to exactly zero.

        NOTE (trailing-zero trimming — see module-level note
        "IMPORTANT — TRAILING-ZERO TRIMMING IN ANGLES"):
        After rounding, if EVERY fractional digit is a zero, the
        decimal separator and those zeros are dropped:

            147.0  →  "147,0"  →  "147"
            -90.0  →  "-90,0"  →  "-90"
             0.0   →  "0,0"    →  "0"
            81.6   →  "81,6"   (non-zero fraction: untouched)

        The special string "-0" (produced when a tiny negative value
        rounds to "-0,0" with fewer than the fractional digits) is
        normalised to "0" so we never emit the LaTeX-hostile "-0".

        IMPORTANT: this method is called ONLY from `parse_complex`
        when rendering the polar form `r∠φ°`. Regular real numbers
        go through `format_num` and do NOT get this trim.
        """
        try:
            value = float(angle)
        except (ValueError, TypeError):
            return str(angle).replace(".", ",")

        if math.isclose(value, 0.0, abs_tol=1e-12):
            value = 0.0

        text = f"{value:.{cls.ANGLE_DECIMALS}f}"

        # NOTE (trailing-zero trimming): the trim MUST run on the
        # rounded text, not on the pre-rounding value — otherwise an
        # angle like 146.9999999 would never be trimmed even though
        # it rounds to 147,0 and should collapse to 147.
        if "." in text:
            integer_part, fractional_part = text.split(".", 1)

            if fractional_part and all(ch == "0" for ch in fractional_part):
                text = integer_part

                # "-0" (from e.g. -0.04 rounded to "-0.0") is a
                # rendering artifact; emit plain "0" instead.
                if text == "-0":
                    text = "0"

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

        NOTE: The asterisk `*` is deliberately NOT escaped — the
        trailing-dot-subscript quirk relies on emitting a literal `*`
        for subscripts like `1.` → `1*`.
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

        NOTE (trailing-dot subscripts):
            Mathcad sometimes stores a subscript with a trailing dot,
            e.g. <ml:id subscript="1.">X</ml:id>, which corresponds to
            the on-screen Mathcad notation `X.1.` (a variable whose
            "index" ends with a dot before the definition symbol). The
            trailing dot is a Mathcad placeholder, not part of the
            variable name. If we render it verbatim the subscript reads
            as `X_{1.}` — as if a digit is missing after the dot.

            We replace ONLY THE LAST character of the subscript — and
            only when it is a dot — with a single `*`. Any dots that
            appear earlier inside the subscript are preserved:

                subscript="1."       →  "1*"
                subscript="2'.1.."   →  "2'.1.*"   (earlier dot kept)
                subscript="2'.1."    →  "2'.1*"

            Earlier versions of this file used `rstrip(".")`, which
            stripped EVERY trailing dot and therefore ate legitimate
            dots in indices like `2'.1..`. Do not reintroduce that —
            Mathcad only ever treats the very last dot as a terminator.

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
            # NOTE (trailing-dot subscripts): replace only the LAST
            # character, and only if it is a dot. Do NOT use rstrip —
            # that would eat legitimate earlier dots.
            if subscript.endswith("."):
                subscript = subscript[:-1] + "*"

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

    @classmethod
    def get_imag_symbol(cls, node=None):
        """
        Return the symbol used for the imaginary unit.

        CRITICAL — DO NOT READ `symbol` FROM THE NODE HERE.
        --------------------------------------------------
        See the module-level note "IMPORTANT — IMAGINARY UNIT".
        Mathcad's symbolic engine ALWAYS re-emits the imaginary unit
        with symbol="i" inside `<symResult>`, even if the user wrote
        `j`. The input side may contain symbol="j". If we honoured
        the attribute we would produce mixed `i`/`j` output in the
        same document (this was the original bug: `a := 1e^(j·deg·120)`
        rendered fine, but after `explicit ALL` the result turned
        into `i`).

        We therefore IGNORE the node and always return
        `cls.IMAGINARY_SYMBOL`. The `node` argument is kept for
        signature compatibility with existing call sites.
        """
        return cls.IMAGINARY_SYMBOL

    @classmethod
    def parse_complex(cls, real_value, imag_value, imag_symbol=None):
        r"""
        Render a rectangular complex number.

        MATHCAD QUIRK (polar vs. rectangular display):
            Mathcad normally displays complex numbers in polar form
            `|z|\angle\phi^\circ`. However, when the number lies
            EXACTLY on one of the axes — i.e. when one of its parts
            (real or imaginary) is numerically zero — the polar form
            is misleading:

                -5 + 0j   →   5\angle 180^\circ   (reads worse than -5)
                +5 + 0j   →   5\angle 0^\circ     (silly)
                 0 + 5j   →   5\angle 90^\circ    (unusual notation)
                 0 - 5j   →   5\angle -90^\circ

            For those cases we fall back to the rectangular form:
            `-5`, `5`, `5j`, `-5j`. This mirrors what Mathcad itself
            shows on screen for on-axis results.

            The off-axis case (both parts non-zero) keeps the polar
            form `r\angle\phi^\circ`.

        NOTE (angle trimming): the polar angle is produced by
        `format_angle_value`, which trims trailing fractional zeros
        after rounding. So an off-axis result whose true angle rounds
        to a whole number renders as `r\angle 147^\circ`, NOT
        `r\angle 147,0^\circ`. This is the ONLY call site of
        `format_angle_value` — see the module-level note
        "IMPORTANT — TRAILING-ZERO TRIMMING IN ANGLES" for the
        rationale and for why the rule is angles-only.

        NOTE (imag_symbol):
            The `imag_symbol` argument is kept for call-site
            compatibility, but the FINAL displayed symbol is always
            `cls.IMAGINARY_SYMBOL`. See the module-level note
            "IMPORTANT — IMAGINARY UNIT" for why the per-node symbol
            cannot be trusted.
        """
        symbol = cls.IMAGINARY_SYMBOL

        # --- Try to interpret both parts as numbers -------------------
        try:
            real_float = float(real_value)
            imag_float = float(imag_value)
            numeric = True
        except (ValueError, TypeError):
            numeric = False

        # --- Non-numeric (symbolic) parts: plain concatenation --------
        if not numeric:
            real_text = cls.format_num(str(real_value))
            imag_text = cls.format_num(str(imag_value))
            sign = "" if imag_text.startswith("-") else "+"
            return f"{real_text}{sign}{imag_text}{symbol}"

        # --- On-axis check --------------------------------------------
        # If either part is numerically zero the point lies on one of
        # the axes (angle 0°, ±90°, 180°). Use rectangular form.
        on_axis = (
            math.isclose(real_float, 0.0, abs_tol=cls.AXIS_EPS)
            or math.isclose(imag_float, 0.0, abs_tol=cls.AXIS_EPS)
        )

        if on_axis:
            return cls._format_rectangular(
                real_float, imag_float, symbol
            )

        # --- Off-axis: polar form -------------------------------------
        modulus, angle = cls.calculate_complex_polar(
            real_float, imag_float
        )

        modulus_text = cls.format_num(str(modulus))
        angle_text = cls.format_angle_value(angle)

        return (
            rf"{modulus_text}"
            rf"\angle"
            rf"{angle_text}^\circ"
        )

    @classmethod
    def _format_rectangular(cls, real_value, imag_value, symbol):
        r"""
        Render a complex number in rectangular form, omitting parts
        that are numerically zero.

        Examples:
            ( 5,  0) → "5"      (pure real)
            (-5,  0) → "-5"
            ( 0,  5) → "5j"     (pure imaginary)
            ( 0, -5) → "-5j"
            ( 0,  1) → "j"
            ( 0, -1) → "-j"
            ( 2,  3) → "2 + 3j" (safety net — normally not reached
                                 via the on-axis branch)
        """
        r_zero = math.isclose(
            real_value, 0.0, abs_tol=cls.AXIS_EPS
        )
        m_zero = math.isclose(
            imag_value, 0.0, abs_tol=cls.AXIS_EPS
        )

        if r_zero and m_zero:
            return "0"

        if m_zero:
            # Pure real number.
            return cls.format_num(str(real_value))

        if r_zero:
            # Pure imaginary number.
            sign = "-" if imag_value < 0 else ""
            abs_imag = abs(imag_value)

            if math.isclose(abs_imag, 1.0, abs_tol=cls.AXIS_EPS):
                return f"{sign}{symbol}"

            imag_text = cls.format_num(str(abs_imag))
            return f"{sign}{imag_text}{symbol}"

        # Both parts non-zero. This branch is not reached from the
        # axis path in `parse_complex`, but we keep it for safety.
        real_text = cls.format_num(str(real_value))
        sign = "-" if imag_value < 0 else "+"
        abs_imag = abs(imag_value)

        if math.isclose(abs_imag, 1.0, abs_tol=cls.AXIS_EPS):
            return f"{real_text} {sign} {symbol}"

        imag_text = cls.format_num(str(abs_imag))
        return f"{real_text} {sign} {imag_text}{symbol}"

    # -------------------------------------------------------------------
    # Sign helpers — used by `plus` / `minus` / `neg` and by the
    # multiplication assembler to avoid dangling binary operators.
    #
    # See the module-level note "IMPORTANT — SIGN CLEANUP AROUND
    # OPERATORS" for the rationale and the exact transformations.
    # -------------------------------------------------------------------

    @staticmethod
    def _strip_leading_sign(text):
        """
        Split a LaTeX fragment into a leading sign and the remainder.

        Returns (sign, body) where:
          * sign is "+" or "-";
          * body is the fragment without its leading sign (and
            without leading whitespace).

        If the fragment does not start with a bare "+"/"-" at top
        level, returns ("+", text.strip()).

        EXAMPLES
            "100"               →  ("+", "100")
            "-100"              →  ("-", "100")
            "+j"                →  ("+", "j")
            "-j"                →  ("-", "j")
            "- 502,04"          →  ("-", "502,04")
            "\\left(-1\\right)" →  ("+", "\\left(-1\\right)")
            "a + b"             →  ("+", "a + b")

        IMPORTANT
            We deliberately do NOT peek inside `\\left(...\\right)`
            or other composite fragments. The leading `-` of a
            parenthesized negative number is not at top level, so it
            is not stripped — which is exactly what we want, because
            the parentheses already disambiguate.
        """
        stripped = text.strip()
        if not stripped:
            return "+", ""

        first = stripped[0]

        if first == "-":
            body = stripped[1:].lstrip()
            if body:
                return "-", body
        elif first == "+":
            body = stripped[1:].lstrip()
            if body:
                return "+", body

        return "+", stripped

    @classmethod
    def _apply_negation(cls, text):
        """
        Return the negation of a LaTeX fragment, folding a leading sign
        when possible.

        EXAMPLES
            "100"                →  "-100"
            "-100"               →  "100"
            "j"                  →  "-j"
            "-j"                 →  "j"
            "\\left(-1\\right)"  →  "-\\left(-1\\right)"

        NOTE: We fold only a top-level leading sign. A `\\left(...)`
        block keeps its leading "-" untouched (unless it is the whole
        fragment, in which case we still just prepend another "-").
        That keeps `-(-1)` visible as such rather than silently
        simplifying to `1`, which could confuse a reader who is
        checking the symbolic derivation.
        """
        sign, body = cls._strip_leading_sign(text)

        if not body:
            return "-"

        if sign == "-":
            return body

        return f"-{body}"

    @staticmethod
    def _needs_parens_after_cdot(text):
        """
        Decide whether a factor must be wrapped in `\\left(...\\right)`
        when it appears after a `\\cdot`.

        WHY THIS EXISTS
            `a \\cdot -b` renders as `a · − b` — LaTeX treats the
            leading "-" of the right operand as a *binary* minus, so
            the product looks like a subtraction with a stray factor.
            `a \\cdot \\left(-b\\right)` is unambiguous.

        We only wrap when:
          * the fragment starts with a bare "-" (top level), AND
          * it is not already parenthesized with `\\left(`.

        Examples:
            "502,04"               →  False
            "-j"                   →  True
            "-502,04j"             →  True
            "-\\frac{1}{2}"        →  True
            "\\left(-1\\right)"    →  False  (already protected)
            "1\\angle84^\\circ"    →  False  (angle fragments never
                                             need parens)
        """
        stripped = text.strip()
        if not stripped:
            return False

        if stripped.startswith(r"\left("):
            return False

        return stripped.startswith("-")

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

    # -------------------------------------------------------------------
    # Extraction of the polar angle from `e^(j·deg·φ)`
    # -------------------------------------------------------------------

    @classmethod
    def _is_i_times_deg_operand(cls, operand):
        """
        Helper: does this operand look like `1j·deg` (imag × deg)?
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
        has_deg = any(cls.is_degree_unit(p) for p in factors)

        return has_imag and has_deg

    @classmethod
    def extract_polar_angle_degrees(cls, node):
        r"""
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

        i_deg_found = False
        angle_parts = []

        for operand in operands:
            if cls._is_i_times_deg_operand(operand):
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
    def extract_polar_angle_numeric(cls, node):
        """
        Same shape as `extract_polar_angle_degrees`, but returns the
        numeric value of the angle as a float, or None if the angle is
        not a simple numeric constant.

        Used by the `pow` handler when `_explicit_all_mode` is set and
        we want to decide whether the angle is on-axis. Non-numeric
        angles (e.g. `e^(j·deg·x)`) return None and fall back to the
        normal polar rendering.
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

        i_deg_found = False
        numeric_value = 1.0
        numeric_count = 0

        for operand in operands:
            if cls._is_i_times_deg_operand(operand):
                if i_deg_found:
                    return None
                i_deg_found = True
                continue

            tag = cls.strip_ns(operand.tag)

            if tag != "real":
                # Non-numeric factor in the exponent — bail out.
                return None

            try:
                numeric_value *= float(
                    (operand.text or "0").strip()
                )
                numeric_count += 1
            except (ValueError, TypeError):
                return None

        if not i_deg_found or numeric_count == 0:
            return None

        return numeric_value

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
    def angle_to_rectangular(cls, angle_deg):
        """
        Convert an on-axis angle (in degrees) to its rectangular
        representation as a short LaTeX fragment.

        Returns None for off-axis angles, so the caller can fall back
        to the polar form. Returns a string for on-axis angles:

            0°      → "1"
            ±180°   → "-1"
            90°     → symbol (j by default)
            -90°    → "-" + symbol

        Used only in explicit-ALL mode (see module-level note).
        """
        normalized = cls.normalize_angle_degrees(angle_deg)
        symbol = cls.IMAGINARY_SYMBOL

        if math.isclose(normalized, 0.0, abs_tol=1e-9):
            return "1"

        if math.isclose(abs(normalized), 180.0, abs_tol=1e-9):
            return "-1"

        if math.isclose(normalized, 90.0, abs_tol=1e-9):
            return symbol

        if math.isclose(normalized, -90.0, abs_tol=1e-9):
            return f"-{symbol}"

        return None

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
    # Angle-mark stack management
    # -------------------------------------------------------------------

    @classmethod
    def push_angle_mark(cls, latex):
        r"""
        Hide a ready-made LaTeX angle (e.g. \angle 84^\circ) inside a
        placeholder token "@@ANGLEn@@" and return that token.

        WHY THIS IS NEEDED — see the module-level note above.
        """
        mark = f"@@ANGLE{cls._angle_mark_counter}@@"
        cls._angle_mark_counter += 1
        cls._angle_mark_stack.append({mark: latex})
        return mark

    @classmethod
    def pop_angle_marks_since(cls, depth):
        """
        Pop and return every mark pushed after the given stack depth.

        Marks BELOW `depth` belong to outer frames and must stay on
        the stack so those frames can resolve them later.

        Returns a flat dict {mark: latex}.
        """
        popped = cls._angle_mark_stack[depth:]
        del cls._angle_mark_stack[depth:]

        merged = {}
        for marks in popped:
            merged.update(marks)

        return merged

    @classmethod
    def pop_all_angle_marks(cls):
        """
        Merge and clear every active angle placeholder.

        Used ONLY by the outermost entry points
        (parse_node_to_latex_root / parse_eval_to_latex_root).

        Returns a flat dict {mark: latex}.
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

        NOTE: This is the *last-resort* check, applied at the outermost
        frames (root handlers). In the common `mult` path the "1" is
        already inserted by `_process_mult_angle_args` before the
        assembled string reaches `restore_angle_symbols`, so the
        preceding character is a digit and no double prefix occurs.

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

        NOTE: This is the final place where the implicit-unity rule
        is enforced. In the common `mult` path, the "1" is already
        inserted by `_process_mult_angle_args` (so `before_text`
        ends with a digit and no extra "1" is added here). The check
        is still needed for angles that were not assembled through
        the mult handler, e.g. a bare `e^(j·deg·120°)` at the top
        level of a define.
        """
        for mark, angle_latex in marks.items():
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
        True if `text` is a bare angle mark (unresolved) or a plain
        resolved angle like "1\\angle 84^\\circ" without extra factors.

        NOTE: This is called on already-parsed fragments, so the text
        may contain:
          * An unresolved placeholder "@@ANGLEn@@".
          * A resolved angle without a wrapping parens.
        """
        stripped = text.strip()

        if MathcadParser.ANGLE_MARK_RE.fullmatch(stripped):
            return True

        # Resolved angle, no parens, no trailing operators.
        # Pattern: optional "1", \angle, content without ^\circ going
        # past it, ends with ^\circ or ^{\circ}.
        if re.fullmatch(
            r"1?\\angle.*?\^\s*(?:\\circ|\{\\circ\})",
            stripped
        ):
            return True

        return False

    @staticmethod
    def _is_angle_mark_wrapped(text):
        """
        True if `text` is an angle fragment (resolved or unresolved),
        optionally wrapped in \left( ... \right).

        NOTE: The regex is deliberately tight. It must NOT match
        strings like "1\\angle84^\\circ · x" (angle followed by extra
        factors), because those are not pure angles.
        """
        stripped = text.strip()

        if MathcadParser.ANGLE_MARK_RE.fullmatch(stripped):
            return True

        if re.fullmatch(
            r"(?:\\left\()?\s*1?\s*\\angle.*?\^\s*"
            r"(?:\\circ|\{\\circ\})\s*(?:\\right\))?",
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
    def _try_parse_number(text):
        """
        Try to parse a plain decimal number written with a comma decimal
        separator (as `format_num` emits). Returns a float, or None.
        """
        stripped = text.strip()
        if not re.fullmatch(r"[+-]?\d+(?:,\d+)?", stripped):
            return None
        try:
            return float(stripped.replace(",", "."))
        except ValueError:
            return None

    @staticmethod
    def _format_number(value):
        """
        Inverse of `_try_parse_number`. Returns a string with comma
        decimal separator, or None if the value cannot be represented
        in plain (non-scientific) form.
        """
        if value == int(value) and abs(value) < 1e15:
            return str(int(value))

        text = f"{value:.10g}"
        if "e" in text.lower() or "E" in text:
            return None

        return text.replace(".", ",")

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
          2. Decides where an implicit `1` modulus is needed. The rule
             (see module-level note "IMPORTANT — 1∠ BEFORE AN ANGLE"):
               - pure number before the angle → no extra "1";
               - anything else → prepend "1".
          3. In explicit-ALL mode, additionally merges adjacent pure
             numeric factors so that, e.g., `100·(-1)` collapses to
             `-100`.
          4. Joins everything with `\cdot` — except in the special case
             where a coefficient directly precedes an angle, where the
             implicit product is written juxtaposed.
          5. Wraps a minus-leading right operand in `\left(...\right)`
             so the leading "-" is not read as a binary minus. See the
             module-level note "IMPORTANT — SIGN CLEANUP AROUND
             OPERATORS".
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
        # MATHCAD QUIRK: `X·e^(j·deg·φ)` is really `X·(1∠φ°)`.
        # When `X` is a pure number it doubles as the modulus of the
        # polar form and we can write `X∠φ°` directly. When `X` is
        # anything else (a letter variable, a subscripted identifier,
        # a parenthesized expression, or nothing at all) the modulus
        # of the polar part is still 1 and must be shown explicitly:
        # `X·1∠φ°`.
        #
        # The rule therefore is binary: pure number → no prefix,
        # otherwise → prefix "1".
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
                # Nothing before the angle in this mult chain.
                f["prefix_one"] = True
                continue

            prev_text = factors[prev_idx]["text"].strip()

            if cls._is_pure_number_latex(prev_text):
                # The number itself is the modulus — no "1".
                continue

            # Anything else needs the explicit "1" modulus.
            f["prefix_one"] = True

        # Step 2.5 (explicit-ALL only): merge adjacent pure-numeric
        # factors. In explicit-ALL mode on-axis angles are rendered as
        # plain numbers (see the module-level note), so a product like
        # `100·e^(j·180°)` becomes [100, "-1"]. Joining those with a
        # `\cdot` would give the awkward `100 \cdot -1`; multiplying
        # them gives the natural `-100`.
        if cls._explicit_all_mode:
            i = 0
            while i < len(factors) - 1:
                if not factors[i]["keep"]:
                    i += 1
                    continue

                a = cls._try_parse_number(factors[i]["text"])
                if a is None:
                    i += 1
                    continue

                # Find the next kept factor.
                j = i + 1
                while j < len(factors) and not factors[j]["keep"]:
                    j += 1
                if j >= len(factors):
                    break

                b = cls._try_parse_number(factors[j]["text"])
                if b is None:
                    i += 1
                    continue

                product_text = cls._format_number(a * b)
                if product_text is None:
                    i += 1
                    continue

                factors[i]["text"] = product_text
                factors[i]["is_pure_number"] = True
                factors[j]["keep"] = False
                i += 1

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
                # First factor: even a leading "-" is fine at the
                # start of an expression.
                result = text
                continue

            # If the current piece is a bare angle mark and the previous
            # piece is a simple coefficient, juxtapose them (a·∠x →
            # a∠x). Otherwise insert an explicit \cdot.
            if (cls._is_bare_angle_mark(text)
                    and cls._is_simple_coefficient(result.strip())):
                result += text
                continue

            # Wrap a minus-leading right operand: `a \cdot -b` would
            # render with a dangling binary minus. `a \cdot (-b)` is
            # unambiguous. See module-level note "IMPORTANT — SIGN
            # CLEANUP AROUND OPERATORS".
            if cls._needs_parens_after_cdot(text):
                text = rf"\left({text.strip()}\right)"

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
        keep placeholders unresolved so that the parent `mult` / `plus`
        handlers can see them as opaque tokens.
        """
        try:
            latex = cls.parse_node_to_latex(node)
        finally:
            marks = cls.pop_all_angle_marks()

        if marks:
            latex = cls.restore_angle_symbols(latex, marks)

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
        # symbol alone is the correct notation. Same for `-1` → `-j`.
        #
        # CRITICAL — DO NOT READ `symbol` FROM THE NODE HERE.
        # ------------------------------------------------------------
        # The `symbol` attribute is NOT reliable. In the input side
        # Mathcad writes whatever the user chose (i or j), but the
        # symbolic engine ALWAYS emits symbol="i" inside <symResult>.
        # Reading the attribute here reintroduces the "j becomes i
        # after explicit all" bug. We therefore use the class constant
        # `cls.IMAGINARY_SYMBOL` unconditionally.
        if tag == "imag":
            symbol = cls.IMAGINARY_SYMBOL
            raw = node.text.strip() if node.text else ""
            number = cls.format_num(raw)

            if not raw or number == "1":
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
            #
            # NOTE (on-axis): parse_complex itself now downgrades
            # on-axis results (angle 0°, ±90°, 180°) to rectangular
            # form, so `5 + 0j` → `5`, `0 + 5j` → `5j`, etc.
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
                    # NOTE: we deliberately ignore the node's own
                    # symbol — see the imag branch above.
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

                    # NOTE: same as above — no per-node symbol.
                    imag_symbol = cls.get_imag_symbol(second_node)
                    return cls.parse_complex(
                        real_value, imag_value, imag_symbol
                    )

            # ============================================================
            # CRITICAL: capture the angle-mark stack depth BEFORE parsing
            # operands. `mult` will later use this to pop only the marks
            # pushed by its own arguments — NOT marks owned by sibling
            # branches (e.g. a numerator whose angle was pushed earlier).
            # ============================================================
            depth_before = cls._angle_stack_depth()

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
                # Pop ONLY the marks pushed during THIS mult's arg
                # parsing. Outer marks stay on the stack and will be
                # resolved by the outer frame's root handler.
                marks_to_restore = cls.pop_angle_marks_since(
                    depth_before
                )
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
            #
            # SIGN CLEANUP (see module-level note "IMPORTANT — SIGN
            # CLEANUP AROUND OPERATORS"): a negative right operand
            # becomes a subtraction instead of `+ -x`. Works for any
            # number of operands, though Mathcad emits binary plus
            # in practice.
            if op == "plus":
                if not args:
                    return ""

                result = args[0]
                for right in args[1:]:
                    sign, body = cls._strip_leading_sign(right)
                    if sign == "-":
                        result = f"{result} - {body}"
                    else:
                        result = f"{result} + {right}"
                return result

            # ---- subtraction ----
            #
            # SIGN CLEANUP: a negative right operand becomes an
            # addition instead of `- -x`.
            if op == "minus":
                if not args:
                    return "-"

                if len(args) == 1:
                    # Unary minus.
                    return cls._apply_negation(args[0])

                result = args[0]
                for right in args[1:]:
                    sign, body = cls._strip_leading_sign(right)
                    if sign == "-":
                        result = f"{result} + {body}"
                    else:
                        result = f"{result} - {right}"
                return result

            # ---- unary negation ----
            # MATHCAD QUIRK: Mathcad sometimes nests a long chain of
            # <neg/> around a placeholder (visible in empty-input
            # regions). We render the outermost minus and fold any
            # nested leading sign.
            if op == "neg":
                if not args:
                    return "-"
                return cls._apply_negation(args[0])

            # ---- exponentiation ---- 
            # Intercepts e^(j·deg·φ) and produces an angle mark; other
            # powers fall through to plain `{base}^{exp}`.
            #
            # In explicit-ALL mode (see module-level note), on-axis
            # angles (0°, ±90°, 180°) are converted directly to their
            # rectangular form, bypassing the angle-mark machinery.
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
                        if cls._explicit_all_mode:
                            angle_value = (
                                cls.extract_polar_angle_numeric(
                                    exp_node
                                )
                            )
                            if angle_value is not None:
                                rect = cls.angle_to_rectangular(
                                    angle_value
                                )
                                if rect is not None:
                                    return rect
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
    def _sym_eval_uses_explicit_all(cls, sym_eval_node):
        """
        Return True if the given <symEval> contains a `<command>` whose
        text mentions both `explicit` and `ALL`.

        Used to enable the on-axis-rectangular rule (see module-level
        note "IMPORTANT — EXPLICIT ALL AND ON-AXIS ANGLES").
        """
        if sym_eval_node is None:
            return False

        command_node = None
        for child in sym_eval_node:
            if cls.strip_ns(child.tag) == "command":
                command_node = child
                break

        if command_node is None:
            return False

        text = " ".join(command_node.itertext())
        tokens = text.split()

        return "explicit" in tokens and "ALL" in tokens

    @classmethod
    def parse_eval_to_latex_root(cls, node):
        """
        Entry point for `<eval>` regions (numeric or symbolic evaluation).

        NOTE: Same placeholder-flushing logic as
        `parse_node_to_latex_root` — we need to resolve angle marks
        before returning to the caller. This is the ONLY place where
        marks left over from a whole `<eval>` subtree get resolved.
        """
        try:
            latex = cls.parse_eval_to_latex(node)
        finally:
            marks = cls.pop_all_angle_marks()

        if marks:
            latex = cls.restore_angle_symbols(latex, marks)

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

        IMPORTANT — EXPLICIT ALL
        ------------------------
        When the symbolic part of the evaluation was driven by
        `explicit ALL`, we flip `_explicit_all_mode` for the duration
        of the `<symResult>` parse. In that mode the `pow` handler
        emits rectangular numbers for on-axis angles instead of
        polar forms. The flag is saved and restored so nested calls
        do not leak (see module-level note "IMPORTANT — EXPLICIT
        ALL AND ON-AXIS ANGLES").
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
                # Detect `explicit ALL` and toggle the mode only for
                # the symResult parse.
                uses_explicit_all = cls._sym_eval_uses_explicit_all(
                    sym_eval_node
                )
                saved_mode = cls._explicit_all_mode
                cls._explicit_all_mode = uses_explicit_all
                try:
                    sym_result_latex = cls.parse_node_to_latex(
                        sym_result_node
                    )
                finally:
                    cls._explicit_all_mode = saved_mode

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

        # Dev fallback: an absolute path used during testing.
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