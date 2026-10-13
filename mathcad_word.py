r"""
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

Numeric-formatting settings (editable in the GUI):
  * `rounding_mode`: "decimal_places" or the legacy "significant" mode.
  * `decimal_places`: maximum fractional digits; trailing zeros are removed.
  * `sig_figs_small` / `sig_figs_large`: legacy significant-digit controls.
  * `ANGLE_DECIMALS`: independent precision for polar angles.
  * `sci_threshold`: optional scientific notation for small numbers (|x| < 10^N).
  * `sci_threshold_large`: optional scientific notation for large numbers (|x| >= 10^M).

Fixed-decimal formatting uses Decimal with ROUND_HALF_UP: a tie rounds
away from zero (2.675 -> 2.68, -2.675 -> -2.68). XML decimals are read
without a float conversion. Complex calculations still use floats, as
before. Formatting changes display only, not the original worksheet.
All numeric XML nodes are formatted, including numbers in input formulas.
If scientific notation is enabled, decimal_places applies to the mantissa;
small nonzero numbers then remain visible instead of rounding to zero.

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

IMPORTANT — EXPLICIT ALL, SUBSTITUTED FACTORS, AND ON-AXIS ANGLES
-----------------------------------------------------------------
The entire purpose of `explicit ALL` in Mathcad is to show the
step-by-step substitution of numeric values for each variable BEFORE
any arithmetic is evaluated. For example, if `a = 3` and `b = 4`,
then `a · b` after `explicit ALL` is stored in `<symResult>` as
`<apply><mult/><real>3</real><real>4</real></apply>` and MUST be
rendered as `3 \cdot 4` — NEVER pre-multiplied into `12`!

CRITICAL INVARIANT FOR FUTURE MAINTAINERS:
    NEVER merge or multiply two `<real>` (or general numeric) operands
    inside `_process_mult_angle_args`! Doing so destroys the 1-to-1
    correspondence between variables in the formula and substituted
    numbers in the `explicit ALL` step (and can even round small
    products like `0,014 · 0,012` to `0`).

The ONLY case where a factor in `<symResult>` is folded into an
adjacent number is when that factor is an ON-AXIS EULER EXPONENTIAL
node `<apply><pow/><id>e</id><apply><mult/>...deg...φ</apply></apply>`
with `φ` at 0° or ±180°. After `explicit ALL`, Mathcad may leave
complex exponentials such as `e^(j·deg·180)` or `e^(j·deg·-90)` in
`<symResult>`. We render on-axis angles (0°, ±90°, 180°) in
rectangular form (`1`, `-1`, `j`, `-j`). When `e^(j·deg·180)` (`-1`)
or `e^(j·deg·0)` (`1`) is the angle part of a polar product like
`100 · e^(j·deg·180)`, we fold thatEuler factor's sign into the
preceding numeric modulus (`-100`). We verify that the factor is
genuinely an on-axis Euler `<pow>` node via `_is_on_axis_euler_node`
before folding.

IMPORTANT — ROUNDING AND TRAILING ZEROS
--------------------------------------
Numbers in fixed-decimal mode and polar angles use one formatter:
round first, then remove only trailing FRACTIONAL zeros and the empty
separator. Examples with three decimal places:
  12.3400 -> "12,34"; 12.0000 -> "12"; 1200 -> "1200".
  -0.0004 -> "0" (negative zero is never displayed).
Never use rstrip("0") on an integer: it would turn 1200 into 12.
Angles have independent precision and are not converted to scientific
notation. In significant-digit mode the legacy precision heuristic is
preserved; exponent output is expanded without the old 20-place limit.

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

REFERENCE — MATHCAD 14/15 XML AST CHEAT SHEET (FOR AI MAINTAINERS)
------------------------------------------------------------------
Because no sample `.xmcd` file may be available in future chats, here
are the exact XML structures Mathcad 14/15 emits:

1. Worksheet & Region Layout:
   ```xml
   <worksheet xmlns:ml="http://schemas.mathsoft.com/math30" ...>
     <regions>
       <region region-id="122" left="24" top="134.25">
         <math><ml:define> ... </ml:define></math>
       </region>
       <region region-id="179" left="12" top="728.25">
         <text><p>Текст параграфа</p></text>
       </region>
     </regions>
   </worksheet>
   ```

2. Definitions (`:=`), Numeric Evaluations (`=`), and Symbolic (`->`):
   * Simple definition (`Sb := 100`):
     `<ml:define><ml:id>Sb</ml:id><ml:real>100</ml:real></ml:define>`
   * Definition with numeric evaluation (`Ibaz := expr = 502.04`):
     ```xml
     <ml:define>
       <ml:id>Ibaz</ml:id>
       <ml:eval>
         <ml:apply> ...input expression... </ml:apply>
         <result><ml:real>502.04371233880505</ml:real></result>
       </ml:eval>
     </ml:define>
     ```
   * Definition with `explicit ALL` (`X1 := expr -> symResult = result`):
     ```xml
     <ml:define>
       <ml:id>X1</ml:id>
       <ml:eval>
         <ml:symEval>
           <ml:apply> ...original formula with <ml:id>... </ml:apply>
           <ml:command>
             <ml:sequence>
               <ml:id>explicit</ml:id><ml:id>ALL</ml:id>
             </ml:sequence>
           </ml:command>
           <ml:symResult>
             <ml:apply> ...formula with substituted <ml:real>... </ml:apply>
           </ml:symResult>
         </ml:symEval>
         <result><ml:real>0.30255634724443792</ml:real></result>
       </ml:eval>
     </ml:define>
     ```

3. Complex Numbers & Polar Exponentials in XML:
   * Numeric evaluator result (`<result>`):
     `<ml:complex><ml:real>-0.5</ml:real><ml:imag>0.866</ml:imag></ml:complex>`
   * Symbolic result (`<ml:symResult>`) for a substituted rectangular complex:
     `<ml:parens><ml:apply><ml:plus/><ml:real>-0.5</ml:real><ml:imag symbol="i">0.866</ml:imag></ml:apply></ml:parens>`
   * Polar exponential `1 · e^(j·deg·120)` (in input or `<ml:symResult>`):
     ```xml
     <ml:apply>
       <ml:mult/>
       <ml:real>1</ml:real>
       <ml:apply>
         <ml:pow/>
         <ml:id>e</ml:id>
         <ml:apply>
           <ml:mult/>
           <ml:apply>
             <ml:mult/><ml:imag symbol="j">1</ml:imag><ml:id>deg</ml:id>
           </ml:apply>
           <ml:real>120</ml:real>
         </ml:apply>
       </ml:apply>
     </ml:apply>
     ```

4. Matrices & Operators:
   * `<ml:matrix rows="R" cols="C">` contains `R * C` flat child nodes in
     row-major order (no `<row>` wrapper tags).
   * `<ml:apply>` has the operator tag as `children[0]` and operands as
     `children[1:]`. Built-in operator tags include: `plus`, `minus`, `mult`,
     `div`, `pow`, `sqrt`, `absval`, `conjugate`, `transpose`, `indexer`,
     `factorial`, `neg`, `equal`. Function calls (`sin`, `cos`, `Im`, `Re`,
     `lsolve`, or user functions) have `<ml:id>funcName</ml:id>` as `children[0]`.

5. SUMMARY OF 5 INVARIANTS (NEVER VIOLATE THESE):
   1. NEVER multiply adjacent `<ml:real>` nodes in `_process_mult_angle_args`.
   2. NEVER read `symbol` from `<ml:imag>` (always use `IMAGINARY_SYMBOL`).
   3. NEVER use `rstrip(".")` on subscripts (only replace the last char if `.`).
   4. NEVER use `rstrip("0")` on integer strings without a decimal point.
   5. NEVER call `pop_all_angle_marks()` inside `mult` (use `pop_angle_marks_since`).
"""

import math
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
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

    # Режим «знаки после запятой» — новый режим по умолчанию.
    # Старый режим значащих цифр доступен через rounding_mode="significant".
    rounding_mode = "decimal_places"
    decimal_places = 3
    sig_figs_small = 4
    sig_figs_large = 8

    # Углы имеют отдельную точность; конечные дробные нули удаляются всегда.
    ANGLE_DECIMALS = 1

    # Scientific-notation thresholds:
    #   * sci_threshold: numbers with |x| < 10^sci_threshold -> "m·10^k"
    #   * sci_threshold_large: numbers with |x| >= 10^sci_threshold_large -> "m·10^k"
    # `None` disables the corresponding threshold.
    sci_threshold = None
    sci_threshold_large = None

    # -------------------------------------------------------------------
    # Imaginary-unit symbol
    # -------------------------------------------------------------------
    #
    # The single source of truth for how the imaginary unit is rendered.
    # We FORCE "j" here on purpose; see the module-level note
    # "IMPORTANT — IMAGINARY UNIT" for why the per-node `symbol`
    # attribute cannot be trusted.
    IMAGINARY_SYMBOL = "j"

    # Tolerance for "this value is zero" checks (real/imag parts).
    AXIS_EPS = 1e-12

    # -------------------------------------------------------------------
    # explicit-ALL flag
    # -------------------------------------------------------------------
    #
    # True only while we are parsing a `<symResult>` whose sibling
    # `<command>` contains `explicit ALL`. In that mode, on-axis
    # angles (`e^(j·deg·0°)`, `e^(j·deg·±90°)`, `e^(j·deg·180°)`) are
    # rendered as plain rectangular numbers instead of polar form.
    _explicit_all_mode = False

    # -------------------------------------------------------------------
    # Angle-mark machinery (deferred ∠ substitution)
    # -------------------------------------------------------------------
    #
    # WHY THIS EXISTS:
    # When `pow` encounters `e^(j·deg·φ)`, we want to emit `\angle φ°`.
    # However, the parent `mult` handler needs to inspect the operand
    # list first to decide whether to merge a preceding `<real>1</real>`
    # or insert an implicit `1` modulus (e.g. `Sb \cdot 1\angle φ°`).
    #
    # `pow` therefore pushes the angle LaTeX onto `_angle_mark_stack`
    # and returns an opaque placeholder token `@@ANGLEn@@`.
    #
    # CRITICAL INVARIANT:
    # `mult` must ONLY pop marks pushed by its own operands via
    # `pop_angle_marks_since(depth_before)`. Calling `pop_all_angle_marks`
    # inside `mult` will steal marks belonging to sibling subtrees (such
    # as a numerator's angle mark while parsing a denominator's `mult`),
    # leaving raw `@@ANGLEn@@` text in the final Word document!
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

    @staticmethod
    def _as_decimal(value):
        """Читает XML-число без потери десятичных цифр; принимает также запятую."""
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value).strip().replace(",", "."))

    @staticmethod
    def _trim_fractional_zeros(text):
        """Удаляет только конечные дробные нули и нормализует отрицательный ноль."""
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        if text in ("-0", "+0", ""):
            return "0"
        return text

    @classmethod
    def format_fixed_decimal(cls, value, places):
        """Округляет до places дробных знаков по ROUND_HALF_UP и удаляет нули.

        NOTE: Decimal создаётся из исходной строки, а не из float.
        Контекст расширяется для больших целых и переноса 9.999 -> 10.
        Нули удаляются после округления, без изменения целой части.
        """
        places = int(places)
        if not 0 <= places <= 20:
            raise ValueError("Число знаков после запятой должно быть от 0 до 20.")
        number = cls._as_decimal(value)
        if not number.is_finite():
            return str(number)
        if number.is_zero():
            return "0"
        with localcontext() as context:
            context.prec = max(28, len(number.as_tuple().digits),
                               number.adjusted() + places + 3)
            rounded = number.quantize(Decimal(1).scaleb(-places),
                                      rounding=ROUND_HALF_UP)
        return cls._trim_fractional_zeros(format(rounded, "f")).replace(".", ",")

    @classmethod
    def format_scientific(cls, value, sig_figs=None):
        r"""Выводит m \cdot 10^{k} без конечных нулей в мантиссе.

        В новом режиме decimal_places задаёт дробные знаки мантиссы.
        В старом режиме sig_figs задаёт число её значащих цифр.
        Перенос округлённой мантиссы 10 -> 1 увеличивает показатель.
        """
        number = cls._as_decimal(value)
        if not number.is_finite():
            return str(number)
        if number.is_zero():
            return "0"
        exponent = number.copy_abs().adjusted()
        if cls.rounding_mode == "decimal_places":
            places = cls.decimal_places
        else:
            if sig_figs is None:
                sig_figs = cls._compute_target_sig_figs(number.copy_abs())
            places = max(int(sig_figs) - 1, 0)
        with localcontext() as context:
            context.prec = max(28, len(number.as_tuple().digits) + 3, places + 3)
            mantissa = number.scaleb(-exponent)
            text = cls.format_fixed_decimal(mantissa, places)
            if abs(cls._as_decimal(text)) >= 10:
                exponent += 1
                text = cls.format_fixed_decimal(mantissa / Decimal(10), places)
        if exponent == 0:
            return text
        return rf"{text} \cdot 10^{{{exponent}}}"

    @classmethod
    def format_num(cls, num_str):
        """Форматирует числовые узлы: входные формулы, результаты и матрицы.

        decimal_places: не более N дробных знаков, без конечных нулей.
        significant: прежняя эвристика значащих цифр, включая бонус для 1/2.
        Научная форма выбирается по исходному модулю ДО округления
        (для малых чисел при |x| < 10^sci_threshold, для больших при
        |x| >= 10^sci_threshold_large).
        Нечисловая строка возвращается с заменой точки на запятую.
        """
        if num_str is None or str(num_str).strip() == "":
            return ""
        try:
            number = cls._as_decimal(num_str)
        except (InvalidOperation, ValueError, TypeError):
            return str(num_str).replace(".", ",")
        if not number.is_finite():
            return str(number)
        if number.is_zero():
            return "0"
        absolute = number.copy_abs()
        if cls.sci_threshold is not None:
            threshold = Decimal(1).scaleb(int(cls.sci_threshold))
            if absolute < threshold:
                return cls.format_scientific(number)
        if cls.sci_threshold_large is not None:
            threshold_large = Decimal(1).scaleb(int(cls.sci_threshold_large))
            if absolute >= threshold_large:
                return cls.format_scientific(number)
        if cls.rounding_mode == "decimal_places":
            return cls.format_fixed_decimal(number, cls.decimal_places)
        if cls.rounding_mode != "significant":
            raise ValueError("Неизвестный режим округления.")
        target = cls._compute_target_sig_figs(absolute)
        # Сохраняем старое округление через float для обычных конечных значений.
        value = float(number)
        if math.isfinite(value) and value != 0:
            formatted = f"{value:.{target}g}"
        else:
            with localcontext() as context:
                context.prec = target
                formatted = str(+number)
        # Decimal раскрывает экспоненту точно, не обнуляя значения < 1e-20.
        plain = format(Decimal(formatted), "f")
        return cls._trim_fractional_zeros(plain).replace(".", ",")

    @classmethod
    def format_angle_value(cls, angle):
        """Угол: отдельная точность ANGLE_DECIMALS, без конечных дробных нулей.

        NOTE: Не зависит от режима обычных чисел и научного порога.
        """
        try:
            number = cls._as_decimal(angle)
        except (InvalidOperation, ValueError, TypeError):
            return str(angle).replace(".", ",")
        return cls.format_fixed_decimal(number, cls.ANGLE_DECIMALS)

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
            We replace ONLY THE LAST character of the subscript — and
            only when it is a dot — with a single `*`. Any dots that
            appear earlier inside the subscript are preserved:
                subscript="1."       →  "1*"
                subscript="2'.1.."   →  "2'.1.*"
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
        only on the first/last line.
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
        Mathcad's symbolic engine ALWAYS re-emits the imaginary unit
        with symbol="i" inside `<symResult>`, even if the user wrote `j`.
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
            (real or imaginary) is numerically zero — we fall back to
            the rectangular form: `-5`, `5`, `5j`, `-5j`.
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

        # Both parts non-zero.
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
    # -------------------------------------------------------------------

    @staticmethod
    def _strip_leading_sign(text):
        """
        Split a LaTeX fragment into a leading sign and the remainder.

        Returns (sign, body) where:
          * sign is "+" or "-";
          * body is the fragment without its leading sign (and
            without leading whitespace).
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
    def _is_on_axis_euler_node(cls, node):
        """
        Return True if `node` is an `<apply><pow/>` representing an
        Euler exponential `e^(j·deg·φ)` with a numeric on-axis angle φ
        (0°, ±90°, 180°).

        WHY THIS EXISTS (CRITICAL FOR MAINTAINERS)
        ------------------------------------------
        In `_explicit_all_mode`, `e^(j·deg·180°)` is rendered by `pow`
        as `"-1"` (and `e^(j·deg·0°)` as `"1"`). When preceded by a
        numeric modulus in a polar product like `100 · e^(j·deg·180°)`,
        we want to fold the ±1 into the modulus (`-100`).

        HOWEVER, we must NEVER fold two ordinary `<real>` numbers!
        In `<symResult>` after `explicit ALL`, a product of two real
        variables `a · b` (e.g. `a = 3`, `b = 4`) is represented as
        `<apply><mult/><real>3</real><real>4</real></apply>`. Both
        operands must remain separate (`3 \\cdot 4`) so the reader sees
        the substituted value of each variable. Checking
        `_is_on_axis_euler_node(node)` guarantees that only genuine
        on-axis Euler exponentials are folded into their modulus.
        """
        if node is None or cls.strip_ns(node.tag) != "apply":
            return False

        children = list(node)
        if len(children) < 3 or cls.strip_ns(children[0].tag) != "pow":
            return False

        if not cls.is_euler_base(children[1]):
            return False

        angle_value = cls.extract_polar_angle_numeric(children[2])
        if angle_value is None:
            return False

        return cls.angle_to_rectangular(angle_value) is not None

    @classmethod
    def parse_complex_node(cls, node):
        """
        Parse a <complex> element: <real>r</real> <imag>m</imag>.
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
        """
        mark = f"@@ANGLE{cls._angle_mark_counter}@@"
        cls._angle_mark_counter += 1
        cls._angle_mark_stack.append({mark: latex})
        return mark

    @classmethod
    def pop_angle_marks_since(cls, depth):
        """
        Pop and return every mark pushed after the given stack depth.
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
        """
        return r"\angle" in latex or latex.rstrip().endswith(MathcadParser.IMAGINARY_SYMBOL)

    # -------------------------------------------------------------------
    # mult-operator post-processing
    # -------------------------------------------------------------------

    @staticmethod
    def _is_bare_angle_mark(text):
        """
        True if `text` is a bare angle mark (unresolved) or a plain
        resolved angle like "1\\angle 84^\\circ" without extra factors.
        """
        stripped = text.strip()

        if MathcadParser.ANGLE_MARK_RE.fullmatch(stripped):
            return True

        if re.fullmatch(
            r"1?\\angle.*?\^\s*(?:\\circ|\{\\circ\})",
            stripped
        ):
            return True

        return False

    @staticmethod
    def _is_angle_mark_wrapped(text):
        r"""
        True if `text` is an angle fragment (resolved or unresolved),
        optionally wrapped in \left( ... \right).
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
        """True if text is purely a number (plain or scientific notation, with comma)."""
        return bool(
            re.fullmatch(
                r"[+-]?[\d,.]+(?:\s*\\cdot\s*10\^\{[+-]?\d+\})?",
                text.strip()
            )
        )

    @staticmethod
    def _is_simple_coefficient(text):
        """True for a number (plain or scientific) or a plain identifier (with optional subscript)."""
        stripped = text.strip()
        if not stripped:
            return False
        if MathcadParser._is_pure_number_latex(stripped):
            return True
        if re.fullmatch(r"[A-Za-z][A-Za-z0-9]*(_\{[^}]*\})?", stripped):
            return True
        return False

    @classmethod
    def _process_mult_angle_args(cls, args, children):
        r"""
        Post-process a chain of `mult` operands to produce clean LaTeX.

        This function:
          1. Merges the explicit `1` in `1 · ∠120°` so we don't emit
             "1 · 1∠120°".
          2. Decides where an implicit `1` modulus is needed:
               - pure number before the angle → no extra "1";
               - anything else → prepend "1".
          3. In explicit-ALL mode, folds an on-axis Euler exponential
             (`e^(j·deg·0°)` → `"1"`, `e^(j·deg·180°)` → `"-1"`) into
             its preceding numeric modulus so `100 · e^(j·deg·180°)`
             becomes `-100` (and drops a unit `1` before `±j`).
             CRITICAL: It NEVER multiplies two ordinary `<real>` numbers!
             If `a = 3` and `b = 4`, `a · b` after `explicit ALL` must
             stay `3 \cdot 4`.
          4. Joins everything with `\cdot` — except where a simple
             coefficient directly precedes a bare angle mark.
          5. Wraps a minus-leading right operand in `\left(...\right)`.
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
                "is_on_axis_euler": (
                    cls._explicit_all_mode
                    and cls._is_on_axis_euler_node(node)
                ),
                "has_one_angle": bool(
                    re.search(r"(?<!\d)1\\angle", arg)
                ),
                "keep": True,
                "prefix_one": False,
            })

        # Step 1: merge the explicit 1 that appears immediately before
        # an angle.
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

        # Step 2.5 (explicit-ALL only): fold an on-axis Euler angle
        # factor (`e^(j·deg·0°)` -> "1", `e^(j·deg·180°)` -> "-1") into
        # its preceding numeric modulus (e.g. `100 · e^(j·deg·180°)` ->
        # `-100`), or drop an explicit unit `1` before `±j`.
        #
        # CRITICAL: Do NOT merge two ordinary numeric factors! In
        # `<symResult>` after `explicit ALL`, `a · b` (with `a=3, b=4`)
        # is `<mult/><real>3</real><real>4</real>`. Neither `<real>`
        # node has `is_on_axis_euler == True`, so `3 \cdot 4` is
        # preserved intact.
        if cls._explicit_all_mode:
            for i, f in enumerate(factors):
                if not f["keep"] or not f["is_on_axis_euler"]:
                    continue

                prev_idx = -1
                for j in range(i - 1, -1, -1):
                    if factors[j]["keep"] and factors[j]["text"].strip():
                        prev_idx = j
                        break

                if prev_idx == -1:
                    continue

                prev = factors[prev_idx]
                euler_text = f["text"].strip()

                if prev["is_pure_number"] and not prev["is_on_axis_euler"]:
                    if euler_text == "1":
                        f["keep"] = False
                    elif euler_text == "-1":
                        prev["text"] = cls._apply_negation(prev["text"])
                        f["keep"] = False
                    elif prev["is_unit_one"] and euler_text in (
                        cls.IMAGINARY_SYMBOL,
                        f"-{cls.IMAGINARY_SYMBOL}",
                    ):
                        prev["keep"] = False

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

            if (cls._is_bare_angle_mark(text)
                    and cls._is_simple_coefficient(result.strip())):
                result += text
                continue

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
        if tag == "apply":
            if not children:
                return ""

            op_node = children[0]
            op = cls.strip_ns(op_node.tag)

            # ---- Special-case: a + bj → polar form ----
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

            # Capture angle-mark stack depth BEFORE parsing operands.
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
            if op == "minus":
                if not args:
                    return "-"

                if len(args) == 1:
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
            if op == "neg":
                if not args:
                    return "-"
                return cls._apply_negation(args[0])

            # ---- exponentiation ----
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

                if args:
                    base_latex = args[0]
                    # NOTE: If a single <ml:real> base was formatted into
                    # scientific notation (m \cdot 10^{k}), Mathcad's XML
                    # has no <ml:parens> around it. Wrap it in \left(...\right)
                    # so (m \cdot 10^{k})^{2} does not read as m \cdot 10^{k^2}.
                    if (
                        r"\cdot 10^{" in base_latex
                        and not base_latex.strip().startswith(r"\left(")
                    ):
                        base_latex = rf"\left({base_latex}\right)"

                    if len(args) > 1:
                        return f"{{{base_latex}}}^{{{args[1]}}}"
                    return f"{{{base_latex}}}^{{?}}"
                return ""

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

            # ---- factorial (n!) ----
            if op == "factorial":
                if not args:
                    return "!"
                return f"{args[0]}!"

            # ---- user-defined function call ----
            if op in ("id", "sym"):
                function_name = cls.parse_identifier(
                    op_node, function=True
                )
                return (
                    rf"{function_name}\left("
                    + ", ".join(args)
                    + r"\right)"
                )

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
        if tag == "sequence":
            return ", ".join(
                cls.parse_node_to_latex(child)
                for child in children
            )

        # --- empty placeholder ----------------------------------------
        if tag == "placeholder":
            return r"\square"

        # --- nested eval / symEval ------------------------------------
        if tag in ("eval", "symEval"):
            return cls.parse_eval_to_latex(node)

        # --- fallback: try each child in order ------------------------
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
        """
        result_node = cls.find_first_by_tag(node, "result")
        sym_eval_node = cls.find_first_by_tag(node, "symEval")

        parts = []

        if sym_eval_node is not None:
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
        """
        # Ensure clean state in case a previous conversion raised an error.
        cls._explicit_all_mode = False
        cls._angle_mark_stack.clear()
        cls._angle_mark_counter = 0

        try:
            tree = ET.parse(filepath)
            root = tree.getroot()
        except Exception as exc:
            raise ValueError(
                f"Ошибка чтения XML/XMCD файла: {exc}"
            ) from exc

        regions = [
            element
            for element in root.iter()
            if cls.strip_ns(element.tag) == "region"
        ]

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
                    if len(child) >= 2:
                        left_side = cls.parse_node_to_latex(child[0])
                        right_side = child[1]

                        if cls.strip_ns(right_side.tag) in ("eval", "symEval"):
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

                elif tag in ("eval", "symEval"):
                    latex = cls.parse_eval_to_latex_root(child)

                else:
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
        self.root.geometry("680x980")
        self.root.minsize(640, 860)
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

        # ---- Настройки отображения чисел и независимая точность углов ----
        frame_settings = ttk.LabelFrame(
            self.root, text="Настройки округления", padding=10
        )
        frame_settings.pack(fill=tk.X, pady=(0, 15))
        frame_settings.columnconfigure(0, weight=1)

        self.rounding_mode_var = tk.StringVar(value="decimal_places")
        self.decimal_places_var = tk.IntVar(value=3)
        self.angle_decimals_var = tk.IntVar(value=1)
        self.sig_figs_small_var = tk.IntVar(value=4)
        self.sig_figs_large_var = tk.IntVar(value=8)

        ttk.Radiobutton(
            frame_settings, text="По числу знаков после запятой",
            variable=self.rounding_mode_var, value="decimal_places",
            command=self.toggle_rounding_mode
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Label(frame_settings, text="Максимум знаков после запятой:").grid(
            row=1, column=0, sticky=tk.W, pady=2
        )
        self.decimal_spinbox = ttk.Spinbox(
            frame_settings, from_=0, to=20,
            textvariable=self.decimal_places_var, width=6
        )
        self.decimal_spinbox.grid(row=1, column=1, padx=10, pady=2)
        ttk.Label(
            frame_settings,
            text="Конечные нули удаляются: 12,340 → 12,34; 12,000 → 12.",
            foreground="gray", wraplength=590
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(2, 6))

        ttk.Radiobutton(
            frame_settings, text="По значащим цифрам (прежний режим)",
            variable=self.rounding_mode_var, value="significant",
            command=self.toggle_rounding_mode
        ).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Label(frame_settings, text="Для малых чисел и дробей:").grid(
            row=4, column=0, sticky=tk.W, pady=2
        )
        self.sig_small_spinbox = ttk.Spinbox(
            frame_settings, from_=1, to=15,
            textvariable=self.sig_figs_small_var, width=6
        )
        self.sig_small_spinbox.grid(row=4, column=1, padx=10, pady=2)
        ttk.Label(frame_settings, text="Для больших чисел:").grid(
            row=5, column=0, sticky=tk.W, pady=2
        )
        self.sig_large_spinbox = ttk.Spinbox(
            frame_settings, from_=1, to=20,
            textvariable=self.sig_figs_large_var, width=6
        )
        self.sig_large_spinbox.grid(row=5, column=1, padx=10, pady=2)
        ttk.Label(frame_settings, text="Знаков после запятой в углах:").grid(
            row=6, column=0, sticky=tk.W, pady=2
        )
        ttk.Spinbox(
            frame_settings, from_=0, to=20,
            textvariable=self.angle_decimals_var, width=6
        ).grid(row=6, column=1, padx=10, pady=2)
        self.toggle_rounding_mode()

        ttk.Separator(frame_settings, orient="horizontal").grid(
            row=7, column=0, columnspan=2, sticky="ew", pady=(8, 6)
        )
        self.use_sci_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frame_settings, text="Научная форма для малых чисел",
            variable=self.use_sci_var, command=self.toggle_sci
        ).grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=2)
        ttk.Label(frame_settings, text="Порог 10^N для малых чисел (например -4):").grid(
            row=9, column=0, sticky=tk.W, pady=2
        )
        self.sci_threshold_var = tk.IntVar(value=-4)
        self.sci_spinbox = ttk.Spinbox(
            frame_settings, from_=-15, to=0,
            textvariable=self.sci_threshold_var, width=6, state="disabled"
        )
        self.sci_spinbox.grid(row=9, column=1, padx=10, pady=2)

        self.use_sci_large_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frame_settings, text="Научная форма для больших чисел",
            variable=self.use_sci_large_var, command=self.toggle_sci_large
        ).grid(row=10, column=0, columnspan=2, sticky=tk.W, pady=(6, 2))
        ttk.Label(frame_settings, text="Порог 10^M для больших чисел (например 4):").grid(
            row=11, column=0, sticky=tk.W, pady=2
        )
        self.sci_large_threshold_var = tk.IntVar(value=4)
        self.sci_large_spinbox = ttk.Spinbox(
            frame_settings, from_=1, to=20,
            textvariable=self.sci_large_threshold_var, width=6, state="disabled"
        )
        self.sci_large_spinbox.grid(row=11, column=1, padx=10, pady=2)

        ttk.Label(
            frame_settings,
            text=("При |x| < 10^N (малые) или |x| ≥ 10^M (большие) используется "
                  "m·10^k. В режиме знаков после запятой выбранная точность "
                  "применяется к мантиссе m."),
            foreground="gray", wraplength=590, justify=tk.LEFT
        ).grid(row=12, column=0, columnspan=2, sticky=tk.W, pady=(4, 0))

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

        self.log("Проверка окружения...")
        self.root.after(
            500,
            lambda: WordConverter.check_pandoc(self.log)
        )

    def toggle_rounding_mode(self):
        """Оставляет активными только поля выбранного режима округления."""
        fixed = self.rounding_mode_var.get() == "decimal_places"
        self.decimal_spinbox.config(state="normal" if fixed else "disabled")
        state = "disabled" if fixed else "normal"
        self.sig_small_spinbox.config(state=state)
        self.sig_large_spinbox.config(state=state)

    @staticmethod
    def _read_bounded_int(variable, label, minimum, maximum):
        """Проверяет ручной ввод: пределы Spinbox сами его не ограничивают."""
        try:
            value = variable.get()
        except (tk.TclError, ValueError, TypeError) as exc:
            raise ValueError(f"{label}: введите целое число.") from exc
        if not isinstance(value, int) or not minimum <= value <= maximum:
            raise ValueError(f"{label}: допустимы числа от {minimum} до {maximum}.")
        return value

    def read_formatting_settings(self):
        """Возвращает проверенные настройки, не изменяя состояние парсера."""
        mode = self.rounding_mode_var.get()
        if mode not in ("decimal_places", "significant"):
            raise ValueError("Выберите режим округления.")
        settings = {
            "rounding_mode": mode,
            "ANGLE_DECIMALS": self._read_bounded_int(
                self.angle_decimals_var, "Точность углов", 0, 20
            ),
            "sci_threshold": None,
            "sci_threshold_large": None,
        }
        if mode == "decimal_places":
            settings["decimal_places"] = self._read_bounded_int(
                self.decimal_places_var, "Знаки после запятой", 0, 20
            )
        else:
            settings["sig_figs_small"] = self._read_bounded_int(
                self.sig_figs_small_var, "Значащие цифры для малых чисел", 1, 15
            )
            settings["sig_figs_large"] = self._read_bounded_int(
                self.sig_figs_large_var, "Значащие цифры для больших чисел", 1, 20
            )
        if self.use_sci_var.get():
            settings["sci_threshold"] = self._read_bounded_int(
                self.sci_threshold_var, "Порог научной формы (малые числа)", -15, 0
            )
        if self.use_sci_large_var.get():
            settings["sci_threshold_large"] = self._read_bounded_int(
                self.sci_large_threshold_var, "Порог научной формы (большие числа)", 1, 20
            )
        return settings

    def toggle_sci(self):
        """Enable/disable the small-number sci-threshold spinbox."""
        self.sci_spinbox.config(
            state="normal" if self.use_sci_var.get() else "disabled"
        )

    def toggle_sci_large(self):
        """Enable/disable the large-number sci-threshold spinbox."""
        self.sci_large_spinbox.config(
            state="normal" if self.use_sci_large_var.get() else "disabled"
        )

    def log(self, message):
        """Append a line to the log pane and scroll to the bottom."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

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
        paths = self.root.tk.splitlist(event.data)
        if paths:
            self.set_input_file(paths[0])

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

        try:
            settings = self.read_formatting_settings()
        except ValueError as exc:
            messagebox.showerror("Ошибка настроек округления", str(exc))
            return

        output_file = filedialog.asksaveasfilename(
            title="Сохранить результат как...",
            defaultextension=".docx",
            initialfile="Результат_Mathcad.docx",
            filetypes=[("Word Document", "*.docx")]
        )
        if not output_file:
            return

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

        for name, value in settings.items():
            setattr(MathcadParser, name, value)
        if MathcadParser.rounding_mode == "decimal_places":
            self.log(
                f"[INFO] Округление: до {MathcadParser.decimal_places} "
                "знаков после запятой; конечные нули удаляются."
            )
        else:
            self.log(
                f"[INFO] Значащие цифры: малые = {MathcadParser.sig_figs_small}, "
                f"большие = {MathcadParser.sig_figs_large}."
            )
        self.log(f"[INFO] Точность углов: {MathcadParser.ANGLE_DECIMALS} знаков.")
        if MathcadParser.sci_threshold is not None:
            self.log(
                f"[INFO] Научная форма (малые): |x| < 10^{MathcadParser.sci_threshold}."
            )
        else:
            self.log("[INFO] Научная форма для малых чисел отключена.")
        if MathcadParser.sci_threshold_large is not None:
            self.log(
                f"[INFO] Научная форма (большие): |x| >= 10^{MathcadParser.sci_threshold_large}."
            )
        else:
            self.log("[INFO] Научная форма для больших чисел отключена.")

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