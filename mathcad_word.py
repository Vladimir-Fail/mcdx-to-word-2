import math
import os
import re
import sys
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pypandoc


try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_SUPPORTED = True
except ImportError:
    DND_SUPPORTED = False


class MathcadParser:
    """Разбор XML/XMCD Mathcad и преобразование в Markdown/LaTeX."""

    sig_figs_small = 4
    sig_figs_large = 8

    @staticmethod
    def strip_ns(tag):
        """Удаляет namespace из XML-тега или атрибута."""
        if isinstance(tag, str) and "}" in tag:
            return tag.split("}", 1)[1]

        return tag

    @classmethod
    def format_num(cls, num_str):
        """Форматирование чисел со значащими цифрами."""
        if not num_str:
            return ""

        try:
            value = float(num_str)

            if value == 0:
                return "0"

            abs_value = abs(value)

            sci_str = f"{abs_value:.15e}"
            first_digit = sci_str[0]

            if abs_value >= 1:
                int_digits = len(str(int(abs_value)))

                base_sig_figs = max(
                    cls.sig_figs_small,
                    min(cls.sig_figs_large, int_digits)
                )
            else:
                base_sig_figs = cls.sig_figs_small

            if first_digit in ("1", "2"):
                target_sig_figs = base_sig_figs + 1
            else:
                target_sig_figs = base_sig_figs

            formatted = f"{value:.{target_sig_figs}g}"

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
    def find_first_by_tag(cls, node, tag):
        """Ищет первый дочерний элемент по имени тега."""
        if node is None:
            return None

        for child in node.iter():
            if cls.strip_ns(child.tag) == tag:
                return child

        return None

    @staticmethod
    def escape_latex(text):
        """Экранирует специальные символы LaTeX."""
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

    @classmethod
    def parse_identifier(cls, node, function=False):
        """
        Обрабатывает имя переменной или функции.

        Примеры:

            R.max -> R_{max}
            x.1   -> x_{1}
        """
        if node is None:
            return "?"

        base = (node.text or "").strip()
        subscript = None

        # Индекс может быть записан атрибутом XML.
        for key, value in node.attrib.items():
            if cls.strip_ns(key) == "subscript":
                subscript = value
                break

        # Индекс может быть отдельным дочерним узлом.
        for child in node:
            tag = cls.strip_ns(child.tag)

            if tag in ("subscript", "sub"):
                subscript = "".join(child.itertext()).strip()

            elif tag in ("name", "id", "sym") and not base:
                base = "".join(child.itertext()).strip()

        # Основной вариант Mathcad:
        # имя с точкой является именем с литеральным индексом.
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
        """Преобразует текстовый результат функции в LaTeX."""
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

    @staticmethod
    def normalize_angle_degrees(angle):
        """
        Нормализует угол в диапазон (-180, 180].

        Возможные значения:
            180  -> 180
            -180 -> 180
            270  -> -90
            360  -> 0
        """
        # Сначала убираем возможные ошибки накопления
        # вычислений около границ диапазона.
        if math.isclose(angle, 0.0, abs_tol=1e-12):
            return 0.0

        angle = math.fmod(angle, 360.0)

        if angle > 180.0:
            angle -= 360.0

        elif angle <= -180.0:
            angle += 360.0

        # Повторно убираем очень маленькие значения вроде
        # -1.4e-14 после вычислений.
        if math.isclose(angle, 0.0, abs_tol=1e-12):
            return 0.0

        if math.isclose(abs(angle), 180.0, abs_tol=1e-12):
            return 180.0

        return angle

    @classmethod
    def calculate_complex_polar(cls, real_value, imag_value):
        """
        Вычисляет модуль и угол комплексного числа.

        Для z = a + ib:

            |z| = hypot(a, b)
            angle = atan2(b, a) в градусах

        Результат:
            модуль, угол
        """
        real_value = float(real_value)
        imag_value = float(imag_value)

        # math.hypot точнее и устойчивее, чем:
        # sqrt(real_value ** 2 + imag_value ** 2)
        modulus = math.hypot(
            real_value,
            imag_value
        )

        # ВАЖНО:
        # первый аргумент — мнимая часть,
        # второй аргумент — действительная часть.
        angle = math.degrees(
            math.atan2(
                imag_value,
                real_value
            )
        )

        angle = cls.normalize_angle_degrees(angle)

        return modulus, angle

    @classmethod
    def parse_complex(cls, real_value, imag_value):
        """
        Преобразует a + ib в:

            |z| angle angle°

        Например:

            1 + i  -> 1,414 angle 45°
            1 - i  -> 1,414 angle -45°
        """
        try:
            modulus, angle = cls.calculate_complex_polar(
                real_value,
                imag_value
            )

        except (ValueError, TypeError):
            # Если Mathcad сохранил нечисловой результат,
            # оставляем исходный декартов вид.
            real_text = cls.format_num(
                str(real_value)
            )

            imag_text = cls.format_num(
                str(imag_value)
            )

            sign = (
                ""
                if imag_text.startswith("-")
                else "+"
            )

            return f"{real_text}{sign}{imag_text}i"

        modulus_text = cls.format_num(
            str(modulus)
        )

        angle_text = cls.format_num(
            str(angle)
        )

        return (
            rf"{modulus_text}"
            rf"\angle "
            rf"{angle_text}^\circ"
        )

    @staticmethod
    def is_degree_unit(node):
        """
        Проверяет, что XML-узел — единица измерения deg.

        Пример Mathcad: <id>deg</id>
        """
        if node is None:
            return False

        if MathcadParser.strip_ns(node.tag) != "id":
            return False

        return (node.text or "").strip() == "deg"

    @classmethod
    def is_euler_base(cls, node):
        """
        Проверяет, что узел — константа e (основание
        экспоненты).

        Mathcad записывает её как <id>e</id> или <e/>.
        """
        if node is None:
            return False

        tag = cls.strip_ns(node.tag)

        if tag == "e":
            return True

        if tag in ("id", "sym"):
            base = (node.text or "").strip()

            # Имя с индексом — это переменная, а не константа e
            for key in node.attrib:
                if cls.strip_ns(key) == "subscript":
                    return False

            return base == "e"

        return False

    @classmethod
    def extract_polar_angle_degrees(cls, node):
        """
        Пытается распознать показатель степени вида
        1i·deg·φ (запись угла в градусах в экспоненциальной
        форме e^(1i·deg·φ)).

        Структура XML Mathcad:

            <apply>
                <mult/>
                <apply>
                    <mult/>
                    <imag symbol="j">1</imag>   (или 1i / 1j)
                    <id>deg</id>
                </apply>
                <real>120</real>               (сам угол)
            </apply>

        Возвращает LaTeX-строку угла (число или выражение)
        либо None, если узел не является полярным показателем.
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
            """Группа 1i · deg (порядок множителей любой)."""
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
        """
        Преобразует угол φ из показателя exp(1i·deg·φ)
        в запись с символом угла: ∠φ°.
        """
        if not angle_latex:
            return r"\angle"

        return rf"\angle\ {angle_latex}^{{\circ}}"

    @classmethod
    def parse_complex_node(cls, node):
        """Извлекает real и imag из XML-узла complex."""
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

        return cls.parse_complex(
            real_value,
            imag_value
        )

    @staticmethod
    def is_complex_latex(latex):
        """Проверяет, содержит ли LaTeX-строка комплексное число."""
        return r"\angle" in latex or latex.rstrip().endswith("i")

    MATRIX_MAX_WIDTH = 1000

    @classmethod
    def build_matrix_latex(cls, rows, cols, delimiter="bmatrix"):
        r"""
        Собирает LaTeX-матрицу из уже распознанных ячеек.

        Если матрица широкая (например 5x9), строки разбиваются
        на несколько блоков pmatrix, чтобы формула помещалась
        на странице Word и не обрезалась при переносе строк.
        """
        if not rows or not cols:
            return ""

        total_width = sum(
            max(
                len(cell)
                for cell in column
            )
            for column in zip(*rows)
        ) + 2 * (len(cols) - 1)

        # Узкая матрица выводится целиком.
        if total_width <= cls.MATRIX_MAX_WIDTH:
            body = r" \\ ".join(
                " & ".join(row)
                for row in rows
            )

            return rf"\begin{{{delimiter}}}{body}\end{{{delimiter}}}"

        # Разбиение по столбцам: подбираем ширину блока так,
        # чтобы каждая блок-матрица умещалась в заданную ширину.
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
        Преобразует XML-узел matrix Mathcad в LaTeX-матрицу.

        Ячейки хранятся в потоке flat row-major; количество
        ячеек равно rows * cols. Пустые узлы <matrix/>
        (настройки отображения) игнорируются.
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

        # На случай несоответствия атрибутов числу ячеек.
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

        # Ширина каждой колонки в символах (для разбиения).
        col_widths = [
            max(len(cell[i]) for cell in rows)
            for i in range(cols_count)
        ]

        # Комплексные числа переводятся в показательную
        # форму, из-за чего ячейки становятся широкими.
        # Для таких матриц блоки формируем по одной
        # колонке за раз.
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
                    rows,
                    col_widths,
                    delimiter="bmatrix"
                )
            finally:
                cls.MATRIX_MAX_WIDTH = saved_max_width

        return cls.build_matrix_latex(
            rows,
            col_widths,
            delimiter="bmatrix"
        )

    @classmethod
    def parse_node_to_latex(cls, node):
        """Рекурсивно преобразует XML-узел Mathcad в LaTeX."""
        if node is None:
            return "?"

        tag = cls.strip_ns(node.tag)
        children = list(node)

        if tag == "math":
            return " ".join(
                cls.parse_node_to_latex(child)
                for child in children
            )

        # Обычные числа.
        if tag == "real":
            return cls.format_num(
                node.text.strip()
                if node.text
                else ""
            )

        # Имена переменных и функций.
        if tag in ("id", "sym"):
            return cls.parse_identifier(node)

        # Текстовые значения.
        if tag in ("str", "string"):
            return cls.parse_string(node)

        # Комплексное число.
        if tag == "complex":
            return cls.parse_complex_node(node)

        # Результаты вычислений.
        if tag in ("result", "symResult"):
            parts = [
                cls.parse_node_to_latex(child)
                for child in children
            ]

            parts = [
                part for part in parts
                if part and part != "?"
            ]

            if parts:
                return " ".join(parts)

            if node.text and node.text.strip():
                return cls.parse_string(node)

            return ""

        # Отдельная мнимая часть.
        if tag == "imag":
            symbol = node.attrib.get(
                "symbol",
                "i"
            )

            number = cls.format_num(
                node.text.strip()
                if node.text
                else ""
            )

            return f"{number}{symbol}"

        # Матрицы и векторы Mathcad.
        if tag == "matrix":
            return cls.parse_matrix_node(node, children)

        if tag == "parens":
            child_latex = (
                cls.parse_node_to_latex(children[0])
                if children
                else ""
            )

            return rf"\left({child_latex}\right)"

        if tag == "apply":
            if not children:
                return ""

            op_node = children[0]
            op = cls.strip_ns(op_node.tag)

            # Частый вариант Mathcad 15:
            #
            # <apply>
            #     <plus />
            #     <real>a</real>
            #     <imag>b</imag>
            # </apply>
            #
            # Это означает a + ib.
            if op == "plus" and len(children) >= 3:
                first_node = children[1]
                second_node = children[2]

                first_tag = cls.strip_ns(
                    first_node.tag
                )

                second_tag = cls.strip_ns(
                    second_node.tag
                )

                if (
                    first_tag == "real"
                    and second_tag == "imag"
                ):
                    real_value = (
                        first_node.text.strip()
                        if first_node.text
                        else "0"
                    )

                    imag_value = (
                        second_node.text.strip()
                        if second_node.text
                        else "0"
                    )

                    return cls.parse_complex(
                        real_value,
                        imag_value
                    )

            args = [
                cls.parse_node_to_latex(child)
                for child in children[1:]
            ]

            if op == "mult":
                if len(args) > 1:
                    return (
                        f"{args[0]} \\cdot {args[1]}"
                    )

                return args[0] if args else ""

            if op == "div":
                if len(args) > 1:
                    return (
                        f"\\frac{{{args[0]}}}"
                        f"{{{args[1]}}}"
                    )

                return (
                    f"\\frac{{{args[0]}}}{{?}}"
                    if args
                    else ""
                )

            if op == "plus":
                if len(args) > 1:
                    return f"{args[0]} + {args[1]}"

                return args[0] if args else ""

            if op == "minus":
                if len(args) > 1:
                    return f"{args[0]} - {args[1]}"

                return f"-{args[0]}" if args else "-"

            if op == "neg":
                return f"-{args[0]}" if args else "-"

            if op == "pow":
                # Полярная запись угла Mathcad:
                #
                #     e^(1i·deg·φ)  ->  ∠φ°
                #
                # Проверяем исходные XML-узлы, а не LaTeX:
                # основание — константа e, показатель —
                # произведение 1i·deg и самого угла.
                base_node = (
                    children[1]
                    if len(children) > 1
                    else None
                )

                exp_node = (
                    children[2]
                    if len(children) > 2
                    else None
                )

                if cls.is_euler_base(base_node):
                    angle_latex = cls.extract_polar_angle_degrees(
                        exp_node
                    )

                    if angle_latex is not None:
                        return cls.format_angle_symbol(
                            angle_latex
                        )

                if len(args) > 1:
                    return (
                        f"{{{args[0]}}}"
                        f"^{{{args[1]}}}"
                    )

                return (
                    f"{{{args[0]}}}^{{?}}"
                    if args
                    else ""
                )

            if op == "sqrt":
                return (
                    f"\\sqrt{{{args[0]}}}"
                    if args
                    else "\\sqrt{?}"
                )

            if op == "absval":
                return (
                    f"\\left| {args[0]} \\right|"
                    if args
                    else "\\left| ? \\right|"
                )

            if op == "conjugate":
                return (
                    f"\\overline{{{args[0]}}}"
                    if args
                    else "\\overline{?}"
                )

            # Индекс элемента матрицы/вектора: A[i, j].
            if op == "indexer":
                if not args:
                    return ""

                index_latex = ", ".join(args[1:])

                return rf"{args[0]}_{{\text[{index_latex}]}}"

            # Транспонирование матрицы: A^T.
            if op == "transpose":
                return (
                    f"{{{args[0]}}}^{{T}}"
                    if args
                    else "^{T}"
                )

            if op == "equal":
                if len(args) > 1:
                    return (
                        f"{args[0]} = {args[1]}"
                    )

                return (
                    f"{args[0]} = ?"
                    if args
                    else "="
                )

            # Вызов функции.
            if op in ("id", "sym"):
                function_name = cls.parse_identifier(
                    op_node,
                    function=True
                )

                return (
                    rf"{function_name}\left("
                    + ", ".join(args)
                    + r"\right)"
                )

            return "?"

        if tag == "function":
            bound_vars = next(
                (
                    child for child in children
                    if cls.strip_ns(child.tag)
                    == "boundVars"
                ),
                None
            )

            function_node = next(
                (
                    child for child in children
                    if cls.strip_ns(child.tag)
                    in ("id", "sym")
                ),
                None
            )

            function_latex = (
                cls.parse_identifier(
                    function_node,
                    function=True
                )
                if function_node is not None
                else ""
            )

            variables_latex = (
                ", ".join(
                    cls.parse_node_to_latex(child)
                    for child in bound_vars
                )
                if bound_vars is not None
                else ""
            )

            return (
                rf"{function_latex}\left("
                + variables_latex
                + r"\right)"
            )

        if tag == "sequence":
            return ", ".join(
                cls.parse_node_to_latex(child)
                for child in children
            )

        if tag == "placeholder":
            return r"\square"

        # Обёртки provenance, unitedValue и т. п.
        if children:
            for child in children:
                result = cls.parse_node_to_latex(
                    child
                )

                if result and result != "?":
                    return result

        return ""

    @classmethod
    def parse_eval_to_latex(cls, node):
        """
        Обрабатывает eval.

        Возможный результат:

            выражение = символьный результат = итог
        """
        result_node = cls.find_first_by_tag(
            node,
            "result"
        )

        sym_eval_node = cls.find_first_by_tag(
            node,
            "symEval"
        )

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
                sym_eval_node,
                "symResult"
            )

            if expression_node is not None:
                expression_latex = (
                    cls.parse_node_to_latex(
                        expression_node
                    )
                )

                if expression_latex:
                    parts.append(expression_latex)

            if sym_result_node is not None:
                sym_result_latex = (
                    cls.parse_node_to_latex(
                        sym_result_node
                    )
                )

                if sym_result_latex:
                    parts.append(sym_result_latex)

        else:
            expression_node = next(
                (
                    child for child in node
                    if cls.strip_ns(child.tag)
                    != "result"
                ),
                None
            )

            if expression_node is not None:
                expression_latex = (
                    cls.parse_node_to_latex(
                        expression_node
                    )
                )

                if expression_latex:
                    parts.append(expression_latex)

        if result_node is not None:
            result_latex = (
                cls.parse_node_to_latex(
                    result_node
                )
            )

            if result_latex:
                parts.append(result_latex)

        return " = ".join(parts)

    @classmethod
    def process_file(cls, filepath):
        """Читает XMCD/XML и создаёт Markdown с формулами."""
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
            if cls.strip_ns(element.tag)
            == "region"
        ]

        def get_top(region):
            try:
                return float(
                    region.attrib.get("top", 0)
                )
            except (TypeError, ValueError):
                return 0.0

        regions.sort(key=get_top)

        content = []

        for region in regions:
            # Текстовые регионы.
            text_node = cls.find_first_by_tag(
                region,
                "text"
            )

            if text_node is not None:
                text_values = []

                for paragraph in text_node.iter():
                    if cls.strip_ns(
                        paragraph.tag
                    ) == "p":
                        paragraph_text = (
                            "".join(
                                paragraph.itertext()
                            ).strip()
                        )

                        if paragraph_text:
                            text_values.append(
                                paragraph_text
                            )

                if text_values:
                    content.append(
                        "\n".join(text_values)
                        + "\n\n"
                    )
                else:
                    text_value = (
                        "".join(
                            text_node.itertext()
                        ).strip()
                    )

                    if text_value:
                        content.append(
                            f"{text_value}\n\n"
                        )

            # Математические регионы.
            math_node = cls.find_first_by_tag(
                region,
                "math"
            )

            if math_node is None:
                continue

            for child in math_node:
                tag = cls.strip_ns(child.tag)
                latex = ""

                if tag == "define":
                    if len(child) >= 2:
                        left_side = (
                            cls.parse_node_to_latex(
                                child[0]
                            )
                        )

                        right_side = child[1]

                        if cls.strip_ns(
                            right_side.tag
                        ) == "eval":
                            right_latex = (
                                cls.parse_eval_to_latex(
                                    right_side
                                )
                            )
                        else:
                            right_latex = (
                                cls.parse_node_to_latex(
                                    right_side
                                )
                            )

                        if right_latex:
                            latex = (
                                f"{left_side} = "
                                f"{right_latex}"
                            )
                        else:
                            latex = left_side

                elif tag == "eval":
                    latex = cls.parse_eval_to_latex(
                        child
                    )

                else:
                    latex = cls.parse_node_to_latex(
                        child
                    )

                if latex:
                    content.append(
                        f"$$ {latex} $$\n\n"
                    )

        return "".join(content)


class WordConverter:
    """Конвертация Markdown/LaTeX в DOCX."""

    @staticmethod
    def check_pandoc(log_callback):
        """Проверяет наличие Pandoc."""
        try:
            version = pypandoc.get_pandoc_version()

            log_callback(
                f"[OK] Обнаружен Pandoc версии: {version}"
            )

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
        markdown_text,
        output_file,
        template_file,
        log_callback
    ):
        """Конвертирует Markdown с формулами в DOCX."""
        try:
            log_callback(
                f"[INFO] Начинается конвертация в "
                f"'{output_file}'..."
            )

            extra_args = []

            if template_file and os.path.exists(
                template_file
            ):
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

            log_callback(
                "[OK] Успех! Документ сохранен."
            )

            return True

        except Exception as exc:
            log_callback(
                f"[ERROR] Ошибка конвертации: {exc}"
            )

            return False


class ConverterApp:
    def __init__(self, root):
        self.root = root

        self.root.title(
            "Mathcad to Word Converter (со стилями)"
        )
        self.root.geometry("600x670")
        self.root.configure(
            padx=20,
            pady=20
        )

        self.input_file = None
        self.template_file = (
            self.find_default_template()
        )

        self.setup_ui()

        if DND_SUPPORTED:
            self.root.drop_target_register(
                DND_FILES
            )

            self.root.dnd_bind(
                "<<Drop>>",
                self.on_drop
            )

    def find_default_template(self):
        """Ищет DOCX-шаблон рядом со скриптом."""
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(
                sys.executable
            )
        else:
            base_dir = os.path.dirname(
                os.path.abspath(__file__)
            )

        template_dir = os.path.join(
            base_dir,
            "шаблон"
        )

        expected_template = os.path.join(
            template_dir,
            "Шаблон стилей.docx"
        )

        if os.path.exists(expected_template):
            return expected_template

        if os.path.isdir(template_dir):
            for filename in os.listdir(
                template_dir
            ):
                if filename.lower().endswith(".docx"):
                    return os.path.join(
                        template_dir,
                        filename
                    )

        fallback_dir = (
            r"C:\VS code здесь\mcdx to word\шаблон"
        )

        if os.path.isdir(fallback_dir):
            for filename in os.listdir(
                fallback_dir
            ):
                if filename.lower().endswith(".docx"):
                    return os.path.join(
                        fallback_dir,
                        filename
                    )

        return ""

    def setup_ui(self):
        # Исходный файл.
        frame_input = ttk.LabelFrame(
            self.root,
            text="Исходный файл Mathcad (.xmcd, .xml)",
            padding=10
        )

        frame_input.pack(
            fill=tk.X,
            pady=(0, 15)
        )

        self.lbl_input = ttk.Label(
            frame_input,
            text="Файл не выбран",
            foreground="gray"
        )

        self.lbl_input.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True
        )

        ttk.Button(
            frame_input,
            text="Выбрать файл",
            command=self.browse_input
        ).pack(
            side=tk.RIGHT,
            padx=5
        )

        if DND_SUPPORTED:
            ttk.Label(
                frame_input,
                text="(Или перетащите файл сюда)",
                font=("Segoe UI", 8, "italic")
            ).pack(
                side=tk.BOTTOM,
                pady=5
            )

        # Настройки округления.
        frame_settings = ttk.LabelFrame(
            self.root,
            text="Настройки округления (значащие цифры)",
            padding=10
        )

        frame_settings.pack(
            fill=tk.X,
            pady=(0, 15)
        )

        self.sig_figs_small_var = tk.IntVar(
            value=4
        )

        self.sig_figs_large_var = tk.IntVar(
            value=8
        )

        ttk.Label(
            frame_settings,
            text=(
                "Для малых чисел и дробей "
                "(по умолчанию 4):"
            )
        ).grid(
            row=0,
            column=0,
            sticky=tk.W,
            pady=2
        )

        ttk.Spinbox(
            frame_settings,
            from_=1,
            to=15,
            textvariable=self.sig_figs_small_var,
            width=5
        ).grid(
            row=0,
            column=1,
            sticky=tk.W,
            padx=10,
            pady=2
        )

        ttk.Label(
            frame_settings,
            text=(
                "Для больших чисел "
                "(по умолчанию 8):"
            )
        ).grid(
            row=1,
            column=0,
            sticky=tk.W,
            pady=2
        )

        ttk.Spinbox(
            frame_settings,
            from_=1,
            to=20,
            textvariable=self.sig_figs_large_var,
            width=5
        ).grid(
            row=1,
            column=1,
            sticky=tk.W,
            padx=10,
            pady=2
        )

        # Шаблон Word.
        frame_template = ttk.LabelFrame(
            self.root,
            text="Шаблон стилей Word (.docx)",
            padding=10
        )

        frame_template.pack(
            fill=tk.X,
            pady=(0, 15)
        )

        self.lbl_template = ttk.Label(
            frame_template,
            text=(
                self.template_file
                or "Шаблон не найден"
            ),
            foreground=(
                "black"
                if self.template_file
                else "red"
            )
        )

        self.lbl_template.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True
        )

        ttk.Button(
            frame_template,
            text="Изменить",
            command=self.browse_template
        ).pack(
            side=tk.RIGHT,
            padx=5
        )

        # Кнопка конвертации.
        self.btn_convert = ttk.Button(
            self.root,
            text="Конвертировать и сохранить как...",
            command=self.process_and_save,
            state=tk.DISABLED
        )

        self.btn_convert.pack(
            fill=tk.X,
            pady=10,
            ipady=5
        )

        # Окно лога.
        frame_log = ttk.LabelFrame(
            self.root,
            text="Статус и Логи",
            padding=5
        )

        frame_log.pack(
            fill=tk.BOTH,
            expand=True
        )

        self.log_text = tk.Text(
            frame_log,
            height=10,
            state=tk.DISABLED,
            bg="#f4f4f9",
            font=("Consolas", 9)
        )

        self.log_text.pack(
            fill=tk.BOTH,
            expand=True
        )

        self.log("Проверка окружения...")

        self.root.after(
            500,
            lambda: WordConverter.check_pandoc(
                self.log
            )
        )

    def log(self, message):
        """Добавляет сообщение в лог."""
        self.log_text.config(
            state=tk.NORMAL
        )

        self.log_text.insert(
            tk.END,
            message + "\n"
        )

        self.log_text.see(tk.END)

        self.log_text.config(
            state=tk.DISABLED
        )

        self.root.update()

    def browse_input(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл Mathcad",
            filetypes=[
                (
                    "Mathcad XML",
                    "*.xmcd *.xml"
                ),
                (
                    "All files",
                    "*.*"
                )
            ]
        )

        if filepath:
            self.set_input_file(filepath)

    def on_drop(self, event):
        """Обрабатывает перетаскивание файла."""
        filepath = event.data

        if (
            filepath.startswith("{")
            and filepath.endswith("}")
        ):
            filepath = filepath[1:-1]

        self.set_input_file(filepath)

    def set_input_file(self, filepath):
        if filepath.lower().endswith(
            (".xmcd", ".xml")
        ):
            self.input_file = filepath

            self.lbl_input.config(
                text=os.path.basename(filepath),
                foreground="black"
            )

            self.btn_convert.config(
                state=tk.NORMAL
            )

            self.log(
                f"[INFO] Выбран файл: {filepath}"
            )

        else:
            messagebox.showwarning(
                "Неверный формат",
                "Выберите файл .xmcd или .xml"
            )

    def browse_template(self):
        filepath = filedialog.askopenfilename(
            title="Выберите шаблон стилей Word",
            filetypes=[
                (
                    "Word Documents",
                    "*.docx"
                )
            ]
        )

        if filepath:
            self.template_file = filepath

            self.lbl_template.config(
                text=os.path.basename(filepath),
                foreground="black"
            )

            self.log(
                f"[INFO] Шаблон изменен на: {filepath}"
            )

    def process_and_save(self):
        if not self.input_file:
            return

        output_file = filedialog.asksaveasfilename(
            title="Сохранить результат как...",
            defaultextension=".docx",
            initialfile="Результат_Mathcad.docx",
            filetypes=[
                (
                    "Word Document",
                    "*.docx"
                )
            ]
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

                self.log(
                    "[ERROR] Файл заблокирован."
                )

                return

            except Exception as exc:
                messagebox.showerror(
                    "Ошибка",
                    f"Не удалось получить доступ к файлу:\n"
                    f"{exc}"
                )

                self.log(
                    f"[ERROR] Ошибка доступа: {exc}"
                )

                return

        self.log("\n--- Запуск обработки ---")

        try:
            MathcadParser.sig_figs_small = (
                self.sig_figs_small_var.get()
            )

            MathcadParser.sig_figs_large = (
                self.sig_figs_large_var.get()
            )

            self.log(
                "[INFO] Округление: "
                f"малые числа = "
                f"{MathcadParser.sig_figs_small}, "
                f"большие числа = "
                f"{MathcadParser.sig_figs_large}"
            )

        except Exception:
            self.log(
                "[WARN] Ошибка чтения настроек округления. "
                "Используются значения 4 и 8."
            )

            MathcadParser.sig_figs_small = 4
            MathcadParser.sig_figs_large = 8

        self.log(
            "[INFO] Парсинг XML файла Mathcad..."
        )

        try:
            markdown_content = (
                MathcadParser.process_file(
                    self.input_file
                )
            )

            if not markdown_content.strip():
                self.log(
                    "[WARN] Файл не содержит текста "
                    "или распознанных формул."
                )

        except Exception as exc:
            self.log(
                f"[ERROR] Ошибка парсинга: {exc}"
            )

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
        """Открывает файл приложением по умолчанию."""
        try:
            if sys.platform == "win32":
                os.startfile(filepath)

            elif sys.platform == "darwin":
                os.system(f"open '{filepath}'")

            else:
                os.system(f"xdg-open '{filepath}'")

        except Exception as exc:
            self.log(
                f"[ERROR] Не удалось открыть файл: {exc}"
            )


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