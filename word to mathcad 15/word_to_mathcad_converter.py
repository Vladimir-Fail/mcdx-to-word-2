import os
import sys
import zipfile
import xml.etree.ElementTree as ET
import ast
import re
import html
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Попытка импорта библиотеки для Drag-and-Drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_SUPPORTED = True
except ImportError:
    DND_SUPPORTED = False


class WordMathParser:
    """Класс для разбора OMML (Office Math) из файлов .docx и их конвертации в AST Mathcad"""
    
    # Пространства имен Word
    NS = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'
    }

    @classmethod
    def extract_math_elements(cls, docx_path):
        """Извлекает все математические блоки из docx"""
        try:
            with zipfile.ZipFile(docx_path, 'r') as doc:
                # Читаем основной XML файл документа
                doc_xml = doc.read('word/document.xml')
            
            root = ET.fromstring(doc_xml)
            math_elements = root.findall('.//m:oMath', cls.NS)
            return math_elements
        except Exception as e:
            raise ValueError(f"Ошибка при чтении .docx: {e}")

    @classmethod
    def omml_to_py_string(cls, node):
        """Рекурсивно преобразует OMML-дерево в математическую строку для парсера AST"""
        tag = node.tag.split('}')[-1]
        
        # Контейнеры пропускаем и парсим детей
        if tag in ['oMath', 'e', 'num', 'den', 'sup', 'sub', 'oMathPara']:
            return "".join(cls.omml_to_py_string(c) for c in node)
            
        # Текстовый блок
        elif tag == 'r':
            t_node = node.find('m:t', cls.NS)
            if t_node is not None and t_node.text:
                txt = t_node.text
                # Очистка и замена символов для совместимости с Python AST
                txt = txt.replace('−', '-').replace('×', '*').replace('·', '*').replace('÷', '/')
                txt = txt.replace('=', '==') # Для правильного разбора уравнений
                txt = txt.replace(',', '.')  # Дробные числа
                # Удаляем пробелы, чтобы они не ломали переменные
                return txt.strip()
            return ""
            
        # Дробь
        elif tag == 'f':
            num = node.find('m:num', cls.NS)
            den = node.find('m:den', cls.NS)
            num_s = cls.omml_to_py_string(num) if num is not None else ""
            den_s = cls.omml_to_py_string(den) if den is not None else ""
            return f"(({num_s}) / ({den_s}))"
            
        # Верхний индекс (Степень)
        elif tag == 'sSup':
            base = node.find('m:e', cls.NS)
            sup = node.find('m:sup', cls.NS)
            base_s = cls.omml_to_py_string(base) if base is not None else ""
            sup_s = cls.omml_to_py_string(sup) if sup is not None else ""
            return f"(({base_s}) ** ({sup_s}))"
            
        # Нижний индекс (Подстрочный текст)
        elif tag == 'sSub':
            base = node.find('m:e', cls.NS)
            sub = node.find('m:sub', cls.NS)
            base_s = cls.omml_to_py_string(base) if base is not None else ""
            sub_s = cls.omml_to_py_string(sub) if sub is not None else ""
            # В Mathcad подстрочный символ обычно является частью имени переменной
            return f"{base_s.strip()}_{sub_s.strip()}"
            
        # Корень
        elif tag == 'rad':
            deg = node.find('m:deg', cls.NS)
            e = node.find('m:e', cls.NS)
            e_s = cls.omml_to_py_string(e) if e is not None else ""
            
            deg_content = cls.omml_to_py_string(deg) if deg is not None else ""
            if deg_content.strip():
                # Корень n-ой степени
                return f"(({e_s}) ** (1/({deg_content})))"
            else:
                # Квадратный корень
                return f"sqrt({e_s})"
                
        # Скобки (Делимитеры)
        elif tag == 'd':
            e = node.find('m:e', cls.NS)
            inner = cls.omml_to_py_string(e) if e is not None else ""
            return f"({inner})"
            
        # Неизвестный тег - идем вглубь
        else:
            return "".join(cls.omml_to_py_string(c) for c in node)

    @classmethod
    def sanitize_py_string(cls, expr):
        """Интеллектуальная расстановка знаков умножения (неявное умножение Word -> явное Mathcad)"""
        # Вставляем умножение между цифрой и буквой (вкл. кириллицу и греческий)
        expr = re.sub(r'(\d)([a-zA-Z\u0370-\u03ff\u0400-\u04ff])', r'\1*\2', expr)
        # Вставляем умножение между цифрой и открывающей скобкой: 2(x) -> 2*(x)
        expr = re.sub(r'(\d)\(', r'\1*(', expr)
        # Вставляем умножение после закрывающей скобки: )x -> )*x
        expr = re.sub(r'\)([a-zA-Z\u0370-\u03ff\u0400-\u04ff\d])', r')*\1', expr)
        return expr

    @classmethod
    def py_ast_to_mathcad_xml(cls, node, is_root=False):
        """Транслирует логическое дерево Python AST в синтаксис Mathcad XML (.xmcd)"""
        if node is None:
            return '<ml:id xml:space="preserve">?</ml:id>'
            
        # Поддержка Python 3.8+
        if hasattr(ast, 'Constant') and isinstance(node, ast.Constant):
            if isinstance(node.value, (int, float)):
                return f"<ml:real>{node.value}</ml:real>"
            return f'<ml:id xml:space="preserve">{node.value}</ml:id>'
            
        # Поддержка старых версий Python
        if hasattr(ast, 'Num') and isinstance(node, getattr(ast, 'Num', type(None))):
            return f"<ml:real>{node.n}</ml:real>"
        if hasattr(ast, 'Str') and isinstance(node, getattr(ast, 'Str', type(None))):
            return f'<ml:id xml:space="preserve">{node.s}</ml:id>'
            
        # Корень дерева
        if isinstance(node, ast.Module):
            if not node.body: return '<ml:id xml:space="preserve">?</ml:id>'
            return cls.py_ast_to_mathcad_xml(node.body[0], is_root=True)
            
        if isinstance(node, ast.Expr):
            return cls.py_ast_to_mathcad_xml(node.value, is_root=is_root)
            
        # Сравнения (Равенства / Уравнения)
        if isinstance(node, ast.Compare):
            left = cls.py_ast_to_mathcad_xml(node.left)
            # В уравнениях нас интересует первый компаратор
            right = cls.py_ast_to_mathcad_xml(node.comparators[0])
            if is_root:
                # Главное равенство становится присваиванием (для расчетов в Mathcad)
                return f'<ml:define xmlns:ml="http://schemas.mathsoft.com/math30">{left}{right}</ml:define>'
            else:
                # Внутреннее равенство - логическое равно
                return f"<ml:apply><ml:equal/>{left}{right}</ml:apply>"
                
        # Бинарные операции (+, -, *, /, ^)
        if isinstance(node, ast.BinOp):
            left = cls.py_ast_to_mathcad_xml(node.left)
            right = cls.py_ast_to_mathcad_xml(node.right)
            
            if isinstance(node.op, ast.Add): tag = "ml:plus"
            elif isinstance(node.op, ast.Sub): tag = "ml:minus"
            elif isinstance(node.op, ast.Mult): tag = "ml:mult"
            elif isinstance(node.op, ast.Div): tag = "ml:div"
            elif isinstance(node.op, ast.Pow): tag = "ml:pow"
            else: tag = "ml:unknown"
            
            return f"<ml:apply><{tag}/>{left}{right}</ml:apply>"
            
        # Унарные операции (-x)
        if isinstance(node, ast.UnaryOp):
            operand = cls.py_ast_to_mathcad_xml(node.operand)
            if isinstance(node.op, ast.USub):
                return f"<ml:apply><ml:neg/>{operand}</ml:apply>"
            return operand
            
        # Переменные (Идентификаторы)
        if isinstance(node, ast.Name):
            return f'<ml:id xml:space="preserve">{node.id}</ml:id>'
            
        # Вызов функций (Например, sin(x) или sqrt(x))
        if isinstance(node, ast.Call):
            if isinstance(node.func, ast.Name):
                func_name = node.func.id
                if func_name == "sqrt":
                    arg = cls.py_ast_to_mathcad_xml(node.args[0]) if node.args else '<ml:id xml:space="preserve">?</ml:id>'
                    return f"<ml:apply><ml:sqrt/>{arg}</ml:apply>"
                else:
                    args = "".join(cls.py_ast_to_mathcad_xml(a) for a in node.args)
                    return f'<ml:apply><ml:id xml:space="preserve">{func_name}</ml:id>{args}</ml:apply>'
            else:
                func = cls.py_ast_to_mathcad_xml(node.func)
                args = "".join(cls.py_ast_to_mathcad_xml(a) for a in node.args)
                return f"<ml:apply>{func}{args}</ml:apply>"
                
        # Fallback для неподдерживаемых конструкций (например, массивов)
        return '<ml:id xml:space="preserve">Unsupported</ml:id>'

class MathcadWriter:
    """Генератор структуры файла .xmcd (Mathcad 15)"""
    
    @staticmethod
    def generate_xmcd_file(math_regions, output_filepath):
        # Обновленный XML Заголовок файла Mathcad с более полными настройками стилей для обхода ошибки DTD
        xml_content = [
            '<?xml version="1.0" encoding="UTF-8" standalone="no"?>',
            '<worksheet version="3.0.3" xmlns="http://schemas.mathsoft.com/worksheet30" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
            'xmlns:ws="http://schemas.mathsoft.com/worksheet30" '
            'xmlns:ml="http://schemas.mathsoft.com/math30" '
            'xmlns:u="http://schemas.mathsoft.com/units10" '
            'xmlns:p="http://schemas.mathsoft.com/provenance10">',
            '  <settings>',
            '    <presentation>',
            '      <textRendering>',
            '        <textStyles>',
            '          <textStyle name="Normal">',
            '            <blockAttr margin-left="0" margin-right="0" text-indent="inherit" text-align="left" list-style-type="inherit" tabs="inherit"/>',
            '            <inlineAttr font-family="Arial" font-charset="0" font-size="10" font-weight="normal" font-style="normal" underline="false" line-through="false" vertical-align="baseline"/>',
            '          </textStyle>',
            '          <textStyle name="Math Text">', # Добавлен стиль Math Text, который часто требуется
            '            <blockAttr margin-left="0" margin-right="0" text-indent="inherit" text-align="left" list-style-type="inherit" tabs="inherit"/>',
            '            <inlineAttr font-family="Arial" font-charset="0" font-size="10" font-weight="normal" font-style="normal" underline="false" line-through="false" vertical-align="baseline"/>',
            '          </textStyle>',
            '        </textStyles>',
            '      </textRendering>',
            '      <mathRendering equation-color="#000">',
            '        <operators multiplication="narrow-dot" derivative="derivative" literal-subscript="large" definition="colon-equal" global-definition="triple-equal" local-definition="left-arrow" equality="bold-equal" symbolic-evaluation="right-arrow"/>',
            '        <mathStyles>',
            '          <mathStyle name="Variables" font-family="Times New Roman" font-charset="0" font-size="10" font-weight="normal" font-style="normal" underline="false"/>',
            '          <mathStyle name="Constants" font-family="Times New Roman" font-charset="0" font-size="10" font-weight="normal" font-style="normal" underline="false"/>',
            '          <mathStyle name="User 1" font-family="Arial" font-charset="0" font-size="10" font-weight="normal" font-style="normal" underline="false"/>',
            '          <mathStyle name="User 2" font-family="Courier New" font-charset="0" font-size="10" font-weight="normal" font-style="normal" underline="false"/>',
            '          <mathStyle name="User 3" font-family="Arial" font-charset="0" font-size="10" font-weight="bold" font-style="normal" underline="false"/>',
            '          <mathStyle name="User 4" font-family="Times New Roman" font-charset="0" font-size="10" font-weight="normal" font-style="italic" underline="false"/>',
            '          <mathStyle name="User 5" font-family="Times New Roman" font-charset="0" font-size="10" font-weight="normal" font-style="normal" underline="false"/>',
            '          <mathStyle name="User 6" font-family="Arial" font-charset="0" font-size="10" font-weight="normal" font-style="normal" underline="false"/>',
            '          <mathStyle name="User 7" font-family="Times New Roman" font-charset="0" font-size="10" font-weight="normal" font-style="normal" underline="false"/>',
            '          <mathStyle name="Math Text Font" font-family="Times New Roman" font-charset="0" font-size="14" font-weight="normal" font-style="normal" underline="false"/>',
            '        </mathStyles>',
            '        <dimensionNames mass="mass" length="length" time="time" current="current" thermodynamic-temperature="temperature" luminous-intensity="luminosity" amount-of-substance="substance" display="false"/>',
            '        <symbolics derivation-steps-style="vertical-insert" show-comments="false" evaluate-in-place="false"/>',
            '        <results numeric-only="true">',
            '          <general precision="3" show-trailing-zeros="false" radix="dec" complex-threshold="10" zero-threshold="15" imaginary-value="i" exponential-threshold="3"/>',
            '          <matrix display-style="auto" expand-nested-arrays="false"/>',
            '          <unit format-units="true" simplify-units="true" fractional-unit-exponent="false"/>',
            '        </results>',
            '      </mathRendering>',
            '      <pageModel show-page-frame="false" show-header-frame="false" show-footer-frame="false" header-footer-start-page="1" paper-code="1" orientation="portrait" print-single-page-width="false" page-width="612" page-height="792">',
            '        <margins left="86.4" right="86.4" top="86.4" bottom="86.4"/>',
            '        <header use-full-page-width="false"/>',
            '        <footer use-full-page-width="false"/>',
            '      </pageModel>',
            '      <colorModel background-color="#fff" default-highlight-color="#ffff80"/>',
            '      <language math="ru" UI="ru"/>',
            '    </presentation>',
            '    <calculation>',
            '      <builtInVariables array-origin="0" convergence-tolerance="0.001" constraint-tolerance="0.001" random-seed="1" prn-precision="4" prn-col-width="8"/>',
            '      <calculationBehavior automatic-recalculation="true" matrix-strict-singularity-check="false" optimize-expressions="false" exact-boolean="true" strings-use-origin="false" zero-over-zero="error">',
            '        <compatibility multiple-assignment="MC12" local-assignment="MC11"/>',
            '      </calculationBehavior>',
            '      <units>',
            '        <currentUnitSystem name="si" customized="false"/>',
            '      </units>',
            '    </calculation>',
            '    <editor view-annotations="false" view-regions="false">',
            '      <ruler is-visible="false" ruler-unit="in"/>',
            '      <grid granularity-x="6" granularity-y="6"/>',
            '    </editor>',
            '    <fileFormat image-type="image/png" image-quality="75" save-numeric-results="true" exclude-large-results="true" save-text-images="false"/>',
            '    <miscellaneous>',
            '      <item id="0" name="creator" value="Python Word to Mathcad Converter"/>',
            '    </miscellaneous>',
            '  </settings>',
            '  <regions>'
        ]
        
        # Добавляем регионы сверху вниз с отступом 30px
        current_top = 15
        
        for i, region in enumerate(math_regions):
            xml_content.append(f'    <region region-id="{i+1}" left="15" top="{current_top}">')
            
            if "error_text" in region:
                # Если формулу не удалось разобрать логически, выводим её как текст (комментарий), чтобы не терять
                escaped_text = html.escape(region["error_text"])
                xml_content.append('      <text>')
                xml_content.append(f'        <p>Нераспознанная формула: {escaped_text}</p>')
                xml_content.append('      </text>')
            else:
                # Успешно разобранная формула
                xml_content.append('      <math optimize="false" disable-calc="false">')
                xml_content.append(f'        {region["xml"]}')
                xml_content.append('      </math>')
                
            xml_content.append('    </region>')
            current_top += 35 # Сдвигаем следующую формулу ниже
            
        xml_content.append('  </regions>')
        xml_content.append('</worksheet>')
        
        # Сохранение файла
        with open(output_filepath, 'w', encoding='utf-8') as f:
            f.write("\n".join(xml_content))

class ReverseConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word to Mathcad Converter")
        self.root.geometry("600x500")
        self.root.configure(padx=20, pady=20)
        
        self.input_file = None
        self.setup_ui()

        # Настройка Drag and Drop (если доступно)
        if DND_SUPPORTED:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_drop)

    def setup_ui(self):
        # 1. Зона выбора файла Word
        frame_input = ttk.LabelFrame(self.root, text="Исходный документ Word (.docx)", padding=10)
        frame_input.pack(fill=tk.X, pady=(0, 15))

        self.lbl_input = ttk.Label(frame_input, text="Файл не выбран", foreground="gray")
        self.lbl_input.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_browse = ttk.Button(frame_input, text="Выбрать файл", command=self.browse_input)
        btn_browse.pack(side=tk.RIGHT, padx=5)
        
        if DND_SUPPORTED:
            ttk.Label(frame_input, text="(Или перетащите файл сюда)", font=("Segoe UI", 8, "italic")).pack(side=tk.BOTTOM, pady=5)

        # 2. Инфопанель
        info_text = "Программа извлекает формулы из Word (OMML) и переводит их\n" \
                    "в расчетные блоки Mathcad. Обычный текст будет проигнорирован.\n" \
                    "Сложные визуальные структуры (матрицы, интегралы) могут быть перенесены как текст."
        ttk.Label(self.root, text=info_text, justify=tk.CENTER, foreground="#555").pack(pady=10)

        # 3. Кнопка запуска
        self.btn_convert = ttk.Button(self.root, text="Сгенерировать Mathcad (.xmcd) ...", 
                                      command=self.process_and_save, state=tk.DISABLED)
        self.btn_convert.pack(fill=tk.X, pady=10, ipady=5)

        # 4. Логи / Статус
        frame_log = ttk.LabelFrame(self.root, text="Статус и Логи", padding=5)
        frame_log.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(frame_log, height=10, state=tk.DISABLED, bg="#f4f4f9", font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        self.log("Готов к работе. Выберите файл .docx с формулами.")

    def log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()

    def browse_input(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл Word",
            filetypes=[("Word Document", "*.docx"), ("All files", "*.*")]
        )
        if filepath:
            self.set_input_file(filepath)

    def on_drop(self, event):
        filepath = event.data
        if filepath.startswith('{') and filepath.endswith('}'):
            filepath = filepath[1:-1]
        self.set_input_file(filepath)

    def set_input_file(self, filepath):
        if filepath.lower().endswith('.docx'):
            self.input_file = filepath
            self.lbl_input.config(text=os.path.basename(filepath), foreground="black")
            self.btn_convert.config(state=tk.NORMAL)
            self.log(f"[INFO] Выбран файл: {filepath}")
        else:
            messagebox.showwarning("Неверный формат", "Пожалуйста, выберите файл .docx")

    def process_and_save(self):
        if not self.input_file:
            return

        output_file = filedialog.asksaveasfilename(
            title="Сохранить Mathcad файл как...",
            defaultextension=".xmcd",
            initialfile="Расчетный_файл_Mathcad.xmcd",
            filetypes=[("Mathcad XML", "*.xmcd")]
        )

        if not output_file:
            return

        self.log("\n--- Запуск обработки ---")
        
        try:
            self.log("[INFO] Поиск формул в документе Word...")
            math_elements = WordMathParser.extract_math_elements(self.input_file)
            
            if not math_elements:
                self.log("[WARN] В документе не найдены формулы Office Math.")
                messagebox.showinfo("Пусто", "Формулы не найдены.")
                return
                
            self.log(f"[INFO] Найдено формул: {len(math_elements)}")
            
            parsed_regions = []
            success_count = 0
            
            for i, math_el in enumerate(math_elements):
                # 1. OMML в строку
                raw_str = WordMathParser.omml_to_py_string(math_el)
                if not raw_str:
                    continue
                    
                # 2. Подготовка строки для AST
                sanitized_str = WordMathParser.sanitize_py_string(raw_str)
                
                try:
                    # 3. Парсинг строки в логическое дерево с помощью Python AST
                    tree = ast.parse(sanitized_str)
                    
                    # 4. Генерация синтаксиса Mathcad XML
                    mathcad_xml = WordMathParser.py_ast_to_mathcad_xml(tree)
                    
                    parsed_regions.append({"xml": mathcad_xml})
                    success_count += 1
                except SyntaxError:
                    # Если формула слишком сложная (визуальная) и не имеет расчетной логики
                    self.log(f"[WARN] Формула '{raw_str}' сохранена как текст (сложная структура).")
                    parsed_regions.append({"error_text": raw_str})
                    
            self.log(f"[INFO] Успешно сконвертировано в расчетный формат: {success_count} из {len(math_elements)}.")
            
            # Сохранение .xmcd
            MathcadWriter.generate_xmcd_file(parsed_regions, output_file)
            self.log("[OK] Файл успешно сохранен!")
            
            if messagebox.askyesno("Успех", f"Документ сохранен:\n{output_file}\n\nОткрыть файл сейчас?"):
                self.open_file(output_file)
                
        except Exception as e:
            self.log(f"[ERROR] Ошибка в процессе конвертации: {e}")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")

    def open_file(self, filepath):
        try:
            if sys.platform == "win32":
                os.startfile(filepath)
            elif sys.platform == "darwin": # macOS
                os.system(f"open '{filepath}'")
            else: # Linux
                os.system(f"xdg-open '{filepath}'")
        except Exception as e:
            self.log(f"[ERROR] Не удалось открыть файл: {e}")


if __name__ == "__main__":
    if DND_SUPPORTED:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        
    style = ttk.Style(root)
    if sys.platform == "win32":
        try:
            style.theme_use("vista") 
        except: pass
    
    app = ReverseConverterApp(root)
    root.mainloop()