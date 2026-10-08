import os
import sys
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pypandoc

# Попытка импорта библиотеки для Drag-and-Drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_SUPPORTED = True
except ImportError:
    DND_SUPPORTED = False


class MathcadParser:
    """Класс для разбора XML/XMCD файлов Mathcad и конвертации их в Markdown (LaTeX)"""
    
    # Переменные для хранения настроек значащих цифр по умолчанию
    sig_figs_small = 4
    sig_figs_large = 8
    
    @staticmethod
    def strip_ns(tag):
        """Удаляет пространство имен (namespace) из XML тега"""
        if isinstance(tag, str) and '}' in tag:
            return tag.split('}')[1]
        return tag

    @classmethod
    def format_num(cls, num_str):
        """Форматирование чисел (интеллектуальное округление значащих цифр)"""
        if not num_str:
            return ""
        try:
            val = float(num_str)
            if val == 0:
                return "0"
                
            abs_val = abs(val)
            
            # Получаем первую значащую цифру через экспоненциальный формат
            sci_str = f"{abs_val:.15e}"
            first_digit = sci_str[0]
            
            # Базовое количество значащих цифр берем из настроек класса
            if abs_val >= 1:
                int_digits = len(str(int(abs_val)))
                base_sig_figs = max(cls.sig_figs_small, min(cls.sig_figs_large, int_digits))
            else:
                base_sig_figs = cls.sig_figs_small
                
            # Правило инженеров: если первая цифра 1 или 2, берем на 1 значащую цифру больше
            if first_digit in ['1', '2']:
                target_sig_figs = base_sig_figs + 1
            else:
                target_sig_figs = base_sig_figs
                
            # Форматируем с нужным количеством значащих цифр
            formatted = f"{val:.{target_sig_figs}g}"
            
            # Если число ушло в экспоненциальный формат (напр. 1.25e-05 или 4.43e+08)
            # Возвращаем его в классический вид без потери цифр
            if 'e' in formatted.lower():
                val_rounded = float(formatted)
                res = f"{val_rounded:.20f}".rstrip('0').rstrip('.')
                if not res: # защита от пустой строки
                    res = "0"
            else:
                res = formatted
                
            return res.replace('.', ',')
        except ValueError:
            # На случай, если переданная строка не является числом
            return num_str.replace('.', ',')

    @classmethod
    def find_first_by_tag(cls, node, tag):
        """Поиск первого дочернего элемента по тегу (с игнорированием namespace)"""
        for child in node.iter():
            if cls.strip_ns(child.tag) == tag:
                return child
        return None

    @classmethod
    def parse_node_to_latex(cls, node):
        """Рекурсивный разбор дерева математики Mathcad в строку LaTeX"""
        if node is None:
            return "?"
            
        tag = cls.strip_ns(node.tag)
        children = list(node)
        
        if tag == 'math':
            return " ".join([cls.parse_node_to_latex(c) for c in children])
            
        if tag in ('real', 'id', 'sym'):
            return cls.format_num(node.text.strip() if node.text else "")
            
        if tag == 'imag':
            symbol = node.attrib.get('symbol', 'j')
            return cls.format_num(node.text.strip() if node.text else "") + symbol
            
        if tag == 'complex':
            # Безопасный поиск без использования XPath (local-name не поддерживается в стандартном ET)
            r = next((c for c in children if cls.strip_ns(c.tag) == 'real'), None)
            i = next((c for c in children if cls.strip_ns(c.tag) == 'imag'), None)

            rt = cls.format_num(r.text.strip()) if r is not None and r.text else "0"
            it = cls.format_num(i.text.strip()) if i is not None and i.text else "0"
            sign = "" if it.startswith('-') else "+"
            return f"{rt}{sign}{it}j"
            
        if tag == 'parens':
            child_latex = cls.parse_node_to_latex(children[0]) if children else ""
            return f"\\left({child_latex}\\right)"
            
        if tag == 'apply':
            if not children:
                return ""
            op_node = children[0]
            op = cls.strip_ns(op_node.tag)
            args = [cls.parse_node_to_latex(c) for c in children[1:]]

            # Добавлена защита от выхода за пределы массива при незаконченных выражениях
            if op == 'mult':
                return f"{args[0]} \\cdot {args[1]}" if len(args) > 1 else (args[0] if args else "")
            if op == 'div':
                return f"\\frac{{{args[0]}}}{{{args[1]}}}" if len(args) > 1 else (f"\\frac{{{args[0]}}}{{?}}" if args else "")
            if op == 'plus':
                return f"{args[0]} + {args[1]}" if len(args) > 1 else (args[0] if args else "")
            if op == 'minus':
                return f"{args[0]} - {args[1]}" if len(args) > 1 else (f"-{args[0]}" if args else "-")
            if op == 'neg':
                return f"-{args[0]}" if args else "-"
            if op == 'pow':
                return f"{{{args[0]}}}^{{{args[1]}}}" if len(args) > 1 else (f"{{{args[0]}}}^{{?}}" if args else "")
            if op == 'sqrt':
                return f"\\sqrt{{{args[0]}}}" if args else "\\sqrt{?}"
            if op == 'absval':
                return f"\\left| {args[0]} \\right|" if args else "\\left| ? \\right|"
            if op == 'conjugate':
                return f"\\overline{{{args[0]}}}" if args else "\\overline{?}"
            if op == 'equal':
                return f"{args[0]} = {args[1]}" if len(args) > 1 else (f"{args[0]} = ?" if args else "=")
            if op in ('id', 'sym'):
                func_name = op_node.text.strip() if op_node.text else ""
                func_name = func_name.replace('%', '\\%')
                return f"\\mathrm{{{func_name}}}\\left({', '.join(args)}\\right)"
            return "?"
            
        if tag == 'function':
            bvars = next((c for c in children if cls.strip_ns(c.tag) == 'boundVars'), None)
            fn_node = next((c for c in children if cls.strip_ns(c.tag) in ('id', 'sym')), None)
            id_latex = cls.parse_node_to_latex(fn_node)
            vars_latex = ', '.join([cls.parse_node_to_latex(c) for c in list(bvars)]) if bvars is not None else ""
            if id_latex:
                return f"\\mathrm{{{id_latex}}}\\left({vars_latex}\\right)"
            return f"\\left({vars_latex}\\right)"
            
        if tag == 'sequence':
            return ', '.join([cls.parse_node_to_latex(c) for c in children])
            
        if tag == 'placeholder':
            return "\\square"
            
        # Fallback для неизвестных тегов-оберток (например, provenance, unitedValue)
        if children:
            for c in children:
                res = cls.parse_node_to_latex(c)
                if res and res != "?":
                    return res
                    
        return ""

    @classmethod
    def process_file(cls, filepath):
        """Читает файл .xmcd, извлекает текст и формулы, возвращает строку в формате Markdown/LaTeX"""
        try:
            tree = ET.parse(filepath)
            root = tree.getroot()
        except Exception as e:
            raise ValueError(f"Ошибка чтения XML/XMCD файла: {e}")

        # Поиск всех регионов безопасным способом, совместимым с любой версией Python
        regions = []
        for elem in root.iter():
            if cls.strip_ns(elem.tag) == 'region':
                regions.append(elem)
        
        # Сортировка по координате Y (сверху вниз)
        def get_top(r):
            try: return float(r.attrib.get('top', 0))
            except: return 0.0
        regions.sort(key=get_top)

        latex_content = ""

        for region in regions:
            # Обработка текста
            text_node = cls.find_first_by_tag(region, 'text')
            if text_node is not None:
                text_vals = []
                # Ищем параграфы
                for p in text_node.iter():
                    if cls.strip_ns(p.tag) == 'p':
                        p_text = "".join(p.itertext()).strip()
                        if p_text:
                            text_vals.append(p_text)
                
                # Если параграфы найдены, добавляем их с переносами строк
                if text_vals:
                    latex_content += "\n".join(text_vals) + "\n\n"
                else:
                    # Fallback для прямого текста без абзацев
                    text_val = "".join(text_node.itertext()).strip()
                    if text_val:
                        latex_content += f"{text_val}\n\n"

            # Обработка формул
            math_node = cls.find_first_by_tag(region, 'math')
            if math_node is not None:
                for child in list(math_node):
                    tag = cls.strip_ns(child.tag)
                    latex = ""

                    if tag == 'define':
                        if len(child) >= 2:
                            lhs_l = cls.parse_node_to_latex(child[0])
                            rhs_node = child[1]
                            
                            if cls.strip_ns(rhs_node.tag) == 'eval':
                                res_node = cls.find_first_by_tag(rhs_node, 'result')
                                res_val = res_node[0] if res_node is not None and len(res_node) > 0 else None
                                
                                # Проверка на наличие символьного вычисления (explicit, ALL)
                                sym_eval_node = cls.find_first_by_tag(rhs_node, 'symEval')
                                if sym_eval_node is not None:
                                    # Берем исходное выражение (пропуская служебные узлы command и symResult)
                                    expr_node = next((c for c in sym_eval_node if cls.strip_ns(c.tag) not in ('command', 'symResult')), None)
                                    # Берем подставленное выражение (с числами)
                                    sym_result_node = cls.find_first_by_tag(sym_eval_node, 'symResult')
                                    
                                    expr_latex = cls.parse_node_to_latex(expr_node) if expr_node is not None else ""
                                    sym_res_latex = cls.parse_node_to_latex(sym_result_node[0]) if (sym_result_node is not None and len(sym_result_node) > 0) else ""
                                    
                                    latex = f"{lhs_l} = {expr_latex}" if expr_latex else f"{lhs_l}"
                                    if sym_res_latex:
                                        latex += f" = {sym_res_latex}"
                                    if res_val is not None:
                                        latex += f" = {cls.parse_node_to_latex(res_val)}"
                                else:
                                    # Обычное вычисление (без explicit)
                                    expr_node = next((c for c in rhs_node if cls.strip_ns(c.tag) != 'result'), None)
                                    expr_latex = cls.parse_node_to_latex(expr_node) if expr_node is not None else ""
                                    
                                    latex = f"{lhs_l} = {expr_latex}" if expr_latex else f"{lhs_l}"
                                    if res_val is not None:
                                        latex += f" = {cls.parse_node_to_latex(res_val)}"
                            else:
                                latex = f"{lhs_l} = {cls.parse_node_to_latex(rhs_node)}"
                    
                    elif tag == 'eval':
                        if len(child) >= 1:
                            res_node = cls.find_first_by_tag(child, 'result')
                            res_val = res_node[0] if res_node is not None and len(res_node) > 0 else None
                            
                            # Проверка на наличие символьного вычисления (explicit, ALL)
                            sym_eval_node = cls.find_first_by_tag(child, 'symEval')
                            if sym_eval_node is not None:
                                expr_node = next((c for c in sym_eval_node if cls.strip_ns(c.tag) not in ('command', 'symResult')), None)
                                sym_result_node = cls.find_first_by_tag(sym_eval_node, 'symResult')
                                
                                expr_latex = cls.parse_node_to_latex(expr_node) if expr_node is not None else ""
                                sym_res_latex = cls.parse_node_to_latex(sym_result_node[0]) if (sym_result_node is not None and len(sym_result_node) > 0) else ""
                                
                                latex = expr_latex
                                if sym_res_latex:
                                    latex += f" = {sym_res_latex}" if latex else f"{sym_res_latex}"
                                if res_val is not None:
                                    latex += f" = {cls.parse_node_to_latex(res_val)}" if latex else f"{cls.parse_node_to_latex(res_val)}"
                            else:
                                # Обычное вычисление
                                expr_node = next((c for c in child if cls.strip_ns(c.tag) != 'result'), None)
                                expr_latex = cls.parse_node_to_latex(expr_node) if expr_node is not None else ""
                                
                                latex = expr_latex
                                if res_val is not None:
                                    latex += f" = {cls.parse_node_to_latex(res_val)}" if latex else f"{cls.parse_node_to_latex(res_val)}"
                    
                    else:
                        latex = cls.parse_node_to_latex(child)

                    if latex:
                        # Заворачиваем в display math для Pandoc (используем $$ для 100% надежного распознавания)
                        latex_content += f"$$ {latex} $$\n\n"

        return latex_content


class WordConverter:
    """Класс для конвертации сгенерированного Markdown/LaTeX в документ Word"""
    
    @staticmethod
    def check_pandoc(log_callback):
        """Проверяет наличие Pandoc, при необходимости скачивает"""
        try:
            version = pypandoc.get_pandoc_version()
            log_callback(f"[OK] Обнаружен Pandoc версии: {version}")
            return True
        except OSError:
            log_callback("[INFO] Pandoc не найден. Начинаю автоматическую загрузку (это займет время)...")
            try:
                pypandoc.download_pandoc()
                log_callback("[OK] Pandoc успешно скачан и установлен!")
                return True
            except Exception as e:
                log_callback(f"[ERROR] Не удалось скачать Pandoc: {e}")
                return False

    @staticmethod
    def convert_to_word(markdown_text, output_file, template_file, log_callback):
        """Выполняет конвертацию текста в docx"""
        try:
            log_callback(f"[INFO] Начинается конвертация в '{output_file}'...")
            
            extra_args = []
            if template_file and os.path.exists(template_file):
                # Добавляем шаблон стилей
                extra_args.append(f'--reference-doc={template_file}')
                log_callback(f"[INFO] Применяется шаблон: {os.path.basename(template_file)}")
            else:
                log_callback("[WARN] Шаблон стилей не задан или не найден. Используется стандартный стиль.")

            pypandoc.convert_text(
                source=markdown_text,
                to='docx',
                format='markdown',
                outputfile=output_file,
                extra_args=extra_args
            )
            log_callback("[OK] Успех! Документ сохранен.")
            return True
        except Exception as e:
            log_callback(f"[ERROR] Произошла ошибка при конвертации: {e}")
            return False


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mathcad to Word Converter (со стилями)")
        self.root.geometry("600x670") # Слегка увеличили высоту окна для новых настроек
        self.root.configure(padx=20, pady=20)
        
        self.input_file = None
        self.template_file = self.find_default_template()

        self.setup_ui()

        # Настройка Drag and Drop (если доступно)
        if DND_SUPPORTED:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_drop)

    def find_default_template(self):
        """Ищет шаблон в папке 'шаблон' рядом со скриптом программы"""
        # Определяем абсолютный путь к папке, где лежит сам скрипт .py
        if getattr(sys, 'frozen', False):
            # Если скрипт будет скомпилирован в .exe через pyinstaller
            base_dir = os.path.dirname(sys.executable)
        else:
            # Для обычного запуска скрипта
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        # Формируем путь к папке "шаблон"
        template_dir = os.path.join(base_dir, "шаблон")
        expected_template = os.path.join(template_dir, "Шаблон стилей.docx")
        
        # 1. Проверяем точное совпадение имени
        if os.path.exists(expected_template):
            return expected_template
            
        # 2. Если файл называется иначе, берем первый попавшийся .docx в папке "шаблон"
        if os.path.exists(template_dir) and os.path.isdir(template_dir):
            for file in os.listdir(template_dir):
                if file.lower().endswith('.docx'):
                    return os.path.join(template_dir, file)
                    
        # 3. Резервный абсолютный путь (fallback), если относительный поиск не сработал
        fallback_dir = r"C:\VS code здесь\mcdx to word\шаблон"
        if os.path.exists(fallback_dir) and os.path.isdir(fallback_dir):
            for file in os.listdir(fallback_dir):
                if file.lower().endswith('.docx'):
                    return os.path.join(fallback_dir, file)

        return ""

    def setup_ui(self):
        # 1. Зона выбора файла Mathcad
        frame_input = ttk.LabelFrame(self.root, text="Исходный файл Mathcad (.xmcd, .xml)", padding=10)
        frame_input.pack(fill=tk.X, pady=(0, 15))

        self.lbl_input = ttk.Label(frame_input, text="Файл не выбран", foreground="gray")
        self.lbl_input.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_browse = ttk.Button(frame_input, text="Выбрать файл", command=self.browse_input)
        btn_browse.pack(side=tk.RIGHT, padx=5)
        
        if DND_SUPPORTED:
            ttk.Label(frame_input, text="(Или перетащите файл сюда)", font=("Segoe UI", 8, "italic")).pack(side=tk.BOTTOM, pady=5)

        # 1.5. Настройки округления
        frame_settings = ttk.LabelFrame(self.root, text="Настройки округления (значащие цифры)", padding=10)
        frame_settings.pack(fill=tk.X, pady=(0, 15))
        
        self.sig_figs_small_var = tk.IntVar(value=4)
        self.sig_figs_large_var = tk.IntVar(value=8)

        # Окошко для малых чисел
        lbl_small = ttk.Label(frame_settings, text="Для малых чисел и дробей (по умолчанию 4):")
        lbl_small.grid(row=0, column=0, sticky=tk.W, pady=2)
        spin_small = ttk.Spinbox(frame_settings, from_=1, to=15, textvariable=self.sig_figs_small_var, width=5)
        spin_small.grid(row=0, column=1, sticky=tk.W, padx=10, pady=2)

        # Окошко для больших чисел
        lbl_large = ttk.Label(frame_settings, text="Для больших чисел (до запятой, по умолчанию 8):")
        lbl_large.grid(row=1, column=0, sticky=tk.W, pady=2)
        spin_large = ttk.Spinbox(frame_settings, from_=1, to=20, textvariable=self.sig_figs_large_var, width=5)
        spin_large.grid(row=1, column=1, sticky=tk.W, padx=10, pady=2)

        # 2. Зона выбора шаблона стилей
        frame_template = ttk.LabelFrame(self.root, text="Шаблон стилей Word (.docx)", padding=10)
        frame_template.pack(fill=tk.X, pady=(0, 15))

        self.lbl_template = ttk.Label(frame_template, text=self.template_file or "Шаблон не найден", 
                                      foreground="black" if self.template_file else "red")
        self.lbl_template.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_template = ttk.Button(frame_template, text="Изменить", command=self.browse_template)
        btn_template.pack(side=tk.RIGHT, padx=5)

        # 3. Кнопка запуска
        self.btn_convert = ttk.Button(self.root, text="Конвертировать и сохранить как...", 
                                      command=self.process_and_save, state=tk.DISABLED)
        self.btn_convert.pack(fill=tk.X, pady=10, ipady=5)

        # 4. Логи / Статус
        frame_log = ttk.LabelFrame(self.root, text="Статус и Логи", padding=5)
        frame_log.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(frame_log, height=10, state=tk.DISABLED, bg="#f4f4f9", font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Проверка Pandoc при старте
        self.log("Проверка окружения...")
        self.root.after(500, lambda: WordConverter.check_pandoc(self.log))

    def log(self, message):
        """Добавляет строку в текстовое поле логов"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()

    def browse_input(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл Mathcad",
            filetypes=[("Mathcad XML", "*.xmcd *.xml"), ("All files", "*.*")]
        )
        if filepath:
            self.set_input_file(filepath)

    def on_drop(self, event):
        """Обработчик перетаскивания файлов"""
        filepath = event.data
        # TkinterDnD может оборачивать пути в фигурные скобки, если есть пробелы
        if filepath.startswith('{') and filepath.endswith('}'):
            filepath = filepath[1:-1]
        self.set_input_file(filepath)

    def set_input_file(self, filepath):
        if filepath.lower().endswith(('.xmcd', '.xml')):
            self.input_file = filepath
            self.lbl_input.config(text=os.path.basename(filepath), foreground="black")
            self.btn_convert.config(state=tk.NORMAL)
            self.log(f"[INFO] Выбран файл: {filepath}")
        else:
            messagebox.showwarning("Неверный формат", "Пожалуйста, выберите файл .xmcd или .xml")

    def browse_template(self):
        filepath = filedialog.askopenfilename(
            title="Выберите шаблон стилей Word",
            filetypes=[("Word Documents", "*.docx")]
        )
        if filepath:
            self.template_file = filepath
            self.lbl_template.config(text=os.path.basename(filepath), foreground="black")
            self.log(f"[INFO] Шаблон изменен на: {filepath}")

    def process_and_save(self):
        if not self.input_file:
            return

        # Всплывающее окно сохранения файла
        output_file = filedialog.asksaveasfilename(
            title="Сохранить результат как...",
            defaultextension=".docx",
            initialfile="Результат_Mathcad.docx",
            filetypes=[("Word Document", "*.docx")]
        )

        if not output_file:
            return # Пользователь отменил сохранение

        if os.path.exists(output_file):
            try:
                # Пытаемся открыть файл на запись, чтобы проверить, не заблокирован ли он
                with open(output_file, 'a'):
                    pass
            except PermissionError:
                messagebox.showerror(
                    "Ошибка доступа", 
                    f"Не удалось перезаписать файл:\n{output_file}\n\nСкорее всего, этот документ сейчас открыт в Word. Пожалуйста, закройте его и повторите попытку сохранения."
                )
                self.log("[ERROR] Ошибка доступа: файл заблокирован (вероятно, открыт в Word).")
                return
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось получить доступ к файлу:\n{e}")
                self.log(f"[ERROR] Ошибка доступа к файлу: {e}")
                return

        self.log("\n--- Запуск обработки ---")
        
        # Обновляем параметры округления в парсере перед запуском конвертации
        try:
            MathcadParser.sig_figs_small = self.sig_figs_small_var.get()
            MathcadParser.sig_figs_large = self.sig_figs_large_var.get()
            self.log(f"[INFO] Округление: малые числа = {MathcadParser.sig_figs_small}, большие = {MathcadParser.sig_figs_large}")
        except Exception:
            self.log("[WARN] Ошибка чтения настроек округления. Используются значения 4 и 8.")
            MathcadParser.sig_figs_small = 4
            MathcadParser.sig_figs_large = 8
        
        # 1. Извлекаем LaTeX
        self.log("[INFO] Парсинг XML файла Mathcad...")
        try:
            markdown_content = MathcadParser.process_file(self.input_file)
            if not markdown_content.strip():
                self.log("[WARN] Файл не содержит текста или формул, или структура не распознана.")
        except Exception as e:
            self.log(f"[ERROR] Ошибка парсинга: {e}")
            return

        # 2. Конвертируем в Word
        success = WordConverter.convert_to_word(
            markdown_text=markdown_content,
            output_file=output_file,
            template_file=self.template_file,
            log_callback=self.log
        )

        if success:
            if messagebox.askyesno("Успех", f"Документ сохранен:\n{output_file}\n\nОткрыть файл сейчас?"):
                self.open_file(output_file)

    def open_file(self, filepath):
        """Открывает сгенерированный файл в ОС по умолчанию"""
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
    # Если установлена библиотека TkinterDnD2, используем её класс для поддержки перетаскивания.
    # Если нет - используем стандартный Tk.
    if DND_SUPPORTED:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        
    # Стилизация ttk
    style = ttk.Style(root)
    if sys.platform == "win32":
        style.theme_use("vista") # Современный вид на Windows
    
    app = ConverterApp(root)
    root.mainloop()