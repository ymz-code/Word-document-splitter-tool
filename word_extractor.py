import os
import sys
import io
import zipfile
import threading
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import OrderedDict

if sys.platform == "win32":
    # PyInstaller 无控制台模式下 stdout/stderr 可能为 None，需做空值保护
    if sys.stdout is not None and hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if sys.stderr is not None and hasattr(sys.stderr, "buffer"):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from docx import Document

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("w", W_NS)


def extract_one(file_path, out_base, log=print):
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    if not os.access(str(file_path), os.R_OK):
        raise PermissionError(f"无法读取文件，权限不足")

    doc_name = file_path.stem
    doc = Document(str(file_path))

    out_dir = out_base / doc_name
    images_dir = out_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    txt_path = out_dir / f"{doc_name}.txt"

    # --- 1. 提取图片（使用 python-docx）---
    ns = {"w": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    rtype_embed = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

    image_count = 0
    used_image_ids = set()

    for para in doc.paragraphs:
        blips = para._element.findall(".//{%s}blip" % ns["w"])
        for blip in blips:
            rid = blip.get(rtype_embed)
            if rid and rid not in used_image_ids:
                used_image_ids.add(rid)
                image_count += 1
                _save_image(doc, rid, images_dir, image_count, log)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for img in cell._element.findall(".//{%s}blip" % ns["w"]):
                    rid = img.get(rtype_embed)
                    if rid and rid not in used_image_ids:
                        used_image_ids.add(rid)
                        image_count += 1
                        _save_image(doc, rid, images_dir, image_count, log)

    for rel in doc.part.rels.values():
        if "image" in rel.reltype and rel.rId not in used_image_ids:
            used_image_ids.add(rel.rId)
            image_count += 1
            _save_image(doc, rel.rId, images_dir, image_count, log)

    # --- 2. 提取文字（从原始 XML，覆盖文本框/形状等）---
    lines = _extract_paragraphs_from_xml(file_path)

    # --- 3. 补充表格（如果 python-docx 能识别到表格，以 | 分隔格式追加）---
    table_lines = []
    for table in doc.tables:
        for row in table.rows:
            row_texts = []
            for cell in row.cells:
                cell_texts = []
                for p in cell.paragraphs:
                    t = "".join(r.text for r in p.runs).strip()
                    if t:
                        cell_texts.append(t)
                if cell_texts:
                    row_texts.append(" ".join(cell_texts))
            if row_texts:
                table_lines.append(" | ".join(row_texts))

    # 如果 XML 提取结果很少但 python-docx 识别到了表格，补充表格行
    if len(lines) < 5 and table_lines:
        lines.extend(table_lines)

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    log(f"  📝 文字已保存: {txt_path}")
    log(f"  🖼  图片文件夹: {images_dir}")
    return len(lines), image_count


def _save_image(doc, rid, images_dir, index, log=print):
    try:
        rel = doc.part.rels[rid]
        image_data = rel.target_part.blob
        content_type = rel.target_part.content_type
        ext_map = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/gif": ".gif",
            "image/bmp": ".bmp",
            "image/tiff": ".tiff",
            "image/svg+xml": ".svg",
            "image/webp": ".webp",
        }
        ext = ext_map.get(content_type, ".png")
        fname = f"image_{index:03d}{ext}"
        fpath = images_dir / fname
        with open(fpath, "wb") as f:
            f.write(image_data)
    except Exception as e:
        log(f"  ⚠ 保存图片 rid={rid} 失败: {e}")


def _extract_paragraphs_from_xml(docx_path):
    """从 docx 原始 XML 中提取所有段落文本（包括文本框、形状里的文字）"""
    paragraphs = []

    with zipfile.ZipFile(str(docx_path), "r") as z:
        # 要解析的 XML 文件：正文 + 页眉 + 页脚
        xml_files = ["word/document.xml"]
        for name in z.namelist():
            if name.startswith("word/header") or name.startswith("word/footer"):
                xml_files.append(name)

        for xml_name in xml_files:
            if xml_name not in z.namelist():
                continue
            try:
                with z.open(xml_name) as f:
                    tree = ET.parse(f)
                root = tree.getroot()

                # 找所有 <w:p> 段落（无论嵌套多深）
                for p in root.iter(f"{{{W_NS}}}p"):
                    texts = []
                    for t in p.iter(f"{{{W_NS}}}t"):
                        if t.text:
                            texts.append(t.text)
                    line = "".join(texts).strip()
                    if line:
                        paragraphs.append(line)
            except Exception:
                continue

    # 去重：先合并连续相同的行，再全局去重（保持首次出现顺序）
    # 处理两栏布局、文本框复制等导致的重复内容
    deduped = []
    for line in paragraphs:
        if not deduped or deduped[-1] != line:
            deduped.append(line)

    # 全局去重：如果同一行出现多次，只保留第一次
    seen = set()
    final = []
    for line in deduped:
        if line not in seen:
            seen.add(line)
            final.append(line)

    return final


def run_cli():
    import argparse

    parser = argparse.ArgumentParser(description="Word文档拆分工具 - 提取文字和图片")
    parser.add_argument("inputs", nargs="+", help="Word文件或文件夹路径")
    parser.add_argument("-o", "--output", help="输出目录（默认: 第一个文件所在目录）")
    args = parser.parse_args()

    file_paths = []
    for p in args.inputs:
        p = os.path.abspath(p)
        if os.path.isdir(p):
            for f in os.listdir(p):
                if f.lower().endswith(".docx"):
                    file_paths.append(os.path.join(p, f))
        elif os.path.isfile(p) and p.lower().endswith(".docx"):
            file_paths.append(p)

    if not file_paths:
        print("❌ 未找到任何 .docx 文件")
        sys.exit(1)

    out_base = Path(args.output) if args.output else Path(file_paths[0]).parent
    out_base.mkdir(parents=True, exist_ok=True)

    print(f"📄 待处理 {len(file_paths)} 个文件")
    print(f"📁 输出目录: {out_base}\n")

    ok = 0
    fail = 0
    for idx, fp in enumerate(file_paths, 1):
        print(f"[{idx}/{len(file_paths)}] {os.path.basename(fp)}")
        try:
            _, img = extract_one(fp, out_base)
            print(f"  ✅ 完成 - {img} 张图片")
            ok += 1
        except PermissionError as e:
            print(f"  ❌ 无权限访问: {e}")
            print(f"     👉 请关闭Word后重试")
            fail += 1
        except Exception as e:
            print(f"  ❌ 失败: {e}")
            fail += 1

    print(f"\n🎉 成功 {ok} 个，失败 {fail} 个")


class WordExtractorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word文档拆分工具 - 文字与图片提取")
        self.root.geometry("620x520")
        self.root.resizable(False, False)

        self.file_paths = []
        self.output_dir = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        style = ttk.Style()
        style.configure("Title.TLabel", font=("Microsoft YaHei", 14, "bold"))
        style.configure("Info.TLabel", font=("Microsoft YaHei", 9), foreground="#666")

        title = ttk.Label(self.root, text="📄 Word文档拆分工具", style="Title.TLabel")
        title.pack(pady=(18, 4))

        subtitle = ttk.Label(
            self.root,
            text="自动提取文字（保存为txt）和图片（保存为文件夹）",
            style="Info.TLabel",
        )
        subtitle.pack(pady=(0, 14))

        file_frame = ttk.LabelFrame(self.root, text=" 选择Word文件 ", padding=10)
        file_frame.pack(fill="x", padx=20)

        btn_row = ttk.Frame(file_frame)
        btn_row.pack(fill="x")

        ttk.Button(btn_row, text="📁 选择文件", command=self._choose_files).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_row, text="🗂 选择文件夹", command=self._choose_folder).pack(
            side="left", padx=(0, 8)
        )
        ttk.Button(btn_row, text="❌ 清空列表", command=self._clear_files).pack(
            side="left"
        )

        self.file_listbox = tk.Listbox(
            file_frame, height=5, font=("Microsoft YaHei", 9)
        )
        self.file_listbox.pack(fill="x", pady=(8, 0))

        output_frame = ttk.LabelFrame(self.root, text=" 输出目录 ", padding=10)
        output_frame.pack(fill="x", padx=20, pady=(12, 0))

        out_row = ttk.Frame(output_frame)
        out_row.pack(fill="x")

        ttk.Entry(out_row, textvariable=self.output_dir).pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Button(out_row, text="浏览...", command=self._choose_output).pack(
            side="left"
        )

        run_frame = ttk.Frame(self.root)
        run_frame.pack(fill="x", padx=20, pady=14)

        self.run_btn = ttk.Button(
            run_frame, text="🚀 开始提取", command=self._start_extract
        )
        self.run_btn.pack(fill="x")

        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=20)

        log_frame = ttk.LabelFrame(self.root, text=" 运行日志 ", padding=8)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(10, 16))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=6, font=("Consolas", 9), state="disabled"
        )
        self.log_text.pack(fill="both", expand=True)

    def _log(self, msg):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update_idletasks()

    def _choose_files(self):
        paths = filedialog.askopenfilenames(
            title="选择Word文档",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")],
        )
        for p in paths:
            if p not in self.file_paths:
                self.file_paths.append(p)
                self.file_listbox.insert("end", p)
        if paths:
            self._set_default_output()

    def _choose_folder(self):
        folder = filedialog.askdirectory(title="选择包含Word文档的文件夹")
        if not folder:
            return
        added = 0
        for f in os.listdir(folder):
            if f.lower().endswith(".docx"):
                full = os.path.join(folder, f)
                if full not in self.file_paths:
                    self.file_paths.append(full)
                    self.file_listbox.insert("end", full)
                    added += 1
        self._log(f"从文件夹添加了 {added} 个 .docx 文件")
        self._set_default_output()

    def _clear_files(self):
        self.file_paths.clear()
        self.file_listbox.delete(0, "end")
        self.output_dir.set("")

    def _choose_output(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.output_dir.set(d)

    def _set_default_output(self):
        if not self.output_dir.get() and self.file_paths:
            first = self.file_paths[0]
            self.output_dir.set(str(Path(first).parent))

    def _start_extract(self):
        if not self.file_paths:
            messagebox.showwarning("提示", "请先选择Word文件！")
            return
        if not self.output_dir.get():
            messagebox.showwarning("提示", "请设置输出目录！")
            return

        self.run_btn.config(state="disabled")
        self.progress.config(value=0, maximum=len(self.file_paths))
        threading.Thread(target=self._extract_all, daemon=True).start()

    def _extract_all(self):
        total = len(self.file_paths)
        out_base = Path(self.output_dir.get())

        ok_count = 0
        fail_count = 0

        for idx, file_path in enumerate(self.file_paths, 1):
            self._log(f"\n[{idx}/{total}] 处理: {os.path.basename(file_path)}")
            try:
                _, img_count = extract_one(file_path, out_base, log=self._log)
                self._log(f"  ✅ 完成 - 图片: {img_count} 张")
                ok_count += 1
            except PermissionError as e:
                self._log(f"  ❌ 无权限访问: {e}")
                self._log(f"     👉 可能原因: 文件正被Word打开，或被杀毒软件锁定")
                self._log(f"     👉 请关闭Word后重试，或另存为新文件再试")
                fail_count += 1
            except Exception as e:
                self._log(f"  ❌ 失败: {e}")
                fail_count += 1

            self.progress.config(value=idx)

        self._log(f"\n🎉 完成！成功 {ok_count} 个，失败 {fail_count} 个")
        self.root.after(0, lambda: self.run_btn.config(state="normal"))
        self.root.after(
            0,
            lambda: messagebox.showinfo(
                "完成",
                f"成功: {ok_count} 个\n失败: {fail_count} 个\n\n输出目录:\n{out_base}",
            ),
        )

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_cli()
    else:
        WordExtractorGUI().run()
