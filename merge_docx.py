#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["docxcompose", "pywin32; sys_platform == 'win32'"]
# ///
"""把資料夾裡的多份 docx 依檔名排序合併成一份，並讓頁碼與目錄頁次自動正確。

各份 docx 若各自寫死了「起始頁碼」（例如第 5 份設定從第 23 頁開始），
直接合併後只要任何一份增減頁數，後面的頁碼就全錯。這支腳本：

1. 用 docxcompose 合併內容（樣式、編號、圖片）
2. 保留每份檔案的分節，每份都從新的一頁開始
3. 拿掉所有寫死的起始頁碼，只留第一份的，頁碼改為一路連續
4. 目錄檔（檔名含「目錄」）裡手打的頁次數字，改成指向各份開頭書籤的 PAGEREF 欄位
5. 在 Windows 上用本機的 Word 重算目錄頁次並存檔；沒有 Word 時改設定開檔時更新欄位

用法：
    uv run merge_docx.py 報告資料夾                  # 輸出到 報告資料夾/../合併.docx
    uv run merge_docx.py 報告資料夾 -o out.docx
    uv run merge_docx.py 報告資料夾 --toc 02_目錄.docx  # 手動指定目錄檔
    uv run merge_docx.py 報告資料夾 --no-toc            # 不處理目錄
"""

import argparse
import re
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxcompose.composer import Composer

BOOKMARK_PREFIX = "_merge_doc_"  # 底線開頭是 Word 的隱藏書籤，不會出現在書籤清單


class NumberingSafeComposer(Composer):
    """docxcompose 2.2.0 的 add_numberings 對同一個元素只配一次新 numId，
    元素裡若有兩組編號（例如目錄表格同時有「壹、」與「（一）」），兩組會撞成同一個 ID。
    這裡改成一次只交給它一組編號，每組各自拿到新 ID，最後再統一改寫引用。"""

    def add_numberings(self, doc, element):
        refs = element.findall(".//" + qn("w:numId"))
        for num_id in dict.fromkeys(ref.get(qn("w:val")) for ref in refs):
            single = OxmlElement("w:p")
            ref = OxmlElement("w:numId")
            ref.set(qn("w:val"), num_id)
            single.append(ref)
            super().add_numberings(doc, single)
        for ref in refs:
            ref.val = self.num_id_mapping.get(ref.val, ref.val)


def body_sectpr(doc):
    return doc.element.body.find(qn("w:sectPr"))


def strip_page_starts(doc, keep_first):
    """移除所有寫死的起始頁碼；keep_first 時保留第一節的。"""
    for i, sectpr in enumerate(doc.element.body.iter(qn("w:sectPr"))):
        if keep_first and i == 0:
            continue
        pg = sectpr.find(qn("w:pgNumType"))
        if pg is not None:
            pg.attrib.pop(qn("w:start"), None)


def strip_header_footer_refs(sectpr):
    for tag in ("w:headerReference", "w:footerReference"):
        for ref in sectpr.findall(qn(tag)):
            sectpr.remove(ref)


def add_bookmark(doc, name):
    """在文件第一個段落開頭放書籤，供目錄的 PAGEREF 指向。"""
    p = doc.element.body.find(qn("w:p"))
    if p is None:
        p = OxmlElement("w:p")
        doc.element.body.insert(0, p)
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), "0")
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), "0")
    ppr = p.find(qn("w:pPr"))
    pos = 0 if ppr is None else list(p).index(ppr) + 1
    p.insert(pos, start)
    p.insert(pos + 1, end)


def paragraph_text(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def field_runs(rpr, instr, cached):
    """產生一個欄位：begin、指令、separate、快取顯示值、end。"""
    runs = []
    for kind in ("begin", "instr", "separate", "text", "end"):
        r = OxmlElement("w:r")
        if rpr is not None:
            r.append(deepcopy(rpr))
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = instr
        elif kind == "text":
            el = OxmlElement("w:t")
            el.text = cached
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        r.append(el)
        runs.append(r)
    return runs


def link_toc(toc_doc, bookmarks):
    """把目錄裡只含數字的段落，依出現順序換成 PAGEREF 欄位。"""
    number_paras = [
        p for p in toc_doc.element.body.iter(qn("w:p"))
        if re.fullmatch(r"\s*\d+\s*", paragraph_text(p))
    ]
    if len(number_paras) != len(bookmarks):
        print(f"警告：目錄有 {len(number_paras)} 個頁次數字，但要合併 {len(bookmarks)} 份檔案，"
              "數量不符，目錄頁次維持原樣", file=sys.stderr)
        return False
    for p, name in zip(number_paras, bookmarks):
        runs = p.findall(qn("w:r"))
        rpr = runs[0].find(qn("w:rPr")) if runs else None
        cached = paragraph_text(p).strip()
        for r in runs:
            p.remove(r)
        for r in field_runs(rpr, f" PAGEREF {name} \\h ", cached):
            p.append(r)
    return True


# settings.xml 的子元素有固定順序，updateFields 必須排在這些元素之前
SETTINGS_AFTER_UPDATE_FIELDS = (
    "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids",
    "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
    "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade",
    "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary",
    "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
)


def set_update_fields_on_open(doc):
    settings = doc.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        after = {qn(f"w:{name}") for name in SETTINGS_AFTER_UPDATE_FIELDS}
        anchor = next((child for child in settings if child.tag in after), None)
        if anchor is None:
            settings.append(el)
        else:
            anchor.addprevious(el)
    el.set(qn("w:val"), "true")


def merge(files, toc_file):
    bookmarks = [f"{BOOKMARK_PREFIX}{i:02d}" for i in range(len(files))]
    docs = []
    for i, f in enumerate(files):
        doc = Document(f)
        strip_page_starts(doc, keep_first=(i == 0))
        add_bookmark(doc, bookmarks[i])
        docs.append(doc)

    if toc_file is not None:
        link_toc(docs[files.index(toc_file)], bookmarks)

    master = docs[0]
    composer = NumberingSafeComposer(master)
    # 頁尾由第一份提供，後面各節自動沿用；不讓 docxcompose 搬動頁尾與頁碼設定
    composer.first_section_properties_added = True
    body = master.element.body

    for doc in docs[1:]:
        # 把目前最後一節的設定移進一個空段落，讓上一份檔案自成一節
        current = body_sectpr(master)
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        ppr.append(deepcopy(current))
        p.append(ppr)
        body.insert(list(body).index(current), p)
        # 最後一節改用這份檔案的設定（頁尾沿用前面的）
        new = deepcopy(body_sectpr(doc))
        strip_header_footer_refs(new)
        body.replace(current, new)
        composer.append(doc)

    return composer


def update_fields_with_word(path):
    """用本機的 Word 重算所有欄位（目錄頁次）並存檔；不是 Windows、沒裝 Word 或失敗時回傳 False。"""
    if sys.platform != "win32":
        return False
    try:
        import win32com.client
    except ImportError:
        return False
    word = None
    try:
        # DispatchEx 開一個獨立的 Word，不會動到使用者已開啟的文件
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone
        doc = word.Documents.Open(str(path.resolve()), AddToRecentFiles=False)
        doc.Repaginate()
        doc.Fields.Update()
        doc.Save()
        doc.Close()
        return True
    except Exception as e:
        print(f"用 Word 重算目錄頁次失敗：{e}", file=sys.stderr)
        return False
    finally:
        if word is not None:
            word.Quit()


def main():
    parser = argparse.ArgumentParser(description="合併多份 docx 並修正頁碼與目錄頁次")
    parser.add_argument("folder", type=Path, nargs="?", help="放 docx 的資料夾，依檔名排序合併")
    parser.add_argument("-o", "--output", type=Path, help="輸出檔（預設：資料夾旁的 合併.docx）")
    toc = parser.add_mutually_exclusive_group()
    toc.add_argument("--toc", help="目錄檔的檔名（預設：檔名含「目錄」的那份）")
    toc.add_argument("--no-toc", action="store_true", help="不改目錄頁次")
    args = parser.parse_args()
    if args.folder is None:
        # 雙擊執行檔時沒有參數；拖進主控台視窗的路徑可能帶引號
        answer = input("請把放 docx 的資料夾拖進這個視窗，再按 Enter：")
        args.folder = Path(answer.strip().strip('"'))

    # 排除 Word 開檔時產生的 ~$ 暫存鎖定檔
    files = sorted(f for f in args.folder.glob("*.docx") if not f.name.startswith("~$"))
    if len(files) < 2:
        sys.exit(f"{args.folder} 裡的 docx 少於 2 份，不需合併")

    toc_file = None
    if args.toc:
        toc_file = args.folder / args.toc
        if toc_file not in files:
            sys.exit(f"找不到目錄檔：{toc_file}")
    elif not args.no_toc:
        toc_file = next((f for f in files if "目錄" in f.name), None)

    output = args.output or args.folder.parent / "合併.docx"
    composer = merge(files, toc_file)
    composer.save(output)

    print(f"已合併 {len(files)} 份檔案 → {output}")
    print(f"目錄：{toc_file.name if toc_file else '未處理'}")
    if toc_file is None:
        return
    print("正在用 Word 重算目錄頁次⋯⋯")
    if update_fields_with_word(output):
        print("目錄頁次已由 Word 重算並存檔。")
    else:
        set_update_fields_on_open(composer.doc)
        composer.save(output)
        print("無法自動重算目錄頁次。請用 Word 開啟後按 Ctrl+A 全選，再按 F9 更新欄位。")


if __name__ == "__main__":
    if not getattr(sys, "frozen", False):
        main()
    else:
        # 打包成執行檔、雙擊執行時，視窗結束就會關掉；先停住讓使用者看結果
        try:
            main()
        except SystemExit as e:
            if isinstance(e.code, str):
                print(e.code)
        except Exception as e:
            print(f"發生錯誤：{e}")
        input("\n按 Enter 關閉視窗")
