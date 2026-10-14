#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ExifTool 繁體中文(zh_tw)語言檔製作工具（單一檔案）

流程：
  0. listx   : 執行 exiftool -listx -all:all > out.xml（可用 -g exif 改成只匯出 EXIF），直接產生 out.xml（不經 shell，無編碼問題）
  1. prepare : out.xml -> out_prepared.xml     只「新增」缺少的 <desc lang='zh_tw'>，其餘一字不動
  2. check   : out.xml vs out_prepared.xml     驗證修改版只多出 zh_tw 行
  3. export  : out_prepared.xml -> to_translate.txt   (name<TAB>英文<TAB>中文)
  4. import  : to_translate.txt -> out_prepared.xml   把翻譯寫回（就地更新，並自動 check）
  5. build   : out_prepared.xml -> zh_tw.pm

用法：
  python exiftool_zh_tw_tool.py                 (不帶參數 = 互動選單)
  python exiftool_zh_tw_tool.py listx [-g exif] [--exiftool 路徑] [-f]
  python exiftool_zh_tw_tool.py prepare
  python exiftool_zh_tw_tool.py check
  python exiftool_zh_tw_tool.py export [--all]
  python exiftool_zh_tw_tool.py import [-t a.txt b.txt ...]
  python exiftool_zh_tw_tool.py convert -t 舊版翻譯檔.txt   (舊版 id<TAB>文字 → 新格式)
  python exiftool_zh_tw_tool.py build
"""

import argparse
import html
import os
import re
import shutil
import subprocess
import sys
from collections import defaultdict
from xml.sax.saxutils import escape as xml_escape

DEF_SRC = "out.xml"
DEF_PREP = "out_prepared.xml"
DEF_TXT = "to_translate.txt"
DEF_VTXT = "to_translate_values.txt"
DEF_PM = "zh_tw.pm"
DEF_OVR = "name_overrides.txt"
DEF_BASE = "zh_tw_official.pm"
DEF_CONF = "conflicts_report.txt"

RE_TAG_START = re.compile(r"^\s*<tag\s")
RE_TAG_END = re.compile(r"^\s*</tag>")
RE_NAME = re.compile(r"""\bname=(['"])(.*?)\1""")
RE_ID = re.compile(r"""\sid=(['"])(.*?)\1""")
RE_VALUES = re.compile(r"^\s*<(values|key|val)\b")
RE_DESC = re.compile(
    r"""^(\s*)<desc\s+lang=(['"])([^'"]+)\2\s*>(.*?)</desc>\s*$"""
)
RE_DESC_OPEN = re.compile(r"^\s*<desc\b")
RE_VAL = re.compile(r"""^\s*<val\s+lang=(['"])([^'"]+)\1\s*>(.*?)</val>\s*$""")
RE_KEY_START = re.compile(r"^\s*<key\b")
RE_KEY_END = re.compile(r"^\s*</key>")
RE_TABLE = re.compile(r"^\s*<table\s")
RE_ZH_LINE = re.compile(
    r"""^\s*<(?:desc|val)\s+lang=(['"])zh_tw\1\s*>.*</(?:desc|val)>\s*$"""
)
RE_VAL_ZH_PARTS = re.compile(
    r"""^(\s*<val\s+lang=(['"])zh_tw\2\s*>)(.*?)(</val>\s*)$"""
)
RE_ZH_PARTS = re.compile(
    r"""^(\s*<desc\s+lang=(['"])zh_tw\2\s*>)(.*?)(</desc>\s*)$"""
)


# ----------------------------------------------------------------- 檔案 I/O
def read_text(path):
  with open(path, "rb") as f:
    data = f.read()
  if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
    raise ValueError(
        f"{path} 是 UTF-16 編碼（PowerShell 5.1 的 > 常見結果）。"
        "請改在 cmd 執行 exiftool -listx ... > out.xml，或轉成 UTF-8。"
    )
  try:
    return data.decode("utf-8")
  except UnicodeDecodeError as e:
    raise ValueError(f"{path} 不是有效的 UTF-8：{e}")


def write_text(path, text):
  with open(path, "wb") as f:
    f.write(text.encode("utf-8"))


def split_lines(text):
  # 以 \n 切割、\n 合併可完全還原（CRLF 的 \r 會留在行尾）
  return text.split("\n")


def un(s):
  return html.unescape(s)


# ----------------------------------------------------------------- 解析
def parse_tags(lines):
  """逐行掃描 <tag> ... </tag>，回傳每個標籤的資訊（文字皆為原始跳脫狀態）。"""
  tags = []
  n = len(lines)
  i = 0
  cur_table = ""
  while i < n:
    if RE_TABLE.match(lines[i]):
      m0 = RE_NAME.search(lines[i])
      cur_table = un(m0.group(2)) if m0 else ""
      i += 1
      continue
    if not RE_TAG_START.match(lines[i]):
      i += 1
      continue
    start = i
    end = start + 1
    while end < n and not RE_TAG_END.match(lines[end]):
      end += 1
    end = min(end, n - 1)

    m = RE_NAME.search(lines[start])
    mi = RE_ID.search(lines[start])
    t = {
        "name": un(m.group(2)) if m else "",
        "table": cur_table,
        "id": un(mi.group(2)) if mi else "",
        "start": start,
        "end": end,
        "en": None,
        "first": None,
        "zh_idx": None,
        "zh_raw": None,
        "last_desc": None,
        "indent": None,
        "quote": "'",
        "eol": "",
        "bad": False,
        "vals": [],
    }
    vals_from = None
    for k in range(start + 1, end):
      line = lines[k]
      if RE_VALUES.match(line):
        vals_from = k
        break
      dm = RE_DESC.match(line)
      if not dm:
        if RE_DESC_OPEN.match(line):
          t["bad"] = True  # 跨行的 <desc>，不處理
        continue
      indent, quote, lang, raw = dm.groups()
      t["last_desc"] = k
      t["indent"] = indent
      t["quote"] = quote
      t["eol"] = "\r" if line.endswith("\r") else ""
      if lang == "en" and t["en"] is None and raw:
        t["en"] = raw
      if lang == "zh_tw" and t["zh_idx"] is None:
        t["zh_idx"] = k
        t["zh_raw"] = raw
      if raw and t["first"] is None:
        t["first"] = raw
    # 值翻譯：以 <key> 為單位記錄 en / zh_tw 的 <val>
    keys = []
    if vals_from is not None:
      cur = None
      for k in range(vals_from, end):
        line = lines[k]
        if RE_KEY_START.match(line):
          cur = {
              "en": None, "zh": None, "en_idx": None, "zh_idx": None,
              "last": None, "indent": "", "quote": "'", "eol": "",
          }
          continue
        if cur is None:
          continue
        vm = RE_VAL.match(line)
        if vm:
          cur["last"] = k
          cur["indent"] = re.match(r"\s*", line).group(0)
          cur["quote"] = vm.group(1)
          cur["eol"] = "\r" if line.endswith("\r") else ""
          if vm.group(2) == "en" and cur["en"] is None:
            cur["en"], cur["en_idx"] = vm.group(3), k
          elif vm.group(2) == "zh_tw" and cur["zh"] is None:
            cur["zh"], cur["zh_idx"] = vm.group(3), k
        elif RE_KEY_END.match(line):
          if cur["en"] is not None:
            keys.append(cur)
          cur = None
    t["keys"] = keys
    t["vals"] = [(kk["en"], kk["zh"]) for kk in keys if kk["zh"]]
    # 比對基準：en → 第一個有文字的 desc → 標籤名稱
    t["base"] = t["en"] or t["first"] or xml_escape(t["name"])
    tags.append(t)
    i = end + 1
  return tags


def one_line(s):
  return s.replace("\t", " ").replace("\r", " ").replace("\n", " ")


# ----------------------------------------------------------------- 0. 匯出 out.xml
def find_exiftool(user_path=None):
  """依序尋找：指定路徑 → PATH → 目前資料夾 / 本程式資料夾。"""
  if user_path:
    return user_path if os.path.exists(user_path) else shutil.which(user_path)
  found = shutil.which("exiftool")
  if found:
    return found
  here = os.path.dirname(os.path.abspath(__file__))
  for d in (os.getcwd(), here):
    for name in ("exiftool.exe", "exiftool(-k).exe", "exiftool"):
      cand = os.path.join(d, name)
      if os.path.isfile(cand):
        return cand
  return None


def cmd_listx(dst=DEF_SRC, group="-all:all", exiftool=None, force=False):
  """等同 exiftool -listx [群組] > out.xml（直接寫入原始位元組，保持 UTF-8）。
  group：預設 -all:all；可輸入群組名（exif → -exif:all、all → -all:all）、
  完整參數（如 -exif:all），或 none＝完全不加群組參數。"""
  if os.path.exists(dst) and not force:
    raise ValueError(f"{dst} 已存在。若要覆蓋請加 --force（選單中會詢問）。")
  exe = find_exiftool(exiftool)
  if not exe:
    raise ValueError(
        "找不到 exiftool。請安裝並加入 PATH，或用 --exiftool 指定路徑，"
        "或把 exiftool.exe 放在目前資料夾。"
    )
  grp = group.strip()
  if grp.lower() in ("", "none", "不限制"):
    gargs = []
  elif grp.startswith("-"):
    gargs = grp.split()
  else:
    gargs = [f"-{grp}:all"]  # 例如輸入 exif → -exif:all
  cmd = [exe, "-listx"] + gargs
  r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
  if r.returncode != 0 or not r.stdout.strip():
    err = r.stderr.decode("utf-8", "replace").strip()
    raise ValueError(f"exiftool 執行失敗（代碼 {r.returncode}）：{err}")
  if b"<taginfo" not in r.stdout[:2000]:
    raise ValueError("exiftool 的輸出看起來不是 -listx 的 XML，已中止。")
  if not r.stdout.rstrip().endswith(b"</taginfo>"):
    raise ValueError("輸出沒有以 </taginfo> 結尾，內容可能被截斷，已中止。")
  with open(dst, "wb") as f:
    f.write(r.stdout)
  try:
    ver = subprocess.run([exe, "-ver"], stdout=subprocess.PIPE).stdout.decode().strip()
  except Exception:
    ver = "?"
  n = len(parse_tags(split_lines(r.stdout.decode("utf-8", "replace"))))
  tables = r.stdout.count(b"<table ")
  print(f"ExifTool {ver}：{' '.join(cmd[1:])}")
  print(f"已輸出 {dst}（{len(r.stdout) / 1024:.0f} KB，表格 {tables} 個，標籤 {n} 個）")
  if not gargs:
    print("範圍：未加群組參數")


# ----------------------------------------------------------------- 1. prepare
def cmd_prepare(src=DEF_SRC, dst=DEF_PREP, values=True):
  lines = split_lines(read_text(src))
  tags = parse_tags(lines)
  ins = {}
  have = bad = 0
  for t in tags:
    if t["bad"]:
      bad += 1
      continue
    if t["zh_idx"] is not None:
      have += 1
      continue
    if t["last_desc"] is not None:
      pos, indent, q, eol = t["last_desc"], t["indent"], t["quote"], t["eol"]
    else:
      sl = lines[t["start"]]
      indent = re.match(r"\s*", sl).group(0) + " "
      pos, q, eol = t["start"], "'", ("\r" if sl.endswith("\r") else "")
    ins.setdefault(pos, []).append(
        f"{indent}<desc lang={q}zh_tw{q}>{t['base']}</desc>{eol}"
    )
  dadd = sum(len(v) for v in ins.values())

  vadd = vhave = 0
  if values:
    for t in tags:
      if t["bad"]:
        continue
      for kk in t["keys"]:
        if kk["zh"] is not None:
          vhave += 1
        else:
          q = kk["quote"]
          ins.setdefault(kk["last"], []).append(
              f"{kk['indent']}<val lang={q}zh_tw{q}>{kk['en']}</val>{kk['eol']}"
          )
          vadd += 1

  out = []
  for i, line in enumerate(lines):
    out.append(line)
    if i in ins:
      out.extend(ins[i])
  write_text(dst, "\n".join(out))

  print(f"標籤總數：{len(tags)}")
  print(f"  已有 zh_tw 而略過：{have}")
  print(f"  新增 zh_tw 說明：{dadd}（暫以英文原文填入）")
  if values:
    print(f"  值：已有 zh_tw 而略過 {vhave}，新增 zh_tw 值 {vadd}（暫以英文原文填入）")
  if bad:
    print(f"  跨行 <desc> 無法處理而略過：{bad}")
  print(f"已輸出 {dst}")
  return cmd_check(src, dst)


# ----------------------------------------------------------------- 2. check
def cmd_check(orig=DEF_SRC, mod=DEF_PREP):
  a = split_lines(read_text(orig))
  b = split_lines(read_text(mod))
  i = j = added = modified = 0
  while j < len(b):
    if i < len(a) and a[i] == b[j]:
      i += 1
      j += 1
    elif i < len(a) and RE_ZH_LINE.match(a[i]) and RE_ZH_LINE.match(b[j]):
      modified += 1  # 同位置的既有 zh_tw 被改寫
      i += 1
      j += 1
    elif RE_ZH_LINE.match(b[j]):
      added += 1
      j += 1
    else:
      print(f"[比對失敗] 原檔第 {i + 1} 行 與 修改版第 {j + 1} 行不一致：")
      print("  原檔  :", a[i][:120] if i < len(a) else "(已結束)")
      print("  修改版:", b[j][:120])
      return False
  if i != len(a):
    print(f"[比對失敗] 修改版缺少原檔第 {i + 1} 行起的內容")
    return False
  print(f"[比對通過] 修改版相對原檔只多出 {added} 行 zh_tw，其餘完全相同")
  if modified:
    print(f"  另有 {modified} 行原檔既有的 zh_tw 內文被改寫（其餘標籤、符號皆未動）")
  return True


# ----------------------------------------------------------------- 3. export
def cmd_export(src=DEF_PREP, dst=DEF_TXT, export_all=False, vdst=DEF_VTXT, values=True):
  tags = parse_tags(split_lines(read_text(src)))
  rows = {}
  for t in tags:
    if t["bad"] or not t["name"]:
      continue
    en = un(t["base"])
    zh = un(t["zh_raw"]) if t["zh_raw"] else ""
    key = (t["name"], en)
    if key in rows:
      continue
    if export_all or zh == "" or zh == en:
      rows[key] = zh or en
  with open(dst, "wb") as f:
    f.write(
        "\n".join(
            f"{n}\t{one_line(e)}\t{one_line(z)}" for (n, e), z in rows.items()
        ).encode("utf-8")
    )
  print(f"已匯出 {len(rows)} 筆到 {dst}（{'全部' if export_all else '僅未翻譯'}）")
  print("格式：標籤名稱<TAB>英文<TAB>中文。請只修改第三欄，勿動前兩欄與 Tab。")
  if not values:
    return
  vrows = {}
  for t in tags:
    if t["bad"]:
      continue
    for kk in t["keys"]:
      en = un(kk["en"])
      zh = un(kk["zh"]) if kk["zh"] else ""
      if en in vrows:
        continue
      if export_all or zh == "" or zh == en:
        vrows[en] = zh or en
  with open(vdst, "wb") as f:
    f.write(
        "\n".join(f"@\t{one_line(e)}\t{one_line(z)}" for e, z in vrows.items()).encode("utf-8")
    )
  print(f"已匯出值翻譯 {len(vrows)} 筆到 {vdst}（依英文值去重；同一英文值只需翻一次，會套用到所有標籤）")
  print("格式：@<TAB>英文值<TAB>中文。要讓某個標籤用不同譯文，可手動加一行：@標籤名稱<TAB>英文值<TAB>中文（優先於通用譯文）。")


# ----------------------------------------------------------------- 4. import
def load_translations(paths):
  exact = {}
  by_name = defaultdict(set)
  names = set()
  vglobal = {}
  vtag = {}
  bad = 0
  for p in paths:
    for line in read_text(p).split("\n"):
      line = line.rstrip("\r")
      if not line.strip() or line.startswith("#"):
        continue
      parts = line.split("\t")
      if len(parts) >= 3 and parts[0].startswith("@"):
        vname, en, zh = parts[0][1:].strip(), parts[1], "\t".join(parts[2:]).strip()
        if zh:
          if vname:
            vtag[(vname, en)] = zh
          else:
            vglobal[en] = zh
      elif len(parts) >= 3:
        name, en, zh = parts[0].strip(), parts[1], "\t".join(parts[2:]).strip()
        if zh:
          exact[(name, en)] = zh
          by_name[name].add(zh)
          names.add(name)
      elif len(parts) == 2 and parts[1].strip():
        name, zh = parts[0].strip(), parts[1].strip()
        by_name[name].add(zh)
        names.add(name)
      else:
        bad += 1
  return exact, by_name, names, bad, vglobal, vtag


def cmd_import(prep=DEF_PREP, txts=None, orig=DEF_SRC, overwrite=False):
  if not txts:
    txts = [p for p in (DEF_TXT, DEF_VTXT) if os.path.exists(p)]
    if not txts:
      raise ValueError(f"找不到翻譯檔（{DEF_TXT} / {DEF_VTXT}）")
  exact, by_name, names, bad_rows, vglobal, vtag = load_translations(txts)
  lines = split_lines(read_text(prep))
  tags = parse_tags(lines)

  # 保護原檔「本來就有」的 zh_tw（以標籤順序對應原檔），預設不覆蓋
  protected = set()
  vprot = set()
  prot_info = False
  if orig and os.path.exists(orig) and not overwrite:
    otags = parse_tags(split_lines(read_text(orig)))
    if len(otags) == len(tags):
      prot_info = True
      protected = {i for i, o in enumerate(otags) if o["zh_idx"] is not None}
      for i, o in enumerate(otags):
        for j, kk in enumerate(o["keys"]):
          if kk["zh"] is not None:
            vprot.add((i, j))
    else:
      print("警告：原檔與修改版標籤數不同，無法辨識既有 zh_tw，改為只更新仍與英文相同者")

  changed = missing_zh = kept = 0
  matched = set()
  for idx, t in enumerate(tags):
    if t["bad"]:
      continue
    if t["zh_idx"] is None:
      missing_zh += 1
      continue
    en = un(t["base"])
    cur = un(t["zh_raw"])
    zh = exact.get((t["name"], en))
    if zh is None:
      # 備援：僅對「尚未翻譯」的標籤，且該名稱只有一種譯文時，才用名稱比對
      cands = by_name.get(t["name"], set())
      if (cur == "" or cur == en) and len(cands) == 1:
        zh = next(iter(cands))
    if zh is None:
      continue
    matched.add(t["name"])
    new_raw = xml_escape(zh)
    if idx in protected or (not prot_info and not overwrite and cur not in ("", en)):
      if new_raw != t["zh_raw"]:
        kept += 1  # 既有譯文與翻譯檔不同，保留既有
      continue
    if new_raw != t["zh_raw"]:
      m = RE_ZH_PARTS.match(lines[t["zh_idx"]])
      lines[t["zh_idx"]] = m.group(1) + new_raw + m.group(4)
      changed += 1

  # 值翻譯：先找「標籤名稱＋英文值」的專屬譯文，再用依英文值通用的譯文
  vchanged = vkept = vmissing = 0
  if vglobal or vtag:
    for idx, t in enumerate(tags):
      if t["bad"]:
        continue
      for j, kk in enumerate(t["keys"]):
        if kk["zh_idx"] is None:
          vmissing += 1
          continue
        en = un(kk["en"])
        cur = un(kk["zh"])
        zh = vtag.get((t["name"], en)) or vglobal.get(en)
        if zh is None:
          continue
        new_raw = xml_escape(zh)
        if (idx, j) in vprot or (not prot_info and not overwrite and cur not in ("", en)):
          if new_raw != kk["zh"]:
            vkept += 1
          continue
        if new_raw != kk["zh"]:
          m = RE_VAL_ZH_PARTS.match(lines[kk["zh_idx"]])
          lines[kk["zh_idx"]] = m.group(1) + new_raw + m.group(4)
          vchanged += 1

  write_text(prep, "\n".join(lines))
  pending = {
      (t["name"], un(t["base"]))
      for t in parse_tags(lines)
      if not t["bad"] and (not t["zh_raw"] or un(t["zh_raw"]) == un(t["base"]))
  }
  print(f"已更新 {changed} 個 zh_tw 項目 → {prep}")
  if kept:
    print(f"  另有 {kept} 個原檔既有的 zh_tw 與翻譯檔不同，已保留既有譯文（加 --overwrite-existing 可改為覆蓋）")
  if bad_rows:
    print(f"翻譯檔中格式錯誤（欄位不足）的行：{bad_rows}")
  if missing_zh:
    print(f"有 {missing_zh} 個標籤沒有 zh_tw，請先執行 prepare")
  unmatched = sorted(names - matched)
  if unmatched:
    print(f"翻譯檔中有 {len(unmatched)} 個名稱在 XML 找不到，例如：{unmatched[:8]}")
  print(f"目前仍與英文相同（視為未翻譯）：{len(pending)} 種")
  if vglobal or vtag:
    print(f"值翻譯：更新 {vchanged} 個 zh_tw 值")
    if vkept:
      print(f"  另有 {vkept} 個原檔既有的 zh_tw 值與翻譯檔不同，已保留既有譯文")
    if vmissing:
      print(f"  有 {vmissing} 個值沒有 zh_tw 行，請先執行 prepare（含值）")
    vpend = {
        un(kk["en"])
        for t in parse_tags(lines)
        if not t["bad"]
        for kk in t["keys"]
        if not kk["zh"] or un(kk["zh"]) == un(kk["en"])
    }
    print(f"  值：仍與英文相同的英文值 {len(vpend)} 種")
  if orig and os.path.exists(orig):
    cmd_check(orig, prep)


# ----------------------------------------------------------------- 舊版翻譯檔轉換
def cmd_convert(old_txt, src=DEF_SRC, dst="to_translate_converted.txt"):
  """舊版格式 (標籤id<TAB>文字) → 新版格式 (名稱<TAB>英文<TAB>中文)。
  依 out.xml 內相同 id 的標籤出現順序，一對一還原成標籤名稱。"""
  tags = parse_tags(split_lines(read_text(src)))
  queues = defaultdict(list)
  for t in tags:
    if t["id"] and not t["bad"]:
      queues[t["id"]].append(t)
  total_tags = sum(len(v) for v in queues.values())

  rows = {}
  conflicts = unmatched = same = 0
  used = 0
  for line in read_text(old_txt).split("\n"):
    line = line.rstrip("\r")
    if not line.strip():
      continue
    parts = line.split("\t", 1)
    if len(parts) != 2:
      unmatched += 1
      continue
    tid, zh = parts[0].strip(), parts[1].strip()
    q = queues.get(tid)
    if not q:
      unmatched += 1
      continue
    t = q.pop(0)
    used += 1
    en = un(t["base"])
    if not zh or zh == en:
      same += 1
      continue
    key = (t["name"], en)
    if key in rows:
      if rows[key] != zh:
        conflicts += 1
      continue
    rows[key] = zh
  with open(dst, "wb") as f:
    f.write(
        "\n".join(
            f"{n}\t{one_line(e)}\t{one_line(z)}" for (n, e), z in rows.items()
        ).encode("utf-8")
    )
  print(f"舊檔行數對應成功：{used} / XML 內有 id 的標籤 {total_tags}")
  print(f"轉換出 {len(rows)} 筆有翻譯的資料 → {dst}")
  print(f"  與英文相同而略過：{same}")
  if unmatched:
    print(f"  無法對應（id 不在 XML 或格式錯誤）：{unmatched}")
  if conflicts:
    print(f"  同名同英文但譯文不同（採第一個）：{conflicts}")
  if used != total_tags:
    print("  提醒：行數與 XML 標籤數不同，可能有行被刪除或合併，請檢查結果。")


# ----------------------------------------------------------------- 官方 .pm 解析（合併用）
def parse_lang_pm(path):
  """解析 ExifTool Lang 模組的 %...::Translate，回傳 (版本, {名稱: 字串 或 dict})。
  只支援該檔實際使用的語法：單/雙引號字串、裸字鍵、巢狀 {}、=>、逗號、# 註解。"""
  text = read_text(path)
  m = re.search(r"%Image::ExifTool::Lang::\w+::Translate\s*=\s*\(", text)
  if not m:
    raise ValueError(f"{path} 找不到 %Image::ExifTool::Lang::*::Translate")
  vm = re.search(r"\$VERSION\s*=\s*'([^']*)'", text)
  version = vm.group(1) if vm else None
  n = len(text)

  def skip(i):
    while i < n:
      if text[i].isspace():
        i += 1
      elif text[i] == "#":
        while i < n and text[i] != "\n":
          i += 1
      else:
        break
    return i

  def string(i):
    q = text[i]
    i += 1
    buf = []
    while i < n and text[i] != q:
      if text[i] == "\\" and i + 1 < n and text[i + 1] in (q, "\\"):
        buf.append(text[i + 1])
        i += 2
      else:
        buf.append(text[i])
        i += 1
    return "".join(buf), i + 1

  def value(i):
    i = skip(i)
    if text[i] == "{":
      return hash_(i + 1, "}")
    if text[i] in "'\"":
      return string(i)
    raise ValueError(f"{path} 位置 {i} 無法解析：{text[i:i + 30]!r}")

  def hash_(i, end):
    d = {}
    while True:
      i = skip(i)
      if text[i] == end:
        return d, i + 1
      if text[i] in "'\"":
        k, i = string(i)
      else:
        km = re.compile(r"[A-Za-z_]\w*").match(text, i)
        if not km:
          raise ValueError(f"{path} 位置 {i} 無法解析鍵：{text[i:i + 30]!r}")
        k, i = km.group(0), km.end()
      i = skip(i)
      if text[i:i + 2] != "=>":
        raise ValueError(f"{path} 位置 {i} 缺少 =>")
      v, i = value(i + 2)
      d[k] = v
      i = skip(i)
      if i < n and text[i] == ",":
        i += 1

  tr, _ = hash_(m.end(), ")")
  return version, tr


# ----------------------------------------------------------------- 5. build
def perl_quote(s):
  return "'" + s.replace("\\", "\\\\").replace("'", "\\'") + "'"


def load_overrides(path):
  """name_overrides.txt：每行 名稱<TAB>譯文（第三欄起為備註，# 開頭為註解）。"""
  ov = {}
  if not path or not os.path.exists(path):
    return ov
  for line in read_text(path).split("\n"):
    line = line.rstrip("\r")
    if not line.strip() or line.lstrip().startswith("#"):
      continue
    parts = line.split("\t")
    if len(parts) >= 2 and parts[0].strip() and parts[1].strip():
      ov[parts[0].strip()] = parts[1].strip()
  return ov


def cmd_build(src=DEF_PREP, dst=DEF_PM, overrides=DEF_OVR, report=DEF_CONF, base=None):
  if base is None and os.path.exists(DEF_BASE):
    base = DEF_BASE
  text = read_text(src)
  raw_val_zh = len(re.findall(r"""<val\s+lang=['"]zh_tw['"]\s*>(?=[^<])""", text))
  tags = parse_tags(split_lines(text))
  trans = {}
  variants = {}  # name -> [(zh, en, table, id)]
  conflicts = set()
  pconv = {}  # name -> {英文值: 中文值}（保留 XML 內既有的 <val lang='zh_tw'>）
  val_conflicts = 0
  val_pairs = 0
  for t in tags:
    if not t["name"]:
      continue
    name = t["name"]
    for en_raw, zh_raw in t["vals"]:
      val_pairs += 1
      ven, vzh = un(en_raw), one_line(un(zh_raw))
      if not vzh or vzh == ven:
        continue
      d = pconv.setdefault(name, {})
      if ven in d and d[ven] != vzh:
        val_conflicts += 1
        continue
      d.setdefault(ven, vzh)
    if t["bad"] or not t["zh_raw"]:
      continue
    zh = un(t["zh_raw"])
    if zh == un(t["base"]):  # 與英文相同＝未翻譯，不寫入
      continue
    zh = one_line(zh)
    variants.setdefault(name, []).append((zh, un(t["base"]), t["table"], t["id"]))
    if name in trans:
      if trans[name] != zh:
        conflicts.add(name)
      continue
    trans[name] = zh

  # 衝突報告 + 使用者覆寫
  ov = load_overrides(overrides)
  applied = 0
  all_names = {t["name"] for t in tags}
  for name, zh in ov.items():
    if name in all_names:
      trans[name] = zh
      applied += 1
  if conflicts:
    lines_r = [
        "# 同名標籤在不同表有不同譯文。Lang 檔以標籤名稱為鍵，一個名稱只能有一個譯文。",
        "# 預設採用 [採用] 那一種。若要指定，請把想用的那一行（前兩欄）複製到 "
        + (overrides or DEF_OVR),
        "# 格式：名稱<TAB>譯文",
        "",
    ]
    for name in sorted(conflicts):
      chosen = ov.get(name, trans[name])
      seen = {}
      for zh, en, table, tid in variants[name]:
        seen.setdefault(zh, []).append((en, table, tid))
      lines_r.append(f"## {name}（{len(seen)} 種譯文）")
      for zh, occ in seen.items():
        mark = "[採用] " if zh == chosen else ""
        info = "; ".join(f"{tb}({i}) en={e}" for e, tb, i in occ[:3])
        more = f" …共 {len(occ)} 處" if len(occ) > 3 else ""
        lines_r.append(f"{name}\t{zh}\t# {mark}{info}{more}")
      lines_r.append("")
    with open(report, "wb") as f:
      f.write("\n".join(lines_r).encode("utf-8"))

  # 與官方 .pm 合併：XML 沒涵蓋到的官方項目原樣保留；XML（含你的修改）優先
  version = "1.00"
  kept_desc = kept_vals = 0
  if base:
    bver, btr = parse_lang_pm(base)
    version = bver or version
    for name, v in btr.items():
      bdesc = v if isinstance(v, str) else v.get("Description")
      bpc = {} if isinstance(v, str) else (v.get("PrintConv") or {})
      if bdesc is not None and name not in trans:
        trans[name] = bdesc
        kept_desc += 1
      if bpc:
        d = pconv.setdefault(name, {})
        for ven, vzh in bpc.items():
          if ven not in d:
            d[ven] = vzh
            kept_vals += 1

  out = [
      "#------------------------------------------------------------------------------",
      "# File:         zh_tw.pm",
      "#",
      "# Description:  ExifTool Traditional Chinese language translations",
      "#",
      "# Notes:        Generated from exiftool -listx XML by exiftool_zh_tw_tool.py",
      "#------------------------------------------------------------------------------",
      "",
      "package Image::ExifTool::Lang::zh_tw;",
      "",
      "use strict;",
      "use vars qw($VERSION);",
      "",
      f"$VERSION = '{version}';",
      "",
      "%Image::ExifTool::Lang::zh_tw::Translate = (",
  ]
  nvals = 0
  for name in sorted(set(trans) | set(pconv)):
    if name in pconv:
      out.append(f"   {perl_quote(name)} => {{")
      if name in trans:
        out.append(f"      Description => {perl_quote(trans[name])},")
      out.append("      PrintConv => {")
      for ven in sorted(pconv[name]):
        out.append(f"        {perl_quote(ven)} => {perl_quote(pconv[name][ven])},")
        nvals += 1
      out.append("      },")
      out.append("    },")
    else:
      out.append(f"   {perl_quote(name)} => {perl_quote(trans[name])},")
  out += [");", "", "1;  # end", ""]
  with open(dst, "wb") as f:
    f.write("\n".join(out).encode("utf-8"))

  print(f"已產生 {dst}：標籤說明 {len(trans)} 個，值翻譯 {nvals} 個（來自 {len(pconv)} 個標籤）")
  if base:
    print(f"  已合併官方檔 {base}：補回 XML 未涵蓋的說明 {kept_desc} 個、值翻譯 {kept_vals} 個")
  print(
      f"輸入檔 {src} 內 <val lang='zh_tw'> 共 {raw_val_zh} 行，"
      f"讀入 {val_pairs} 組，寫入 .pm 的值翻譯 {nvals} 個"
  )
  if raw_val_zh == 0:
    print(
        "警告：輸入檔沒有任何 <val lang='zh_tw'>，所以 .pm 不可能有值翻譯。\n"
        "      請確認 out_prepared.xml 是由「有中文值翻譯」的原始 out.xml 產生"
        "（out.xml 若是用缺少值翻譯的 zh_tw.pm 匯出，就已經沒有了）。"
    )
  elif raw_val_zh > val_pairs:
    print(f"警告：有 {raw_val_zh - val_pairs} 行 zh_tw 值沒有成功讀入（格式不符），請回報範例。")
  if applied:
    print(f"已套用 {applied} 個自訂覆寫（{overrides}）")
  unresolved = sorted(n for n in conflicts if n not in ov)
  if unresolved:
    print(
        f"注意：{len(unresolved)} 個同名標籤在不同表有不同譯文，預設採用第一個：",
        unresolved[:8],
    )
    print(f"  各種譯文與出處已寫入 {report}；要指定請建立 {overrides}（名稱<TAB>譯文）")
  elif conflicts:
    print(f"同名衝突 {len(conflicts)} 個，皆已由 {overrides} 指定譯文（明細見 {report}）")
  if val_conflicts:
    print(f"注意：{val_conflicts} 個同名標籤的同一個值有不同譯文，僅採用第一個")


# ----------------------------------------------------------------- 6. compare（診斷）
def scan_stats(path):
  lines = split_lines(read_text(path))
  tags = parse_tags(lines)
  tables = set()
  table_zh = set()
  table_zh_lines = 0
  cur_table = None
  in_tag = False
  for line in lines:
    if RE_TAG_START.match(line):
      in_tag = True
    elif RE_TAG_END.match(line):
      in_tag = False
    elif RE_TABLE.match(line):
      m = RE_NAME.search(line)
      cur_table = un(m.group(2)) if m else "?"
      tables.add(cur_table)
    elif not in_tag:
      dm = RE_DESC.match(line)
      if dm and dm.group(3) == "zh_tw":
        table_zh.add(cur_table)
        table_zh_lines += 1
  tag_zh = {t["name"] for t in tags if t["zh_idx"] is not None}
  tag_zh_lines = sum(1 for t in tags if t["zh_idx"] is not None)
  val_zh = set()
  val_zh_lines = 0
  for t in tags:
    for en_raw, _ in t["vals"]:
      val_zh.add((t["name"], un(en_raw)))
      val_zh_lines += 1
  return {
      "tables": tables,
      "tags": len(tags),
      "table_zh": table_zh,
      "table_zh_lines": table_zh_lines,
      "tag_zh": tag_zh,
      "tag_zh_lines": tag_zh_lines,
      "val_zh": val_zh,
      "val_zh_lines": val_zh_lines,
  }


def cmd_compare(a_path, b_path, report=None):
  A, B = scan_stats(a_path), scan_stats(b_path)
  print(f"{'項目':<22}{'A(' + os.path.basename(a_path) + ')':>20}{'B(' + os.path.basename(b_path) + ')':>20}")
  rows = [
      ("表格數", len(A["tables"]), len(B["tables"])),
      ("標籤數", A["tags"], B["tags"]),
      ("表格說明 zh_tw 行數", A["table_zh_lines"], B["table_zh_lines"]),
      ("標籤說明 zh_tw 行數", A["tag_zh_lines"], B["tag_zh_lines"]),
      ("值 zh_tw 行數", A["val_zh_lines"], B["val_zh_lines"]),
  ]
  for name, x, y in rows:
    print(f"{name:<22}{x:>20}{y:>20}")

  miss = {
      "A 有、B 沒有的表格": sorted(A["tables"] - B["tables"]),
      "A 有、B 沒有的表格說明 zh_tw": sorted(x for x in A["table_zh"] - B["table_zh"] if x),
      "A 有、B 沒有的標籤說明 zh_tw": sorted(A["tag_zh"] - B["tag_zh"]),
      "A 有、B 沒有的值 zh_tw（標籤 | 英文值）": sorted(
          f"{n} | {v}" for n, v in A["val_zh"] - B["val_zh"]
      ),
      "B 有、A 沒有的表格": sorted(B["tables"] - A["tables"]),
  }
  print()
  for title, items in miss.items():
    print(f"{title}：{len(items)}" + (f"  例：{items[:5]}" if items else ""))
  if report:
    with open(report, "wb") as f:
      body = []
      for title, items in miss.items():
        body.append(f"## {title}（{len(items)}）")
        body.extend(items)
        body.append("")
      f.write("\n".join(body).encode("utf-8"))
    print(f"\n完整清單已寫入 {report}")


# ----------------------------------------------------------------- 入口
def ask(prompt, default):
  v = input(f"{prompt} [{default}]: ").strip().strip('"')
  return v or default


def menu():
  while True:
    print(
        "\n=== ExifTool zh_tw 工具 ===\n"
        "1. 匯出 out.xml（exiftool -listx -all:all > out.xml）\n"
        "2. 補充 zh_tw（out.xml → out_prepared.xml，含自動比對）\n"
        "3. 匯出翻譯檔\n"
        "4. 匯入翻譯檔\n"
        "5. 產生 zh_tw.pm\n"
        "6. 比較兩份 out.xml（診斷遺失的 zh_tw 項目）\n"
        "0. 離開"
    )
    c = input("選擇: ").strip()
    try:
      if c == "1":
        dst = ask("輸出", DEF_SRC)
        force = True
        if os.path.exists(dst):
          force = ask(f"{dst} 已存在，覆蓋？(y/n)", "n").lower() == "y"
        if force:
          cmd_listx(dst, ask("標籤範圍（群組名如 exif / all，或完整參數）", "-all:all"), None, True)
      elif c == "2":
        cmd_prepare(ask("原檔", DEF_SRC), ask("輸出", DEF_PREP))
      elif c == "3":
        a = ask("匯出全部含已翻譯？(y/n)", "n").lower() == "y"
        cmd_export(ask("修改版", DEF_PREP), ask("輸出", DEF_TXT), a)
      elif c == "4":
        prep = ask("修改版", DEF_PREP)
        txt = ask("翻譯檔（可用逗號分隔多個）", f"{DEF_TXT},{DEF_VTXT}")
        ow = ask("覆蓋原檔既有的 zh_tw？(y/n)", "n").lower() == "y"
        cmd_import(prep, [x.strip() for x in txt.split(",") if x.strip()], DEF_SRC, ow)
      elif c == "5":
        cmd_build(ask("修改版", DEF_PREP), ask("輸出", DEF_PM))
      elif c == "6":
        cmd_compare(
            ask("A（舊/官方）", "out_old.xml"),
            ask("B（新）", DEF_SRC),
            ask("完整清單輸出", "compare_report.txt"),
        )
      elif c == "0":
        return
    except Exception as e:
      print(f"[錯誤] {e}")


def main():
  if len(sys.argv) == 1:
    menu()
    return 0

  ap = argparse.ArgumentParser(description="ExifTool zh_tw 語言檔製作工具")
  sub = ap.add_subparsers(dest="cmd", required=True)

  p = sub.add_parser("listx", help="執行 exiftool -listx 匯出 out.xml")
  p.add_argument("-o", "--output", default=DEF_SRC)
  p.add_argument("-g", "--group", default="-all:all", help="標籤範圍；預設 -all:all，可輸入 exif（=-exif:all）、none（不加參數）或完整參數")
  p.add_argument("--exiftool", default=None, help="exiftool 的路徑")
  p.add_argument("-f", "--force", action="store_true", help="覆蓋既有檔案")

  p = sub.add_parser("prepare", help="補充 zh_tw 項目（說明與值）")
  p.add_argument("-i", "--input", default=DEF_SRC)
  p.add_argument("-o", "--output", default=DEF_PREP)
  p.add_argument("--no-values", action="store_true", help="不補 <val lang='zh_tw'>")

  p = sub.add_parser("check", help="比對原檔與修改版")
  p.add_argument("-a", "--orig", default=DEF_SRC)
  p.add_argument("-b", "--mod", default=DEF_PREP)

  p = sub.add_parser("export", help="匯出翻譯檔")
  p.add_argument("-i", "--input", default=DEF_PREP)
  p.add_argument("-o", "--output", default=DEF_TXT)
  p.add_argument("--all", action="store_true", help="連已翻譯的也匯出")
  p.add_argument("--values-output", default=DEF_VTXT, help="值翻譯輸出檔")
  p.add_argument("--no-values", action="store_true", help="不匯出值翻譯")

  p = sub.add_parser("import", help="匯入翻譯檔（就地更新修改版）")
  p.add_argument("-i", "--input", default=DEF_PREP)
  p.add_argument("-t", "--txt", nargs="+", default=None, help="翻譯檔（預設 to_translate.txt 與 to_translate_values.txt）")
  p.add_argument("--orig", default=DEF_SRC, help="匯入後用來比對的原檔")
  p.add_argument(
      "--overwrite-existing",
      action="store_true",
      help="允許覆蓋原檔既有的 zh_tw（預設保護；覆蓋後自動比對會失敗屬正常）",
  )

  p = sub.add_parser("convert", help="轉換舊版翻譯檔（id<TAB>文字）")
  p.add_argument("-t", "--txt", required=True, help="舊版翻譯檔")
  p.add_argument("-x", "--xml", default=DEF_SRC, help="原檔 out.xml")
  p.add_argument("-o", "--output", default="to_translate_converted.txt")

  p = sub.add_parser("compare", help="比較兩份 out.xml 的 zh_tw 數量與遺失項目")
  p.add_argument("a", help="A：舊的/官方的 out.xml")
  p.add_argument("b", help="B：新的 out.xml")
  p.add_argument("-o", "--output", default="compare_report.txt")

  p = sub.add_parser("build", help="產生 zh_tw.pm")
  p.add_argument("-i", "--input", default=DEF_PREP)
  p.add_argument("-o", "--output", default=DEF_PM)
  p.add_argument("--overrides", default=DEF_OVR, help="同名標籤指定譯文（名稱<TAB>譯文）")
  p.add_argument("--report", default=DEF_CONF, help="同名衝突報告輸出檔")
  p.add_argument(
      "--base",
      default=None,
      help=f"官方 zh_tw.pm（合併用，保留 XML 沒涵蓋的官方項目）；預設自動使用 {DEF_BASE}（若存在）",
  )

  a = ap.parse_args()
  try:
    if a.cmd == "listx":
      cmd_listx(a.output, a.group, a.exiftool, a.force)
    elif a.cmd == "prepare":
      ok = cmd_prepare(a.input, a.output, not a.no_values)
      return 0 if ok else 1
    if a.cmd == "check":
      return 0 if cmd_check(a.orig, a.mod) else 1
    if a.cmd == "export":
      cmd_export(a.input, a.output, a.all, a.values_output, not a.no_values)
    elif a.cmd == "import":
      cmd_import(a.input, a.txt, a.orig, a.overwrite_existing)
    elif a.cmd == "convert":
      cmd_convert(a.txt, a.xml, a.output)
    elif a.cmd == "compare":
      cmd_compare(a.a, a.b, a.output)
    elif a.cmd == "build":
      cmd_build(a.input, a.output, a.overrides, a.report, a.base)
  except Exception as e:
    print(f"[錯誤] {e}")
    return 1
  return 0


if __name__ == "__main__":
  sys.exit(main())
