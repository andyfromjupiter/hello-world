import json
import os
import shutil
import time
import re
import win32com.client as win32

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
HWP_TEMPLATE_PATH = os.path.join(BASE_DIR, "템플릿_변형문제.hwp") 
DATA_FILENAME = "JSON.txt"
OUTPUT_FILENAME = "변형문제 샘플.hwp" 
TEMP_DIR = os.path.join(BASE_DIR, "temp_files")

def init_hwp():
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = True
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        return hwp
    except Exception as e:
        print(f"HWP Error: {e}")
        return None

def load_json_data(filepath):
    for enc in ["utf-8-sig", "utf-8", "cp949"]:
        try:
            with open(filepath, "r", encoding=enc) as f:
                raw_text = f.read().strip()
                if raw_text.startswith("```json"):
                    raw_text = raw_text.replace("```json", "", 1)
                if raw_text.endswith("```"):
                    raw_text = raw_text[::-1].replace("```", "", 1)[::-1]
                
                raw_text = re.sub(r'//.*', '', raw_text)
                return json.loads(raw_text)
        except: continue
    raise ValueError("JSON Parsing Failed")

def set_style(hwp, bold=None, underline=None, color=None, backcolor=None):
    act = hwp.CreateAction("CharShape")
    pset = act.CreateSet()
    act.GetDefault(pset)
    if bold is not None: pset.SetItem("Bold", 1 if bold else 0)
    if underline is not None: pset.SetItem("UnderlineType", 1 if underline else 0)
    if color is not None: pset.SetItem("TextColor", color)
    if backcolor is not None: pset.SetItem("BackColor", backcolor)
    act.Execute(pset)

def insert_text(hwp, text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = str(text)
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def process_and_insert_tags(hwp, text_block):
    lines = str(text_block).split('\n')
    for i, line in enumerate(lines):
        parts = re.split(r'(<u>|</u>|<b>|</b>|<r>|</r>|<y>|</y>)', line)
        for part in parts:
            if part == '<u>': set_style(hwp, underline=True)
            elif part == '</u>': set_style(hwp, underline=False)
            elif part == '<b>': set_style(hwp, bold=True)
            elif part == '</b>': set_style(hwp, bold=False)
            elif part == '<r>': set_style(hwp, color=255)
            elif part == '</r>': set_style(hwp, color=0)
            elif part == '<y>': set_style(hwp, backcolor=13434879) 
            elif part == '</y>': set_style(hwp, backcolor=4294967295) 
            elif part: insert_text(hwp, part)
        if i < len(lines) - 1: hwp.HAction.Run("BreakPara")

def insert_keep_style(hwp, field_name, text):
    text_str = str(text)
    if not text_str or text_str.strip().lower() == "null":
        hwp.PutFieldText(field_name, " ")
        return

    if not re.search(r'(<u>|</u>|<b>|</b>|<r>|</r>|<y>|</y>)', text_str):
        hwp.PutFieldText(field_name, text_str.replace('\n', '\r\n'))
        return

    targets = [field_name] + [f"{field_name}{{{i}}}" for i in range(1, 50)]
    for target in targets:
        if hwp.MoveToField(target, True, True, True): 
            act = hwp.CreateAction("CharShape")
            pset = act.CreateSet()
            act.GetDefault(pset)
            hwp.PutFieldText(target, "")
            if hwp.MoveToField(target, True, False, True):
                act.Execute(pset)
                process_and_insert_tags(hwp, text_str)
            hwp.Run("Cancel")

def insert_table_data(hwp, field_name, data_list):
    targets = [field_name] + [f"{field_name}{{{i}}}" for i in range(1, 50)]
    for target in targets:
        if hwp.MoveToField(target, True, False, True):
            hwp.PutFieldText(target, "")
            for row_idx, row_data in enumerate(data_list):
                for col_idx, cell_data in enumerate(row_data):
                    process_and_insert_tags(hwp, cell_data)
                    if col_idx < len(row_data) - 1: hwp.HAction.Run("TableRightCell")
                if row_idx < len(data_list) - 1:
                    hwp.HAction.Run("TableLowerCell")
                    for _ in range(len(row_data) - 1): hwp.HAction.Run("TableLeftCell")
            hwp.Run("Cancel")

def process_fields_and_rows(hwp, content):
    for key, val in content.items():
        if val is None: val = " "
        is_table_data = isinstance(val, list) and len(val) > 0 and isinstance(val[0], list)
        if not is_table_data:
            if isinstance(val, list): val = "\n".join(str(x) for x in val)
            else: val = str(val)
            if val.strip().lower() == "null" or not val: val = " "
        
        key_variations = {key, key.lower(), key.upper(), key.capitalize()}
        k_lower = key.lower()
        
        if "_" in k_lower:
            key_variations.add(k_lower.replace("_", ""))
            key_variations.add(key.replace("_", ""))
            parts = key.split('_')
            if len(parts) == 2:
                key_variations.add(parts[0].lower() + parts[1].capitalize())
                
        if k_lower in ['n', 'no', 'num', 'number']:
            key_variations.update(['n', 'N', 'No', 'NO', 'no', 'num', 'Num', 'NUM'])
            
        if k_lower in ['ans_tf', 'anstf']:
            key_variations.update(['ans_TF', 'ansTF', 'ANS_TF', 'TFA', 'ans_Tf'])

        for t_key in key_variations:
            try:
                if is_table_data: insert_table_data(hwp, t_key, val)
                else: insert_keep_style(hwp, t_key, val)
            except: pass

    for j in range(1, 31):
        val1 = str(content.get(f"e{j}", "")).strip()
        val2 = str(content.get(f"E{j}", "")).strip()
        
        if (not val1 or val1.lower() == "null") and (not val2 or val2.lower() == "null"):
            for base_name in [f"e{j}", f"E{j}"]:
                targets = [base_name] + [f"{base_name}{{{i}}}" for i in range(1, 20)]
                for target in targets:
                    if hwp.MoveToField(target, True, False, True): 
                        try:
                            act = hwp.CreateAction("CellShape")
                            pset = act.CreateSet()
                            if act.GetDefault(pset):
                                hwp.Run("TableDeleteRow") 
                            else:
                                hwp.PutFieldText(target, " ")
                        except:
                            hwp.PutFieldText(target, " ")

    for prefix in ['w', 'W', 's', 'S', 'v', 'V']:
        for j in range(1, 31):
            val = str(content.get(f"{prefix}{j}", "")).strip()
            if not val or val.lower() == "null":
                for base_name in [f"{prefix}{j}"]:
                    targets = [base_name] + [f"{base_name}{{{i}}}" for i in range(1, 20)]
                    for target in targets:
                        try: hwp.PutFieldText(target, " ")
                        except: pass

    passage_no = ""
    for k in ["n", "N", "No", "NO", "num", "Num", "NUM"]:
        if content.get(k):
            passage_no = str(content.get(k)).strip()
            break

    if passage_no:
        image_path = os.path.join(BASE_DIR, f"{passage_no}.jpeg")
        if os.path.exists(image_path):
            for base_pic in ["pic", "PIC"]:
                targets = [base_pic] + [f"{base_pic}{{{i}}}" for i in range(1, 10)]
                for target in targets:
                    if hwp.MoveToField(target, True, False, True):
                        hwp.PutFieldText(target, ""); hwp.MoveToField(target, True, False, True)
                        hwp.InsertPicture(image_path, True, 3, False, False, 0)
                        hwp.Run("Cancel")

def main():
    if not os.path.exists(DATA_FILENAME): return
    all_data = load_json_data(DATA_FILENAME)
    if isinstance(all_data, dict): all_data = [all_data]
    
    hwp = init_hwp()
    if not hwp: return
    if os.path.exists(TEMP_DIR): shutil.rmtree(TEMP_DIR, ignore_errors=True)
    os.makedirs(TEMP_DIR, exist_ok=True)

    try:
        for i, content in enumerate(all_data):
            hwp.Open(HWP_TEMPLATE_PATH)
            time.sleep(0.3) 
            print(f"Processing... [{i+1}/{len(all_data)}]")
            
            process_fields_and_rows(hwp, content)
            
            temp_path = os.path.join(TEMP_DIR, f"temp_{i:02d}.hwp")
            hwp.SaveAs(temp_path)
            hwp.Clear(1) 
            time.sleep(0.2)

        print("Merging files...")
        time.sleep(1.0)
        
        temp_files = sorted([os.path.join(TEMP_DIR, f) for f in os.listdir(TEMP_DIR) if f.endswith(".hwp")])
        if temp_files:
            hwp.Open(temp_files[0])
            time.sleep(0.5)
            
            for f_path in temp_files[1:]:
                hwp.HAction.Run("Cancel"); hwp.HAction.Run("MoveDocEnd")
                hwp.HAction.Run("MoveRight"); hwp.HAction.Run("MoveRight"); hwp.HAction.Run("MoveDocEnd")
                hwp.HAction.Run("BreakSection"); time.sleep(0.1)
                act = hwp.CreateAction("InsertFile"); pset = act.CreateSet(); act.GetDefault(pset)
                pset.SetItem("FileName", f_path); pset.SetItem("KeepSection", 1); act.Execute(pset)
                time.sleep(0.1)
            
            patterns = [(r"\[[^\]]*\]", True), (r"\([a-zA-Z]\)[ ]*_+", True), (r"\[[ ]*T[ ]*/[ ]*F[ ]*\]", False)]
            for regex, is_bold in patterns:
                hwp.HAction.Run("MoveDocBegin")
                find_ps = hwp.HParameterSet.HFindReplace; hwp.HAction.GetDefault("FindReplace", find_ps.HSet)
                find_ps.HSet.SetItem("FindString", regex); find_ps.HSet.SetItem("FindRegExp", 1)
                while hwp.HAction.Execute("RepeatFind", find_ps.HSet):
                    set_style(hwp, bold=is_bold); hwp.HAction.Run("MoveRight")
                hwp.Run("Cancel") 

            hwp.SaveAs(os.path.join(BASE_DIR, OUTPUT_FILENAME))
            print(f"\nDone! Saved at:\n{os.path.join(BASE_DIR, OUTPUT_FILENAME)}")
            
    except Exception as e:
        print(f"\nError: {e}")
    finally:
        shutil.rmtree(TEMP_DIR, ignore_errors=True)

if __name__ == "__main__":
    main()
