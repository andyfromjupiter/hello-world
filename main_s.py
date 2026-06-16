import customtkinter as ctk
import tkinter.messagebox as messagebox
from tkinter import filedialog
import requests
import json
import os
import shutil
import time
import re
import threading
import win32com.client as win32
import google.generativeai as genai

GAS_URL = "https://script.google.com/macros/s/AKfycbzDNtdOr7ZCpU5YOUHnV4duFdgOruv_vR_DEqavtBPyracsT3wok-7seqisc_hScgo/exec"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMP_DIR = os.path.join(BASE_DIR, "temp_files")
AI_TIMEOUT = 300

def init_hwp():
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = True
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        return hwp
    except Exception as e:
        print(f"HWP Error: {e}")
        return None

def set_style(hwp, bold=None, underline=None, color=None, shadecolor=None):
    act = hwp.CreateAction("CharShape")
    pset = act.CreateSet()
    act.GetDefault(pset)
    if bold is not None: pset.SetItem("Bold", 1 if bold else 0)
    if underline is not None: pset.SetItem("UnderlineType", 1 if underline else 0)
    if color is not None: pset.SetItem("TextColor", color)
    if shadecolor is not None: pset.SetItem("ShadeColor", shadecolor)
    act.Execute(pset)

def insert_text(hwp, text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = str(text)
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

def process_and_insert_tags(hwp, text_block):
    clean_text = str(text_block).replace('\\n', '\n').replace('\r', '')
    parts = re.split(r'(<u>|</u>|<b>|</b>|<r>|</r>|<y>|</y>|<bl>|</bl>)', clean_text)
    for part in parts:
        if part == '<u>': set_style(hwp, underline=True)
        elif part == '</u>': set_style(hwp, underline=False)
        elif part == '<b>': set_style(hwp, bold=True)
        elif part == '</b>': set_style(hwp, bold=False)
        elif part == '<r>': set_style(hwp, color=255)
        elif part == '</r>': set_style(hwp, color=0)
        elif part == '<y>': set_style(hwp, shadecolor=13434879)
        elif part == '</y>': set_style(hwp, shadecolor=4294967295)
        elif part == '<bl>': set_style(hwp, color=16711680)
        elif part == '</bl>': set_style(hwp, color=0)
        elif part: insert_text(hwp, part.replace('\n', '\r\n'))

def insert_keep_style(hwp, field_name, text):
    text_str = str(text).replace('\\n', '\n')
    if not text_str or text_str.strip().lower() == "null":
        hwp.PutFieldText(field_name, " ")
        return
    if not re.search(r'(<u>|</u>|<b>|</b>|<r>|</r>|<y>|</y>|<bl>|</bl>)', text_str):
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
            hwp.MoveToField(target, True, False, True)
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
            if len(parts) == 2: key_variations.add(parts[0].lower() + parts[1].capitalize())
        if k_lower in ['n', 'no', 'num', 'number']: key_variations.update(['n', 'N', 'No', 'NO', 'no', 'num', 'Num', 'NUM'])
        if k_lower in ['ans_tf', 'anstf']: key_variations.update(['ans_TF', 'ansTF', 'ANS_TF', 'TFA', 'ans_Tf'])
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
                            if act.GetDefault(act.CreateSet()): hwp.Run("TableDeleteRow")
                            else: hwp.PutFieldText(target, " ")
                        except: hwp.PutFieldText(target, " ")

    for prefix in ['w', 'W', 's', 'S', 'v', 'V']:
        for j in range(1, 31):
            val = str(content.get(f"{prefix}{j}", "")).strip()
            if not val or val.lower() == "null":
                for target in [f"{prefix}{j}"] + [f"{prefix}{j}{{{i}}}" for i in range(1, 20)]:
                    try: hwp.PutFieldText(target, " ")
                    except: pass

    passage_no = ""
    for k in ["n", "N", "No", "NO", "num", "Num", "NUM"]:
        if content.get(k):
            passage_no = str(content.get(k)).strip()
            break

    if passage_no:
        clean_no = re.sub(r'[^0-9]', '', passage_no)
        if not clean_no:
            clean_no = re.sub(r'<[^>]+>', '', passage_no).replace('⚠️', '').strip()

        image_path = None
        if os.path.exists(BASE_DIR):
            for f in os.listdir(BASE_DIR):
                f_name, f_ext = os.path.splitext(f)
                f_clean_name = re.sub(r'[^0-9]', '', f_name)
                if (f_clean_name == clean_no or f_name == clean_no) and f_ext.lower() in [".jpg", ".jpeg", ".png"]:
                    image_path = os.path.join(BASE_DIR, f)
                    break

        if image_path:
            for target in ["pic", "PIC"] + [f"pic{{{i}}}" for i in range(1, 10)]:
                if hwp.MoveToField(target, True, False, False):
                    hwp.PutFieldText(target, "")
                    time.sleep(0.1)
                    hwp.MoveToField(target, True, False, False)
                    try: hwp.InsertPicture(image_path, True, 3, False, False, 0)
                    except: pass
                    hwp.Run("Cancel")

def run_auto_qa_pipeline(content, template_name):
    logs = []
    if 'ins' in content and 're' in content:
        ins_t = str(content['ins']).strip()
        re_t = str(content['re']).strip()
        if ins_t and ins_t in re_t:
            content['re'] = re_t.replace(ins_t, "").replace("  ", " ")
            logs.append("└ [자체치유] 본문(re) 내 중복 문장(ins)을 삭제했습니다.")

    if 'ans_gr' in content:
        ans_t = str(content['ans_gr'])
        a_count = len(re.findall(r'\(A\)|앞', ans_t, re.IGNORECASE))
        b_count = len(re.findall(r'\(B\)|뒤', ans_t, re.IGNORECASE))
        if (a_count >= 7 or b_count >= 7) and (a_count + b_count > 0):
            logs.append("└ [보안경고] 정답 배열 쏠림 감지! 문항 번호에 붉은색 경고 표시를 추가합니다.")
            if 'n' in content:
                content['n'] = f"<r>⚠️ {content['n']}</r>"
    return content, logs

def parse_passages(text):
    # 번호 라벨(<20번>, [21], (22), 35., B13 등) 기준으로 (번호, 본문) 분리.
    # 같은 번호가 2번 이상 나오면(영어 따로/한글 따로 입력) 하나의 지문으로 병합한다.
    num_pat = re.compile(r'^[ \t]*[\<\(\[]?\s*([A-Za-z]?\d{1,3})\s*[\>\)\]\.번]*[ \t]*$')
    order = []
    bodies = {}
    cur_num, cur_body = None, []

    def flush(n, body_lines):
        if n is None:
            return
        b = '\n'.join(body_lines).strip()
        if not b:
            return
        if n not in bodies:
            bodies[n] = []
            order.append(n)
        bodies[n].append(b)

    for line in text.split('\n'):
        m = num_pat.match(line)
        if m:
            flush(cur_num, cur_body)
            cur_num = m.group(1)
            cur_body = []
        else:
            if cur_num is not None:
                cur_body.append(line)
    flush(cur_num, cur_body)

    return [(n, '\n\n'.join(bodies[n])) for n in order]

def ai_generate_text(model, payload, timeout=AI_TIMEOUT):
    resp = model.generate_content(payload, stream=True, request_options={"timeout": timeout})
    parts = []
    for chunk in resp:
        try:
            if chunk.text:
                parts.append(chunk.text)
        except Exception:
            pass
    return "".join(parts)

class NeoEasternMasterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Neo Eastern AI 마스터 출제 프로그램 v2.5")
        self.geometry("1050x750")
        ctk.set_appearance_mode("light")
        self.hwp_path = ""

        self.auth_frame = ctk.CTkFrame(self)
        self.auth_frame.pack(side="top", pady=10, padx=20, fill="x")

        ctk.CTkLabel(self.auth_frame, text="라이선스 키:").pack(side="left", padx=10, pady=10)
        self.license_entry = ctk.CTkEntry(self.auth_frame, width=180)
        self.license_entry.pack(side="left", padx=5)

        ctk.CTkLabel(self.auth_frame, text="Gemini API 키:").pack(side="left", padx=10)
        self.api_entry = ctk.CTkEntry(self.auth_frame, width=280, show="*")
        self.api_entry.pack(side="left", padx=5)

        self.generate_btn = ctk.CTkButton(self, text="🚀 확정 템플릿 기반 문서 자동 생성 시작", height=55, font=("Malgun Gothic", 16, "bold"), command=self.start_pipeline_thread)
        self.generate_btn.pack(side="bottom", pady=15, padx=20, fill="x")

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(side="top", pady=(0, 5), padx=20, fill="both", expand=True)

        self.input_frame = ctk.CTkFrame(self.main_frame)
        self.input_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        self.file_btn = ctk.CTkButton(self.input_frame, text="📄 메모장(.txt) 파일 통째로 불러오기", command=self.load_txt_file)
        self.file_btn.pack(fill="x", pady=5)

        ctk.CTkLabel(self.input_frame, text="분석할 영어 원문 및 해석 직접 입력창:").pack(anchor="w", pady=2)
        self.text_input = ctk.CTkTextbox(self.input_frame)
        self.text_input.pack(fill="both", expand=True, pady=5)

        self.right_frame = ctk.CTkScrollableFrame(self.main_frame, width=340)
        self.right_frame.pack(side="right", fill="y", padx=10, pady=10)

        self.top_settings = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        self.top_settings.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(self.top_settings, text="📌 템플릿 종류 선택:").pack(anchor="w", pady=(5, 0))

        self.template_combo = ctk.CTkComboBox(self.top_settings, values=[
            "수업용", "클리닉 A - 지칭O", "클리닉 B - 지칭X", "워크북", "미니북", "변형문제", "어법어휘 선택", "고난도 어법", "단어장", "구문", "병맛용"
        ], width=260, command=self.on_template_change)
        self.template_combo.pack(fill="x", pady=5)

        self.variation_frame = ctk.CTkFrame(self.right_frame, border_width=1, border_color="#3b82f6", fg_color="#eff6ff")

        type_lbl_frame = ctk.CTkFrame(self.variation_frame, fg_color="transparent")
        type_lbl_frame.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(type_lbl_frame, text="✅ 타겟 문항 유형 (다중 선택)", font=("", 12, "bold"), text_color="#1e3a8a").pack(side="left")
        ctk.CTkButton(type_lbl_frame, text="모두 해제", width=60, height=20, font=("", 10), command=self.uncheck_all_types).pack(side="right", padx=2)
        ctk.CTkButton(type_lbl_frame, text="전체 선택", width=60, height=20, font=("", 10), command=self.check_all_types).pack(side="right", padx=2)

        self.checkbox_vars = {}
        self.type_grid_frame = ctk.CTkFrame(self.variation_frame, fg_color="transparent")
        self.type_grid_frame.pack(fill="x", padx=10, pady=5)

        types_list = ["요지", "주제", "주장", "내용일치", "어법", "어휘", "빈칸", "빈칸패러", "순서", "삽입", "요약문", "서술형"]
        for i, t_name in enumerate(types_list):
            var = ctk.BooleanVar(value=False)
            self.checkbox_vars[t_name] = var
            cb = ctk.CTkCheckBox(self.type_grid_frame, text=t_name, variable=var, width=80)
            cb.grid(row=i//4, column=i%4, padx=2, pady=5, sticky="w")

        ctk.CTkLabel(self.variation_frame, text="🔥 3단 정밀 난이도 제어", font=("", 12, "bold"), text_color="#1e3a8a").pack(anchor="w", padx=10, pady=(10, 0))

        ctk.CTkLabel(self.variation_frame, text="1. 객관식 선지 수준:", text_color="#333333").pack(anchor="w", padx=10)
        self.diff1_combo = ctk.CTkComboBox(self.variation_frame, values=[
            "[A] 기본 (교과서 수준 평이한 어휘)",
            "[B] 중간 (모의고사 기출 수준 유의어/반의어)",
            "[C] 심화 (수능/고3 수준 고난도/추상적 어휘)"
        ], width=280)
        self.diff1_combo.pack(padx=10, pady=(0, 5))

        ctk.CTkLabel(self.variation_frame, text="2. 어휘 문제 선지 변형:", text_color="#333333").pack(anchor="w", padx=10)
        self.diff2_combo = ctk.CTkComboBox(self.variation_frame, values=[
            "[A] 원문 단어 유지 (지문 암기 확인용)",
            "[B] 혼동어/동반의어로 100% 변형 (변별력 강화)"
        ], width=280)
        self.diff2_combo.pack(padx=10, pady=(0, 5))

        ctk.CTkLabel(self.variation_frame, text="3. 빈칸패러 정답 변형:", text_color="#333333").pack(anchor="w", padx=10)
        self.diff3_combo = ctk.CTkComboBox(self.variation_frame, values=[
            "[A] 중간 (내신 수준 유의어 및 평이한 구문)",
            "[B] 심화 (완전 다른 추상적 구문 및 고난도 단어)"
        ], width=280)
        self.diff3_combo.pack(padx=10, pady=(0, 10))

        self.bottom_settings = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        self.bottom_settings.pack(fill="both", expand=True, pady=(5, 0))

        ctk.CTkLabel(self.bottom_settings, text="📂 연동할 한글(HWP) 파일:").pack(anchor="w", pady=(5, 0))
        self.hwp_btn = ctk.CTkButton(self.bottom_settings, text="템플릿 파일 선택", command=self.select_hwp_template, fg_color="#2c3e50")
        self.hwp_btn.pack(fill="x", pady=5)
        self.hwp_label = ctk.CTkLabel(self.bottom_settings, text="선택된 파일 없음", text_color="gray")
        self.hwp_label.pack(fill="x", pady=2)

        ctk.CTkLabel(self.bottom_settings, text="📊 실시간 시스템 분석 로그:").pack(anchor="w", pady=(10, 0))
        self.log_textbox = ctk.CTkTextbox(self.bottom_settings, height=120, width=280)
        self.log_textbox.pack(fill="both", expand=True, pady=5)

        self.on_template_change(self.template_combo.get())

    def on_template_change(self, choice):
        if choice == "변형문제":
            self.variation_frame.pack(fill="x", pady=5, after=self.top_settings)
        else:
            self.variation_frame.pack_forget()

    def check_all_types(self):
        for var in self.checkbox_vars.values(): var.set(True)

    def uncheck_all_types(self):
        for var in self.checkbox_vars.values(): var.set(False)

    def log(self, message):
        self.after(0, self._append_log, message)

    def _append_log(self, message):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")

    def load_txt_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, "r", encoding="utf-8") as f: content = f.read()
            self.text_input.delete("1.0", "end")
            self.text_input.insert("1.0", content)
            self.log(f"> 메모장 로드 완료: {os.path.basename(file_path)}")

    def select_hwp_template(self):
        file_path = filedialog.askopenfilename(filetypes=[("HWP Files", "*.hwp")])
        if file_path:
            self.hwp_path = file_path
            file_name = os.path.basename(file_path)
            self.hwp_label.configure(text=f"선택됨: {file_name}", text_color="blue")
            self.log(f"> 한글 템플릿 파일 선택 완료: {file_name}")

    def clean_json_string(self, text):
        t = text.strip()
        if t.startswith("```json"):
            t = t[7:]
        elif t.startswith("```"):
            t = t[3:]
        if t.endswith("```"):
            t = t[:-3]
        return re.sub(r'//.*', '', t).strip()

    def start_pipeline_thread(self):
        license_key = self.license_entry.get().strip()
        api_key = self.api_entry.get().strip()
        selected_template = self.template_combo.get()
        input_text = self.text_input.get("1.0", "end-1c").strip()

        if not all([license_key, api_key, self.hwp_path, input_text]):
            messagebox.showerror("입력 누락", "모든 텍스트 필드를 입력하고 한글 템플릿 파일을 선택하십시오.")
            return

        selected_types = []
        diff1, diff2, diff3 = "", "", ""
        if selected_template == "변형문제":
            selected_types = [t_name for t_name, var in self.checkbox_vars.items() if var.get()]
            if not selected_types:
                messagebox.showerror("옵션 누락", "변형문제 출제 시 최소 1개 이상의 타겟 문항 유형을 체크해 주십시오.")
                return
            diff1, diff2, diff3 = self.diff1_combo.get(), self.diff2_combo.get(), self.diff3_combo.get()

        self.generate_btn.configure(state="disabled", text="⚡ 전역 텍스트 AI 연산 진행 중...")
        work_thread = threading.Thread(target=self.process_master_pipeline, args=(license_key, api_key, selected_template, input_text, selected_types, diff1, diff2, diff3), daemon=True)
        work_thread.start()

    def process_master_pipeline(self, license_key, api_key, selected_template, input_text, selected_types, diff1, diff2, diff3):
        self.log(f"\n[시작] {selected_template} 전역 분석 시퀀스 가동")
        try:
            gas_request_url = f"{GAS_URL}?key={license_key}&template={selected_template}"
            response = requests.get(gas_request_url, timeout=10).json()

            if response.get("status") != "SUCCESS":
                print(f"\n[디버그] 구글 시트 거절 사유: {response}\n")
                self.after(0, messagebox.showerror, "인증 에러", "라이선스 정보가 일치하지 않거나 만료되었습니다.")
                self.after(0, self.reset_btn)
                return

            sys_prompt = response.get("system_prompt", "")

            genai.configure(api_key=api_key)
            main_model = genai.GenerativeModel('gemini-3.1-pro-preview', system_instruction=sys_prompt, generation_config={"temperature": 0.3})
            retry_model = genai.GenerativeModel('gemini-3.5-flash', system_instruction=sys_prompt, generation_config={"temperature": 0.3})

            chunk_limit = 4 if selected_template in ["워크북", "미니북"] else 6

            parsed = parse_passages(input_text)
            if parsed:
                self.log(f"> 지문 {len(parsed)}개 분리 완료: {[n for n, _ in parsed]}")
            else:
                raw = [p.strip() for p in input_text.split('\n\n') if p.strip()]
                parsed = [("", body) for body in raw]
                self.log(f"> 번호 미검출 → 빈 줄 기준 {len(parsed)}개로 분할")

            payloads = []
            global_q_num = 1

            if selected_template == "변형문제":
                for q_type in selected_types:
                    override_header = (
                        f"[시스템 강제 오더]\n"
                        f"- 출제 타겟 문항 유형: 무조건 '{q_type}' 1가지 유형으로만 출제할 것!\n"
                        f"- 객관식 선지 수준: {diff1}\n"
                        f"- 어휘 문제 선지 변형 여부: {diff2}\n"
                        f"- 빈칸패러 정답 변형 난이도: {diff3}\n"
                        f"* 지시사항: 다른 유형은 섞지 말고, 오직 지정된 '{q_type}' 유형으로만 출제하며, 3단 난이도 옵션을 100% 엄수할 것.\n"
                        f"* 대화형 텍스트를 절대 출력하지 말고 오직 단일 JSON 배열만 뱉을 것.\n\n"
                    )
                    chunks = [parsed[i:i + chunk_limit] for i in range(0, len(parsed), chunk_limit)]
                    for chunk in chunks:
                        nums = [n for n, _ in chunk]
                        bodies = "\n\n".join((f"{n}\n{b}" if n else b) for n, b in chunk)
                        end_q_num = global_q_num + len(chunk) - 1
                        additional_instructions = (
                            f"[추가 지시사항]\n"
                            f"1. 🚨아래 [출제 대상 지문]에 들어 있는 지문 {len(chunk)}개에 대해서만, 각 지문당 1문제씩 출제하십시오.\n"
                            f"2. 🚨[매우 중요] 문항 일련번호는 무조건 {global_q_num}번부터 {end_q_num}번까지 순서대로 부여하십시오. "
                            f"(반드시 딱 {len(chunk)}개의 문제만 생성하고 즉시 종료할 것!)\n\n"
                        )
                        global_q_num += len(chunk)
                        payload = override_header + additional_instructions + f"[출제 대상 지문]\n{bodies}"
                        payloads.append((f"[{q_type}] 번호 {nums}", payload))
            else:
                override_header = (
                    f"[시스템 강제 오더]\n"
                    f"* 대화형 텍스트(안내 멘트 등)를 절대 출력하지 말고 오직 단일 JSON 배열만 뱉을 것.\n\n"
                )
                chunks = [parsed[i:i + chunk_limit] for i in range(0, len(parsed), chunk_limit)]
                for chunk in chunks:
                    nums = [n for n, _ in chunk]
                    bodies = "\n\n".join((f"{n}\n{b}" if n else b) for n, b in chunk)
                    additional_instructions = (
                        f"[추가 지시사항]\n"
                        f"아래 [출제 대상 지문]의 지문들에 대해서만 JSON 배열을 생성하십시오.\n\n"
                    )
                    payload = override_header + additional_instructions + f"[출제 대상 지문]\n{bodies}"
                    payloads.append((f"번호 {nums}", payload))

            final_json_pool = []
            for idx, (label, chunk_payload) in enumerate(payloads):
                self.log(f"-> AI 연산 가동 중: {label} [{idx+1}/{len(payloads)}]")
                raw_text = None

                # 메인 모델: 스트리밍 + (리밋/데드라인/일시오류) 시 재시도. Flash로 떨어지기 전에 pro를 먼저 다시 시도.
                for attempt in range(3):
                    try:
                        raw_text = ai_generate_text(main_model, chunk_payload, timeout=AI_TIMEOUT)
                        if raw_text and raw_text.strip():
                            break
                        raise RuntimeError("빈 응답")
                    except Exception as ex:
                        msg = str(ex).lower()
                        retryable = any(k in msg for k in [
                            '429', 'quota', 'exhausted', 'resource', 'rate',
                            'deadline', '504', 'timeout', 'unavailable', '503', '빈 응답'
                        ])
                        if retryable and attempt < 2:
                            wait = 20 * (attempt + 1)
                            self.log(f"⏳ 지연/리밋 감지({type(ex).__name__}) → {wait}초 후 메인 재시도 ({attempt+1}/2)")
                            time.sleep(wait)
                            continue
                        self.log(f"❌ 메인 모델 호출 실패: {type(ex).__name__} - {ex}")
                        raw_text = None
                        break

                try:
                    if not raw_text:
                        raise RuntimeError("메인 모델 응답 없음")
                    self.log("   ↳ 응답 수신 완료, JSON 파싱 중...")
                    cleaned_json = self.clean_json_string(raw_text)
                    parsed_batch = json.loads(cleaned_json)
                    if isinstance(parsed_batch, dict): parsed_batch = [parsed_batch]

                    for item in parsed_batch:
                        healed_item, qa_logs = run_auto_qa_pipeline(item, selected_template)
                        for q_log in qa_logs: self.log(q_log)
                        final_json_pool.append(healed_item)
                except Exception as ex:
                    self.log(f"❌ 1차 파싱 실패. Flash 복구 시도... ({type(ex).__name__})")
                    if raw_text:
                        print(f"\n========== [DEBUG: 1차 실패 RAW DATA] ==========\n{raw_text}\n================================================\n")

                    try:
                        retry_text = ai_generate_text(
                            retry_model,
                            f"정확한 단일 JSON 배열 구조로 스키마를 완전 수정하여 다시 출력하라:\n{chunk_payload}",
                            timeout=AI_TIMEOUT
                        )
                        cleaned_retry_json = self.clean_json_string(retry_text)
                        parsed_batch = json.loads(cleaned_retry_json)
                        if isinstance(parsed_batch, dict): parsed_batch = [parsed_batch]
                        final_json_pool.extend(parsed_batch)
                        self.log("✅ 복구 성공!")
                    except Exception as e:
                        self.log(f"🔺 통신 에러: {e}")
                        if 'retry_text' in locals() and retry_text:
                            print(f"\n========== [DEBUG: 복구 실패 RAW DATA] ==========\n{retry_text}\n=================================================\n")

                time.sleep(5)

            if not final_json_pool:
                self.after(0, messagebox.showerror, "데이터 오류", "유효 데이터가 없습니다. 터미널 창(검은 화면)의 디버깅 로그를 확인해 주십시오.")
                self.after(0, self.reset_btn)
                return

            if selected_template == "변형문제":
                self.log("> 변형문제 감지: 쪼개진 문항과 해설을 하나로 통합 병합 중...")
                combined_questions = []
                combined_answers = []

                for data_chunk in final_json_pool:
                    if "1" in data_chunk and str(data_chunk["1"]).strip() not in ["", "null"]:
                        combined_questions.append(str(data_chunk["1"]).strip())
                    if "a1" in data_chunk and str(data_chunk["a1"]).strip() not in ["", "null"]:
                        combined_answers.append(str(data_chunk["a1"]).strip())

                merged_data = {}
                if combined_questions:
                    merged_data["1"] = "\n\n\n".join(combined_questions)
                if combined_answers:
                    merged_data["a1"] = "\n\n\n".join(combined_answers)

                final_json_pool = [merged_data]
                self.log("✅ 전체 문항/해설 통합 완료! 단일 문서로 조립을 시작합니다.")

            hwp = init_hwp()
            if not hwp:
                self.after(0, messagebox.showerror, "HWP 연결 실패", "한글을 열 수 없습니다.")
                self.after(0, self.reset_btn)
                return

            if os.path.exists(TEMP_DIR): shutil.rmtree(TEMP_DIR, ignore_errors=True)
            os.makedirs(TEMP_DIR, exist_ok=True)

            for i, data_object in enumerate(final_json_pool):
                hwp.Open(self.hwp_path)
                time.sleep(0.3)
                process_fields_and_rows(hwp, data_object)
                temp_output_path = os.path.join(TEMP_DIR, f"temp_{i:02d}.hwp")
                hwp.SaveAs(temp_output_path)
                hwp.Clear(1)
                time.sleep(0.2)

            temp_files = sorted([os.path.join(TEMP_DIR, f) for f in os.listdir(TEMP_DIR) if f.endswith(".hwp")])

            if selected_template == "변형문제" and selected_types:
                types_str = " ".join(selected_types)
                output_filename = f"{selected_template}_{types_str}_done.hwp"
            else:
                output_filename = f"{selected_template}_done.hwp"

            if temp_files:
                hwp.Open(temp_files[0])
                time.sleep(0.5)
                for f_path in temp_files[1:]:
                    hwp.HAction.Run("Cancel")
                    hwp.HAction.Run("MoveDocEnd")
                    hwp.HAction.Run("MoveRight")
                    hwp.HAction.Run("MoveRight")
                    hwp.HAction.Run("MoveDocEnd")
                    hwp.HAction.Run("BreakSection")
                    time.sleep(0.1)

                    act = hwp.CreateAction("InsertFile")
                    pset = act.CreateSet()
                    act.GetDefault(pset)
                    pset.SetItem("FileName", f_path)
                    pset.SetItem("KeepSection", 1)
                    act.Execute(pset)
                    time.sleep(0.1)

                patterns = [(r"\[[^\]]*\]", True), (r"\([a-zA-Z]\)[ ]*_+", True), (r"\[[ ]*T[ ]*/[ ]*F[ ]*\]", False)]
                for regex, is_bold in patterns:
                    hwp.HAction.Run("MoveDocBegin")
                    find_ps = hwp.HParameterSet.HFindReplace
                    hwp.HAction.GetDefault("FindReplace", find_ps.HSet)
                    find_ps.HSet.SetItem("FindString", regex)
                    find_ps.HSet.SetItem("FindRegExp", 1)
                    find_ps.HSet.SetItem("IgnoreMessage", 1)
                    find_ps.HSet.SetItem("Direction", 0)
                    while hwp.HAction.Execute("RepeatFind", find_ps.HSet):
                        set_style(hwp, bold=is_bold)
                        hwp.HAction.Run("MoveRight")
                    hwp.Run("Cancel")

                hwp.SaveAs(os.path.join(BASE_DIR, output_filename))

            self.log(f"✅ 완전 무결성 컴파일 완료!")
            self.log(f"📂 저장 경로: {output_filename}")
            self.after(0, messagebox.showinfo, "컴파일 완료", f"[{output_filename}] 파일이 정상적으로 빌드되었습니다.")

        except Exception as e:
            self.log(f"❌ 런타임 오류: {e}")
            self.after(0, messagebox.showerror, "시스템 오류", f"오류 발생.\n{e}")
        finally:
            shutil.rmtree(TEMP_DIR, ignore_errors=True)
            self.after(0, self.reset_btn)

    def reset_btn(self):
        self.generate_btn.configure(state="normal", text="🚀 확정 템플릿 기반 문서 자동 생성 시작")

if __name__ == "__main__":
    app = NeoEasternMasterApp()
    app.mainloop()
