
import zipfile
import re
import os
import json
import datetime
import sys

def get_docx_text(path):
    try:
        if not os.path.exists(path):
            return ""
        document = zipfile.ZipFile(path)
        xml_content = document.read('word/document.xml')
        document.close()
        xml_str = xml_content.decode('utf-8')
        xml_str = re.sub(r'</w:p>', '\n', xml_str) 
        text = re.sub(r'<[^>]+>', '', xml_str)
        return text
    except Exception as e:
        print(f"Error reading {path}: {e}")
        return ""

def parse_id_card(id_card):
    if not id_card or len(id_card) != 18:
        return None, None
    birth_str = id_card[6:14]
    birth_date = f"{birth_str[:4]}-{birth_str[4:6]}-{birth_str[6:]}"
    gender_code = int(id_card[16])
    gender = "1" if gender_code % 2 == 1 else "2"
    return gender, birth_date

def clean_name_field(val):
    if not val: return ""
    return re.sub(r'[^\u4e00-\u9fa5]', '', val).strip()

def clean_address_field(val):
    if not val: return ""
    return re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9\s]', '', val).strip()

def clean_marriage_field(val):
    if not val: return "已婚"
    if "未婚" in val or val == "1": return "未婚"
    if "已婚" in val or val == "2": return "已婚"
    if "丧偶" in val or val == "3": return "丧偶"
    if "离婚" in val or val == "4": return "离婚"
    if "其他" in val or val == "9": return "其他"
    return "已婚"

def clean_relationship_field(val):
    if not val: return "8"
    val = val.strip()
    if re.search(r'本人|户主', val): return "0"
    if re.search(r'配偶|妻|夫|爱人|老婆|老公', val): return "1"
    if re.search(r'子|儿|女', val) and not re.search(r'孙|甥|侄', val): 
        if re.search(r'婿|媳', val): return "8" 
        return "2" 
    if re.search(r'子|儿', val) and '女' not in val: return "2" 
    if '女' in val and '孙' not in val and '侄' not in val: return "3" 
    if re.search(r'孙|外孙', val): return "4"
    if re.search(r'父|母|爸|妈', val): return "5"
    if re.search(r'祖|奶|爷|姥', val): return "6"
    if re.search(r'兄|弟|姐|妹', val): return "7"
    return "8" 

def parse_record(text, filename=""):
    rec = {}
    text = text.replace('：', ':').replace('；', ';')
    
    def extract(pattern, default=""):
        m = re.search(pattern, text)
        return m.group(1).strip() if m else default
    
    def extract_with_stop(source_text, label, stops, max_len=None, default=""):
        if isinstance(stops, str): stops = [stops]
        stops_pattern = "|".join([re.escape(s) for s in stops])
        pattern = f"{label}\\s*[:;](.*?)(?=\\s*(?:{stops_pattern})|$)"
        m = re.search(pattern, source_text)
        if m:
            val = m.group(1).strip()
            if max_len and len(val) > max_len:
                for s in stops:
                    if s in val:
                        val = val.split(s)[0].strip()
                        break
            return val
        return default

    rec['field_1_1'] = extract(r'付款方式\s*[:;]?\s*(\d{1,2})\b', "02")
    rec['field_1_2'] = extract(r'第\s*(\d+)\s*次住院')
    rec['field_1_3'] = extract(r'病案号\s*[:;]\s*(\d+)')
    if not rec['field_1_3']:
        rec['field_1_3'] = extract(r'住院号\s*[:;]\s*\d+\s+(\d+)')
    raw_name = extract_with_stop(text, "姓名", stops=["性别", "11.", "21.", "1.", "2."])
    if raw_name == "Unknown" or not raw_name:
         raw_name = extract(r'姓名\s*[:;]\s*([\u4e00-\u9fa5]{2,4})') or "Unknown"
    rec['field_1_4'] = clean_name_field(raw_name)
    rec['field_1_11'] = extract(r'身份证号\s*[:;]\s*([0-9X]+)')
    gender_from_id, birth_from_id = parse_id_card(rec['field_1_11'])
    if gender_from_id:
        rec['field_1_5'] = gender_from_id
        rec['field_1_6'] = birth_from_id
    else:
        m_gender = re.search(r'性别\s*[:;].*?(1|2)\.[男女]', text)
        rec['field_1_5'] = m_gender.group(1) if m_gender else "1"
        rec['field_1_6'] = extract(r'出生日期\s*[:;]\s*(\d{4}-\d{2}-\d{2})')
    if rec['field_1_6']:
        try:
            born = int(rec['field_1_6'][:4])
            current = 2025
            rec['field_1_7'] = str(current - born)
        except:
             rec['field_1_7'] = extract(r'年龄\s*[:;]\s*(\d+)')
    else:
        rec['field_1_7'] = extract(r'年龄\s*[:;]\s*(\d+)')
    rec['field_1_8'] = "CHN"
    bp_clean = extract_with_stop(text, "出生地", stops=["籍贯", "民族"], max_len=30)
    rec['field_1_9'] = clean_address_field(bp_clean)
    native_clean = extract_with_stop(text, "籍贯", stops=["民族", "职业"], max_len=30)
    rec['field_1_9_native'] = clean_address_field(native_clean)
    rec['field_1_10'] = extract_with_stop(text, "民族", stops=["职业", "婚姻", "婚否"], max_len=10)
    if not rec['field_1_10']:
         m_eth = re.search(r'民族\s*[:;]?\s*([\u4e00-\u9fa5]+)', text)
         rec['field_1_10'] = m_eth.group(1) if m_eth else "汉族"
    val_job = extract_with_stop(text, "职业", stops=["婚姻", "婚否", "现住地址"])
    if not val_job:
        m_job = re.search(r'职业\s*[:]\s*([\u4e00-\u9fa5]+)', text)
        if m_job: val_job = m_job.group(1)
    rec['field_1_12'] = val_job or "农民"
    val_mar = ""
    if '①' in text and '婚姻' in text: val_mar = "1"
    elif '②' in text and '婚姻' in text: val_mar = "2"
    elif '③' in text and '婚姻' in text: val_mar = "3"
    else:
        if "未婚" in text: val_mar = "未婚"
        elif "已婚" in text: val_mar = "已婚"
        elif "丧偶" in text: val_mar = "丧偶"
        elif "离婚" in text: val_mar = "离婚"
        else: val_mar = "2" 
    rec['field_1_13'] = clean_marriage_field(val_mar)
    rec['field_1_14'] = clean_address_field(extract_with_stop(text, "现住地址", stops=["电话", "邮政"]))
    rec['field_1_15'] = extract(r'电话\s*[:]\s*(\d+)')
    rec['field_1_16'] = clean_address_field(extract_with_stop(text, "户口地址", stops=["邮政", "电话", "工作"])) or rec.get('field_1_14', "")
    rec['field_1_17'] = clean_name_field(extract_with_stop(text, "联系人姓名", stops=["关系", "地址"]))
    raw_relation = extract_with_stop(text, "关系", stops=["地址", "电话", "联系人"])
    rec['field_1_18'] = clean_relationship_field(raw_relation)
    if rec['field_1_17'] == rec['field_1_4']:
         rec['field_1_17'] = ""
    rec['field_1_21'] = "2"
    m_admit = re.search(r'入院日期\s*[:]\s*(\d{4}-\d{2}-\d{2})\s+(\d{2})', text)
    if m_admit:
        rec['field_1_22_1'] = m_admit.group(1)
        rec['field_1_22_2'] = m_admit.group(2)
    else:
        rec['field_1_22_1'] = ""
        rec['field_1_22_2'] = ""
    rec['field_1_23'] = "精神科"
    m_disch = re.search(r'出院日期\s*[:]\s*(\d{4}-\d{2}-\d{2})\s+(\d{2})', text)
    if m_disch:
        rec['field_1_24_1'] = m_disch.group(1)
        rec['field_1_24_2'] = m_disch.group(2)
    else:
        rec['field_1_24_1'] = ""
        rec['field_1_24_2'] = ""
    rec['field_1_25'] = "精神科"
    days_match = re.search(r'实.{0,5}住.{0,5}院[:：]?\D*?(\d+)', text)
    if days_match:
        rec['field_1_26'] = days_match.group(1)
    else:
        rec['field_1_26'] = ""
    m_diag = re.search(r'主要诊断\s+([\u4e00-\u9fa5，]+)\s+([A-Z0-9\.]+)', text)
    if m_diag:
        rec['field_1_27'] = m_diag.group(1)
        rec['field_1_28'] = m_diag.group(2)
    else:
        m_name = re.search(r'门[(（]急[)）]诊诊断名称\s*[:](.*?)\s', text)
        rec['field_1_27'] = m_name.group(1).strip() if m_name else ""
        rec['field_1_28'] = ""
    rec['field_1_29'] = "1"
    rec['field_2_1'] = clean_name_field(extract_with_stop(text, "科室主任", ["主[(（]副主[)）]任医师", "主任医师", "副主任医师"], default="张医生"))
    rec['field_2_2'] = clean_name_field(extract_with_stop(text, "主[(（]副主[)）]任医师", ["主治医师"], default="王医生"))
    rec['field_2_3'] = clean_name_field(extract_with_stop(text, "主治医师", ["住院医师"], default="李医生"))
    rec['field_2_4'] = clean_name_field(extract_with_stop(text, "住院医师", ["责任护士"], default="赵医生"))
    rec['field_2_5'] = clean_name_field(extract_with_stop(text, "责任护士", ["进修医师", "实习医师", "编码员"]))
    rec['field_2_6'] = "1"
    rec['field_2_7'] = clean_name_field(extract_with_stop(text, "质控医师", ["质控护士"], default="张医生"))
    rec['field_2_8'] = clean_name_field(extract_with_stop(text, "质控护士", ["质控日期"], default="刘护士"))
    rec['field_2_9'] = extract(r'质控日期\s*[:]?\s*(\d{4}-\d{2}-\d{2})') or rec.get('field_1_24_1', "")
    rec['field_2_10'] = "1"
    def extract_fee(pattern, default="0.00"):
        m = re.search(pattern, text)
        return m.group(1) if m else default
    rec['field_3_1'] = extract_fee(r'总费用\s*[:]\s*([\d\.]+)')
    rec['field_3_2'] = extract_fee(r'自付金额\s*[:]\s*([\d\.]+)')
    rec['field_3_3'] = extract_fee(r'\(2\)一般治疗操作费\s*[:]\s*([\d\.]+)')
    rec['field_3_4'] = extract_fee(r'\(3\)护理费\s*[:]\s*([\d\.]+)')
    rec['field_3_5'] = extract_fee(r'\(4\)其他费用\s*[:]\s*([\d\.]+)')
    rec['field_3_6'] = extract_fee(r'\(6\)实验室诊断费\s*[:]\s*([\d\.]+)')
    rec['field_3_7'] = extract_fee(r'\(7\)影像学诊断费\s*[:]\s*([\d\.]+)')
    rec['field_3_8'] = extract_fee(r'\(13\)西药费\s*[:]\s*([\d\.]+)')
    rec['field_3_9'] = extract_fee(r'\(14\)中成药费\s*[:]\s*([\d\.]+)')
    rec['field_3_10'] = extract_fee(r'\(15\)中草药费\s*[:]\s*([\d\.]+)')
    rec['field_3_11'] = extract_fee(r'其他费\s*[:]\s*([\d\.]+)')
    date_keys = ['field_1_6', 'field_1_22_1', 'field_1_24_1', 'field_2_9']
    for key, val in rec.items():
        if key not in date_keys and isinstance(val, str):
            cleaned = val.replace('-', '').replace('—', '').strip()
            rec[key] = cleaned
    return rec

def main():
    # 自动识别基础目录：如果是打包后的 EXE，则使用 EXE 所在目录
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.getcwd()
        
    # 定义病历存放的子文件夹
    doc_dir = os.path.join(base_dir, "病历文档")
    if not os.path.exists(doc_dir):
        os.makedirs(doc_dir)
        
    output_records = []
    
    print(f"正在扫描目录: {doc_dir}")
    for filename in os.listdir(doc_dir):
        if filename.endswith(".docx") and not filename.startswith("~$"):
            path = os.path.join(doc_dir, filename)
            print(f"正在读取: {filename}")
            text = get_docx_text(path)
            if not text: continue
            
            parts = re.split(r'(?=医疗机构(?!转入)|住\s*院\s*病\s*案\s*首\s*页)', text)
            for part in parts:
                if "住院" in part or "病案首页" in part or "医疗机构" in part:
                    rec = parse_record(part, filename)
                    if rec.get('field_1_4') and rec.get('field_1_4') != "Unknown":
                        if rec.get('field_1_3'):
                            output_records.append(rec)
    
    def sort_key(r):
        try:
            return int(r.get('field_1_3', 0))
        except:
            return 0
    output_records.sort(key=sort_key)
    
    print(f"解析成功: 找到 {len(output_records)} 条病历。")
    
    # if not output_records:
    #    return  <-- REMOVED: We want to proceed even if empty, to clear the HTML.

    html_path = os.path.join(base_dir, "medical_record_lite.html") 
    if not os.path.exists(html_path):
        print(f"错误：找不到网页文件！\n路径: {html_path}")
        return

    print("正在写入网页数据 (同步模式)...")
    try:
        with open(html_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # SYNC MODE: We overwrite the existing 'records' array with the new 'output_records'
        # This ensures the HTML only contains exactly what is in the current DOCX folder.
        pattern_rec = re.compile(r'(let\s+records\s*=\s*)(\[.*?\])(\s*;)', re.DOTALL)
        
        new_json = json.dumps(output_records, ensure_ascii=False, indent=4)
        
        if pattern_rec.search(content):
            # Replace the entire array
            content = pattern_rec.sub(r'\1' + new_json.replace('\\', '\\\\') + r'\3', content)
            
            ts = int(datetime.datetime.now().timestamp())
            content = re.sub(r"const\s+KEY\s*=\s*'med_records_lite_.*?'", f"const KEY = 'med_records_lite_v{ts}'", content)
            
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(content)
                
            print(f"同步成功！网页现包含 {len(output_records)} 条病历（与文件夹一致）。")
        else:
             print("错误：无法在网页中找到数据存储区域。")
             
    except Exception as e:
        print(f"写入失败: {e}")

if __name__ == "__main__":
    main()
