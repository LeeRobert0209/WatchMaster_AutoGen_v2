
import pandas as pd
import re
import os

# input_path = r'data/sample_input.xlsx'
input_path = r'data/南京顺序-维修师介绍.xlsx-11.21.xlsx'
output_path = r'data/维修师数据_清洗版.xlsx'

try:
    # Row 2 is header (0, 1, 2)
    df = pd.read_excel(input_path, header=2)
    
    # Filter rows where Name (Col 1) is not empty.
    # Col 1 name should be in df.columns[1]
    name_col = df.columns[1]
    df = df.dropna(subset=[name_col])

    new_data = []

    def clean_text_content(text):
        if not text:
            return ""
        # 1. Remove spaces around '+' specific patterns first
        text = re.sub(r'\s*\+\s*', '+', text)
        
        # 2. Aggressive spacing removal for Chinese context
        # Remove space between Chinese/Number/Symbol and Chinese/Number/Symbol
        # Keep space only if both sides are likely English letters
        # Pattern: lookbehind for non-ascii OR lookahead for non-ascii implies we can strip space?
        # Simpler: Remove space if it's flanked by (Chinese/Digit/Punctuation)
        
        # Strategy: matching whitespace that is bordered by at least one non-en-letter character
        # \u4e00-\u9fa5 is Chinese.
        # Let's verify if there are pure English sentences. Likely not.
        # "15 年" -> "15" is digit, "年" is Chinese. Remove.
        # "领域 15" -> "领域" Chinese, "15" digit. Remove.
        
        text = re.sub(r'(?<=[\u4e00-\u9fa5\d+％%])\s+(?=[\u4e00-\u9fa5\d+％%])', '', text)
        text = re.sub(r'(?<=[\u4e00-\u9fa5])\s+(?=[a-zA-Z])', '', text) # Chinese space English
        text = re.sub(r'(?=[a-zA-Z])\s+(?<=[\u4e00-\u9fa5])', '', text) # English space Chinese (lookbehind fixed)
        
        # 3. Fix Logic for specific user complaints
        # "近10 年" -> handled by rule above (0 space Year)
        # "4000枚 +" -> handled by rule 1
        
        return text

    for idx, row in df.iterrows():
        # Get raw values
        store = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ""
        # store = store.replace("某某品牌后缀", "").strip() # 可在此处添加特定品牌后缀清洗逻辑
        
        full_text = str(row.iloc[6]) if pd.notna(row.iloc[6]) else ""
        
        # Initialize
        name = "" 
        experience = ""
        monologue = ""
        titles = ["", "", ""]
        descriptions = ["", "", ""]
        
        # --- 1. Extract Monologue (and remove it from main text) ---
        monologue_marker = "匠人独白"
        split_pattern = f"{monologue_marker}[：:]"
        parts = re.split(split_pattern, full_text, maxsplit=1)
        
        if len(parts) > 1:
            main_content = parts[0]
            monologue = parts[1].strip()
            # Clean monologue
            monologue = clean_text_content(monologue)
            monologue = re.sub(r'[。！!.\s,，]+$', '', monologue)
        else:
            main_content = full_text

        # --- 2. Extract Name and Experience ---
        first_line_match = re.search(r"\*?\s*维修技师[：:]\s*(.*?)(?:\n|\s{2,}|匠龄|$)", main_content)
        
        if first_line_match:
            potential_name = first_line_match.group(1).strip()
            if len(potential_name) <= 5 and not re.search(r'[0-9年]', potential_name):
                name = potential_name

        if not name:
             simple_match = re.search(r"维修技师[：:]\s*(\S+)", main_content)
             if simple_match:
                 cand = simple_match.group(1)
                 cand = cand.split("匠龄")[0]
                 if len(cand) <= 4 and not re.search(r'\d', cand) and "年" not in cand:
                     name = cand

        # Extract Experience
        exp_match = re.search(r"匠龄[:：]\s*(.*?)(?:\n|\s{2,}|$)", main_content)
        if exp_match:
            raw_exp = exp_match.group(1).strip()
            if raw_exp:
                experience = raw_exp
        
        if experience:
             experience = experience.replace("匠龄：", "").replace("匠龄:", "").strip()

        if not experience:
             year_match = re.search(r"(近?\d+\+?\s*年\+?|\d+余年)", main_content)
             if year_match:
                 experience = year_match.group(1)

        # 4. Format Experience and Clean
        if experience:
            experience = clean_text_content(experience)
            experience = re.sub(r"(\d+)\+年", r"\1年+", experience)

        # --- 3. Extract 3 Description Blocks ---
        raw_segments = [s.strip() for s in main_content.split('*') if s.strip()]
        
        valid_segments = []
        for seg in raw_segments:
            if "维修技师" in seg or ("匠龄" in seg and len(seg) < 30): 
                continue
            if "：" in seg or ":" in seg:
                valid_segments.append(seg)
                
        for i in range(min(3, len(valid_segments))):
            seg = valid_segments[i]
            if "：" in seg:
                t, d = seg.split("：", 1)
            elif ":" in seg:
                t, d = seg.split(":", 1)
            else:
                t, d = seg, ""
            
            t = clean_text_content(t).strip()
            d = clean_text_content(d).strip()
            
            # Simple check for repeated words like "维修维修" (2 chars repeated)
            # Only replace if strictly adjacent
            d = re.sub(r'([\u4e00-\u9fa5]{2,})\1', r'\1', d) 

            # Fix typos if any known (generic approach)
            # E.g. "  " -> "" handled by clean_text_content
            
            if d and not d.endswith(('。', '！', '!', '.', '…')):
                d += "。"
            
            titles[i] = t
            descriptions[i] = d

        new_data.append({
            "姓名": name,
            "门店": store,
            "匠龄": experience,
            "标题1": titles[0],
            "描述1": descriptions[0],
            "标题2": titles[1],
            "描述2": descriptions[1],
            "标题3": titles[2],
            "描述3": descriptions[2],
            "匠人独白": monologue
        })

    new_df = pd.DataFrame(new_data)
    new_df.to_excel(output_path, index=False)
    print(f"Successfully processed {len(new_df)} rows.")
    print(new_df[['姓名', '门店', '匠龄']].head())
    
except Exception as e:
    print(f"Error: {e}")
