import re
import pandas as pd
def find_mismatch_details(key, df, col_map, xml, tag_counts, attr_counts):
    # --- DataFrame side ---
    if key == "p":
        problem_cols = ["p_id"] if "p_id" in df.columns else []
    else:
        problem_cols = list(set([cols[1] for cols in col_map if key in cols[1]]))
    
    df_contents = {}
    for col in problem_cols:
        df_contents[col] = df[col].dropna()

    # --- XML side ---
    xml_items = []
    if '_' in key:
        tag, attr = key.split('_', 1)
        for i, m in enumerate(re.finditer(
            rf'<{tag}[^>]*\b{attr}\s*=\s*["\']([^"\']*)["\']', xml
        )):
            line = xml[:m.start()].count('\n') + 1
            xml_items.append((i, m.group(1), line))
    else:
        if key == "p":
            matched_starts = set()
            for i, m in enumerate(re.finditer(
                r'<p>("(?:[^"]*)";\s*[\d.]+;\s*")', xml
            )):
                line = xml[:m.start()].count('\n') + 1
                raw = m.group(1)
                normalized = re.match(r'"([^"]+)"', raw).group(1)
                xml_items.append((i, normalized, line))
                matched_starts.add(m.start())
            # find <p> tags that didn't match the expected pattern
            for m in re.finditer(r'<p>', xml):
                if m.start() not in matched_starts:
                    line = xml[:m.start()].count('\n') + 1
                    snippet = xml[m.start():m.start()+120].replace('\n', ' ')
                    print(f"  [debug] unmatched <p> at line {line}: {snippet!r}")
        else:
            for i, m in enumerate(re.finditer(
                rf'<{key}[^>]*>(.*?)</{key}>', xml, re.DOTALL
            )):
                line = xml[:m.start()].count('\n') + 1
                xml_items.append((i, m.group(1), line))

    # --- DF items flattened and sorted by row index ---
    df_items = []
    for col, series in df_contents.items():
        for idx, val in series.items():
            df_items.append((idx, val))
    df_items.sort(key=lambda x: x[0])

    # --- Positional comparison ---
    print(f"\n=== Mismatch: {key} | XML={len(xml_items)} DF={len(df_items)} ===")
    mismatches = 0
    for pos in range(max(len(xml_items), len(df_items))):
        xml_i, xml_val, xml_line = xml_items[pos] if pos < len(xml_items) else (pos, "MISSING", "?")
        df_i,  df_val            = df_items[pos]   if pos < len(df_items)  else (pos, "MISSING")
        if xml_val != df_val:
            print(f"  pos {pos} | XML[{xml_i}] line {xml_line}: '{xml_val}' != DF[{df_i}]: '{df_val}'")
            mismatches += 1
    if mismatches == 0:
        print("  All positions match.")
def check_counts(df, xml_path, schema_path):
    # Read schema and extract tags
    with open(schema_path) as f:
        schema = f.read()
    
    tags = re.findall(r'<!ELEMENT\s+(\w+)', schema)
    tags.remove("text")
    
    # Read XML
    if xml_path is None:
    	return None
    with open(xml_path) as f:
        xml = f.read()
    
    # Count each tag with its attributes
    tag_counts = {}
    attr_counts = {}
    
    for tag in tags:
        matches = re.findall(rf'<{tag}([^>]*?)(?:>|/>)', xml)
        tag_counts[tag] = len(matches)
        
        for match in matches:
            attrs = re.findall(r'(\w+)\s*=\s*["\'][^"\']*["\']', match)
            for attr in attrs:
                key = f"{tag}_{attr}"
                attr_counts[key] = attr_counts.get(key, 0) + 1
    
    # Count from dataframe
    df_counts = {}
    col_map = []
    for col in df.columns:
    	
        col_ = col
        if "_text" in col:
            col_ = col.replace("_text","")
        
        if "p_id" in col_:
            col_ = "p"
        
        col_ = re.sub(r'_[0-9]|[0-9]_','',col_)
        if len(re.findall("_",col_)) >1:
            col_ = re.sub(r'^.*?_','',col_)
        if (re.findall(r'^.*?_EXPANSION|VDD|MOD|MD',col_)) and not (re.findall(r'VDD_type|EXPANSION_type|VDD_lemme',col_)):
            col_ = re.sub(r'(^.*?_)','',col_)
        if (not col_ == "paragraph") and (not col == "source_file"):
            if col_ in df_counts.keys():
                col_map.append((col_,col))
                df_counts[col_] += (len(df[col].dropna()))
            else:
                df_counts[col_] = (len(df[col].dropna()))
                col_map.append((col_,col))
    
    # Compare
    import os
    print("----------------------------------------")
    print("Vérification de la conformité des colonnes du fichier de sortie avec les balises du xml:")
    print("Traitement du fichier: ",os.path.basename(xml_path))
    merged = tag_counts.copy()
    for k, v in attr_counts.items():
        merged[k] = merged.get(k, 0) + v
    
    for key in set(df_counts) | set(merged):
        df_val = df_counts.get(key, 0)
        merged_val = merged.get(key, 0)
        if df_val != merged_val:
            
            print(f"Erreur: {key}: Excel={df_val}, xml={merged_val}")
            find_mismatch_details(key, df, col_map, xml, tag_counts, attr_counts)
        else:
            print(f"OK: {key}: Excel={df_val}, xml={merged_val}")
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
