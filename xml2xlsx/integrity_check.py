import re
import pandas as pd
import os


def find_mismatch_details(key, df, col_map, xml, tag_counts, attr_counts):
    """Generate mismatch detail messages instead of printing."""
    messages = []
    
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
                    messages.append(f"  [debug] unmatched <p> at line {line}: {snippet!r}")
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
    messages.append(f"\n=== Mismatch: {key} | XML={len(xml_items)} DF={len(df_items)} ===")
    mismatches = 0
    for pos in range(max(len(xml_items), len(df_items))):
        xml_i, xml_val, xml_line = xml_items[pos] if pos < len(xml_items) else (pos, "MISSING", "?")
        df_i,  df_val            = df_items[pos]   if pos < len(df_items)  else (pos, "MISSING")
        if xml_val != df_val:
            messages.append(f"  pos {pos} | XML[{xml_i}] line {xml_line}: '{xml_val}' != DF[{df_i}]: '{df_val}'")
            mismatches += 1
    if mismatches == 0:
        messages.append("  All positions match.")
    
    return messages


def check_counts(df, xml_path, schema_path):
    """
    Verify XML and DataFrame integrity.
    
    Returns: list of message strings (can be empty if all OK)
    """
    messages = []
    
    # Read schema and extract tags
    try:
        with open(schema_path,encoding="utf-8") as f:
            schema = f.read()
    except (FileNotFoundError, TypeError):
        # schema_path is None or file doesn't exist
        return messages
    
    tags = re.findall(r'<!ELEMENT\s+(\w+)', schema)
    if "text" in tags:
        tags.remove("text")
    
    # Read XML
    if xml_path is None:
        return messages
    
    try:
        with open(xml_path,encoding="utf-8") as f:
            xml = f.read()
    except FileNotFoundError:
        messages.append(f"Erreur: fichier XML non trouvé: {xml_path}")
        return messages
    
    # Count each tag with its attributes
    tag_counts = {}
    attr_counts = {}
    
    for tag in tags:
        matches = re.findall(rf'<{tag}([^>]*?)(?:>|/>)', xml)
        tag_counts[tag] = len(matches)
        
        for match in matches:
	        # Capture les attributs avec valeur NON-VIDE
        	attrs = re.findall(r'(\w+)\s*=\s*["\']([^"\']+)["\']', match)
        	for attr, value in attrs:
        		if value:  # Seulement si la valeur n'est pas vide
            			key = f"{tag}_{attr}"
            			attr_counts[key] = attr_counts.get(key, 0) + 1
    
    # Count from dataframe
    df_counts = {}
    col_map = []
    for col in df.columns:
        col_ = col
        if "_text" in col:
            col_ = col.replace("_text", "")
        
        if "p_id" in col_:
            col_ = "p"
        
        col_ = re.sub(r'_[0-9]|[0-9]_', '', col_)
        if len(re.findall("_", col_)) > 1:
            col_ = re.sub(r'^.*?_', '', col_)
        if (re.findall(r'^.*?_EXPANSION|VDD|MOD|MD', col_)) and not (re.findall(r'VDD_type|EXPANSION_type|VDD_lemme', col_)):
            col_ = re.sub(r'(^.*?_)', '', col_)
        if (not col_ == "paragraph") and (not col == "source_file") and (not col == "xml_comments") and not (col == 'p_index'):
            if col_ in df_counts.keys():
                col_map.append((col_, col))
                df_counts[col_] += (len(df[col].dropna()))
            else:
                df_counts[col_] = (len(df[col].dropna()))
                col_map.append((col_, col))
    
    # Compare and build messages
    messages.append("----------------------------------------")
    messages.append("Vérification de la conformité des colonnes du fichier de sortie avec les balises du xml:")
    messages.append(f"Traitement du fichier: {os.path.basename(xml_path)}")
    errors = 0
    merged = tag_counts.copy()
    for k, v in attr_counts.items():
        merged[k] = merged.get(k, 0) + v
    
    for key in set(df_counts) | set(merged):
        df_val = df_counts.get(key, 0)
        merged_val = merged.get(key, 0)
        if df_val != merged_val:
            messages.append(f"❌ Erreur: {key}: Excel={df_val}, xml={merged_val}")
            mismatch_details = find_mismatch_details(key, df, col_map, xml, tag_counts, attr_counts)
            messages.extend(mismatch_details)
            errors +=1
        else:
            messages.append(f"✓ OK: {key}: Excel={df_val}, xml={merged_val}")
    if errors >0:
    	messages.append(f"{errors} ont été détectées")
    return messages
