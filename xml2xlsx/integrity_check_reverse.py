"""
integrity_check_reverse.py — Verify XML integrity after Excel→XML reverse injection.
"""

import xml.etree.ElementTree as ET
from collections import defaultdict


def count_elements(root):
    counts = defaultdict(int)
    for elem in root.iter():
        if callable(elem.tag):  # skip comments
            continue
        counts[elem.tag] += 1
    return dict(counts)

def get_all_text(root):
    texts = set()
    for elem in root.iter():
        if callable(elem.tag):  # skip comments
            continue
        if elem.tag == 'p':
            continue
        text = ' '.join((elem.text or '').split())
        if text:
            texts.add(text)
    return texts

def check_reverse_integrity(original_xml_path, updated_xml_path, children_map):
    try:
        from lxml import etree
        orig_tree = etree.parse(original_xml_path)
        upd_tree = etree.parse(updated_xml_path)
    except ET.ParseError as e:
        return {
            'valid': False,
            'messages': [f"❌ Erreur de parsing XML: {e}"],
            'tag_changes': {},
            'new_tags': [],
            'deleted_tags': []
        }

    orig_root = orig_tree.getroot()
    upd_root = upd_tree.getroot()
    orig_counts = count_elements(orig_root)
    upd_counts = count_elements(upd_root)
    orig_text = get_all_text(orig_root)
    upd_text = get_all_text(upd_root)

    messages = []
    tag_changes = {}
    valid = True

    # ── Check original tags still exist ────────────────────────────────────
    for tag, orig_count in orig_counts.items():
        upd_count = upd_counts.get(tag, 0)
        if upd_count == 0:
            messages.append(f"❌ Tag '{tag}' supprimé ({orig_count} → 0)")
            tag_changes[tag] = {'original': orig_count, 'updated': upd_count, 'status': 'MISSING'}
            valid = False
        elif upd_count < orig_count:
            messages.append(f"⚠️  Tag '{tag}' réduit ({orig_count} → {upd_count})")
            tag_changes[tag] = {'original': orig_count, 'updated': upd_count, 'status': 'CHANGED'}
            valid = False
        elif upd_count > orig_count:
            messages.append(f"✓ Tag '{tag}' augmenté ({orig_count} → {upd_count})")
            tag_changes[tag] = {'original': orig_count, 'updated': upd_count, 'status': 'OK'}
        else:
            messages.append(f"✓ Tag '{tag}': {orig_count} (inchangé)")
            tag_changes[tag] = {'original': orig_count, 'updated': upd_count, 'status': 'OK'}

    # ── Check for completely new tags ──────────────────────────────────────
    new_tags = []
    for tag, upd_count in upd_counts.items():
        if tag not in orig_counts:
            new_tags.append(tag)
            messages.append(f"➕ Nouveau tag '{tag}' ({upd_count} instance(s))")
            tag_changes[tag] = {'original': 0, 'updated': upd_count, 'status': 'NEW'}

    # ── Check for text deletion ────────────────────────────────────────────
    def normalize_text(t):
        return ' '.join(t.replace('\u2019', "'").replace('\u2018', "'").split())

    full_xml_text = normalize_text(''.join(
    ''.join(p.itertext())
    for p in upd_root.iter('p')
))

    deleted_text = set()
    for text in orig_text - upd_text:
        normalized = normalize_text(text)
        idx = full_xml_text.find(normalized[:40])
        if not (normalized in full_xml_text):
            idx2 = full_xml_text.find(normalized[:40])
            if idx2 >= 0:
                for i, (a, b) in enumerate(zip(normalized, full_xml_text[idx2:idx2+len(normalized)])):
                    if a != b:
                        break
            deleted_text.add(text)

    if deleted_text:
        messages.append(f"⚠️  {len(deleted_text)} élément(s) de texte supprimé(s):")
        for text in list(deleted_text)[:5]:
            truncated = text[:50] + "..." if len(text) > 50 else text
            messages.append(f"    - \"{truncated}\"")
        if len(deleted_text) > 5:
            messages.append(f"    ... et {len(deleted_text) - 5} autres")
        valid = False

    # ── Summary ───────────────────────────────────────────────────────────
    if valid and not new_tags:
        messages.insert(0, "✅ Intégrité confirmée: aucune perte de données")
    elif valid and new_tags:
        messages.insert(0, f"✅ OK: {len(new_tags)} nouveau(x) tag(s) ajouté(s), aucune perte")
    else:
        messages.insert(0, "❌ Problèmes détectés dans l'intégrité")

    return {
        'valid': valid,
        'messages': messages,
        'tag_changes': tag_changes,
        'new_tags': new_tags,
        'deleted_tags': list(deleted_text)[:10] if deleted_text else []
    }
