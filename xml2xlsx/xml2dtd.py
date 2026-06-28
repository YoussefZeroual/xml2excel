#!/usr/bin/env python3
"""
xml2dtd.py — Déduit une DTD externe à partir d'un fichier XML.
Usage : python xml2dtd.py input.xml [output.dtd]
"""

import sys
from collections import defaultdict
from lxml import etree


def local(tag):
    """Supprime le namespace d'un tag lxml."""
    if not isinstance(tag, str):
        return None
    return tag.split("}", 1)[-1] if "}" in tag else tag


def analyze(tree):
    """
    Parcourt l'arbre XML et collecte :
    - les enfants possibles de chaque élément
    - les attributs de chaque élément
    - si un élément contient du texte (#PCDATA)
    """
    children   = defaultdict(list)   # parent -> [enfants ordonnés]
    attrs      = defaultdict(dict)   # elem   -> {attr: {values}}
    has_text   = defaultdict(bool)   # elem   -> True si contient du texte
    child_sets = defaultdict(set)    # pour détecter les enfants uniques vs multiples
    child_counts = defaultdict(lambda: defaultdict(int))  # parent -> child -> nb occurrences max par parent

    for elem in tree.iter():
        tag = local(elem.tag)
        if not isinstance(elem.tag, str):  # commentaires, PI
            continue

        # Texte direct
        if (elem.text and elem.text.strip()) or (elem.tail and elem.tail.strip()):
            has_text[tag] = True

        # Attributs
        for attr, val in elem.attrib.items():
            attr = local(attr)
            if attr not in attrs[tag]:
                attrs[tag][attr] = set()
            attrs[tag][attr].add(val)

        # Enfants
        seen_in_this_parent = defaultdict(int)
        for child in elem:
            if not isinstance(child.tag, str):
                continue
            ctag = local(child.tag)
            if ctag not in child_sets[tag]:
                child_sets[tag].add(ctag)
                children[tag].append(ctag)
            seen_in_this_parent[ctag] += 1

        for ctag, cnt in seen_in_this_parent.items():
            if cnt > child_counts[tag][ctag]:
                child_counts[tag][ctag] = cnt

    return children, attrs, has_text, child_counts


def content_model(tag, children, has_text, child_counts):
    """Construit le modèle de contenu DTD pour un élément."""
    kids = children.get(tag, [])
    txt  = has_text.get(tag, False)

    if not kids and not txt:
        return "EMPTY"

    if not kids and txt:
        return "(#PCDATA)"

    if kids and txt:
        # Contenu mixte
        parts = ["#PCDATA"] + kids
        return "(" + " | ".join(parts) + ")*"

    # Si les enfants peuvent s'intercaler (plusieurs types dans un ordre variable),
    # on utilise un modèle de choix répété plutôt qu'une séquence stricte.
    # Heuristique : si un élément parent a plus de types d'enfants que d'enfants
    # dans la séquence observée, c'est probablement un contenu mixte interleaved.
    if len(kids) > 1:
        # Vérifier si au moins un type d'enfant apparaît plusieurs fois
        any_repeated = any(child_counts[tag].get(c, 1) > 1 for c in kids)
        if any_repeated:
            parts = [c for c in kids]
            return "(" + " | ".join(parts) + ")*"

    # Séquence d'enfants
    parts = []
    for child in kids:
        max_occ = child_counts[tag].get(child, 1)
        indicator = "+" if max_occ > 1 else ""
        parts.append(child + indicator)
    return "(" + ", ".join(parts) + ")"


def attr_type(values):
    """Déduit le type d'attribut DTD depuis les valeurs observées."""
    if len(values) <= 5 and all(len(v) < 30 and ' ' not in v for v in values):
        return "(" + " | ".join(sorted(values)) + ")"
    return "CDATA"


def generate_dtd(tree):
    """Génère le texte de la DTD."""
    children, attrs, has_text, child_counts = analyze(tree)

    # Ordre topologique : racine d'abord, puis BFS
    root_tag = local(tree.getroot().tag)
    visited  = []
    queue    = [root_tag]
    seen     = set()

    while queue:
        tag = queue.pop(0)
        if tag in seen:
            continue
        seen.add(tag)
        visited.append(tag)
        for child in children.get(tag, []):
            if child not in seen:
                queue.append(child)

    lines = []
    for tag in visited:
        model = content_model(tag, children, has_text, child_counts)
        lines.append(f"<!ELEMENT {tag} {model}>")

        tag_attrs = attrs.get(tag, {})
        if tag_attrs:
            lines.append(f"<!ATTLIST {tag}")
            for attr, values in sorted(tag_attrs.items()):
                atype = attr_type(values)
                # Heuristique : si une seule valeur observée → valeur par défaut
                default = f'"{next(iter(values))}"' if len(values) == 1 else "#IMPLIED"
                lines.append(f"  {attr:<20} {atype:<30} {default}")
            lines[-1] += ">"  # ferme le ATTLIST sur la dernière ligne
            lines.append("")

    return "\n".join(lines)


def main():
    if len(sys.argv) < 2:
        print("Usage: python xml2dtd.py input.xml [output.dtd]")
        sys.exit(1)

    xml_path = sys.argv[1]
    dtd_path = sys.argv[2] if len(sys.argv) > 2 else xml_path.rsplit('.', 1)[0] + ".dtd"

    parser = etree.XMLParser(load_dtd=False, resolve_entities=False)
    tree   = etree.parse(xml_path, parser)

    dtd_text = generate_dtd(tree)

    with open(dtd_path, 'w', encoding='utf-8') as f:
        f.write(dtd_text)

    print(f"DTD générée : {dtd_path}")
    print(dtd_text)


if __name__ == "__main__":
    main()
