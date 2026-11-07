#!/usr/bin/env python3
"""

Installation of Python and Required Libs:
    On Debian/Ubuntu:
        sudo apt install python3 python3-pip python3-docx -y
"""

import sys
import os
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_hyperlink_to_paragraph(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    c = OxmlElement('w:color')
    c.set(qn('w:val'), '0000FF')
    rPr.append(c)

    u = OxmlElement('w:val')
    u.set(qn('w:val'), 'single')
    rPr.append(u)

    new_run.append(rPr)
    
    text_node = OxmlElement('w:t')
    text_node.text = text
    new_run.append(text_node)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def is_url_text(text):
    text_lower = text.strip().lower()
    return text_lower.startswith("http://") or text_lower.startswith("https://") or text_lower.startswith("www.")

# Only replace anchor text with [idx] if it does NOT contain 
# characters other than: (, ), [, ], :, ., and DIGITs.
def has_disallowed_chars(txt):
    allowed = set("()[]:.,0123456789 ")
    for ch in txt:
        if ch not in allowed:
            return True
    return False

def convert_hyperlinks_to_addendum(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File '{file_path}' not found.")
        sys.exit(1)
        
    if not file_path.endswith('.docx'):
        print("Error: Please provide a valid .docx file.")
        sys.exit(1)

    print(f"Processing: {file_path}...")
    
    doc = docx.Document(file_path)
    rels = doc.part.rels
    
    url_to_index = {}
    unique_urls = []

    # First pass: collect unique URLs and assign order
    for rel in rels.values():
        if "hyperlink" in rel.reltype:
            target = rel.target_ref
            if target not in url_to_index:
                unique_urls.append(target)
                url_to_index[target] = len(unique_urls)

    if not unique_urls:
        print("No hyperlinks found in the document.")
        return

    # Second pass: check anchor text of each hyperlink in paragraphs
    for p in doc.paragraphs:
        for child in list(p._p):
            if child.tag.endswith('hyperlink'):
                r_id = child.get(qn('r:id'))
                if r_id in rels:
                    target_url = rels[r_id].target_ref
                    if target_url in url_to_index:
                        idx = url_to_index[target_url]
                        
                        # Extract the current anchor text inside the hyperlink element
                        current_text = "".join([t.text for t in child.iter(qn('w:t')) if t.text])
                        
                        # Only replace anchor text with [idx] if it does NOT contain 
                        # characters other than: (, ), [, ], :, ., and DIGITs.
                        if not has_disallowed_chars(current_text):
                            for subnode in list(child):
                                child.remove(subnode)
                            
                            r = OxmlElement('w:r')
                            t = OxmlElement('w:t')
                            t.text = f"[{idx}]"
                            r.append(t)
                            child.append(r)

    # Append Addendum Section with clickable links
    heading_p = doc.add_paragraph()
    heading_run = heading_p.add_run("Addendum: Reference Links")
    heading_run.bold = True
    heading_run.font.size = docx.shared.Pt(14)

    for index, url in enumerate(unique_urls, start=1):
        p = doc.add_paragraph(f"[{index}] ")
        add_hyperlink_to_paragraph(p, url, url)

    base, ext = os.path.splitext(file_path)
    output_path = f"{base} - with addendum of links{ext}"
    doc.save(output_path)
    
    print(f"Done! Converted {len(unique_urls)} link(s). Saved as: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 links_to_addendum.py <FileName.docx>")
        sys.exit(1)
        
    convert_hyperlinks_to_addendum(sys.argv[1])