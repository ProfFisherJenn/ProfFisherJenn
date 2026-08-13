#!/usr/bin/env python3
"""
Word to Canvas QTI Converter
Matches folder structure: input_docs/ -> output_qti/
"""

import os
import sys
import re
import uuid
import zipfile
from pathlib import Path
from datetime import datetime
from docx import Document
from lxml import etree as ET

# ==============================
# Utilities
# ==============================
def log(msg: str):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def pause():
    try:
        if os.name == "nt":
            os.system("pause")
        else:
            input("Press Enter to exit...")
    except Exception:
        pass

def open_folder(path: Path):
    try:
        if os.name == "nt":
            os.startfile(str(path))
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
    except Exception as e:
        log(f"Could not open folder {path}: {e}")

# ==============================
# Paths
# ==============================
BASE_DIR = Path(__file__).resolve().parent
LOG_PATH = BASE_DIR / "qti_conversion_log.txt"
INPUT_DIR = BASE_DIR / "input_docs"
OUTPUT_DIR = BASE_DIR / "output_qti"
INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

try:
    LOG_PATH.unlink(missing_ok=True)
except Exception:
    pass

log("=== Word → Canvas QTI Converter started ===")
log(f"Script folder: {BASE_DIR}")

# ==============================
# QTI Generation Functions
# ==============================

def generate_id(prefix=""):
    """Generate a unique identifier"""
    return f"{prefix}{uuid.uuid4().hex[:12]}"

def clean_text(text):
    """Clean and normalize text"""
    if not text:
        return ""
    text = re.sub(r'\s+', ' ', text.strip())
    return text

def parse_docx(docx_path):
    """Parse Word document and extract questions"""
    doc = Document(docx_path)
    
    # Extract all text with formatting info
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            has_bold = any(run.bold for run in para.runs if run.text.strip())
            lines.append({
                'text': text,
                'bold': has_bold
            })
    
    questions = []
    i = 0
    current_type = None
    
    while i < len(lines):
        line_obj = lines[i]
        line = line_obj['text']
        
        # Skip headers and instructions
        if 'Module' in line or 'Instructions' in line or 'Correct matches:' in line:
            i += 1
            continue
        
        # Detect question type headers
        if line == 'Multiple Choice':
            current_type = 'multiple_choice'
            i += 1
            continue
        elif line == 'True/False':
            current_type = 'true_false'
            i += 1
            continue
        elif line == 'Multiple Answer':
            current_type = 'multiple_answer'
            i += 1
            continue
        elif line == 'Matching':
            current_type = 'matching'
            i += 1
            continue
        
        # Try to detect question based on pattern
        if ('?' in line or (current_type == 'true_false' and 'True or False' in line)) and not line.lower().startswith('correct'):
            # Determine type based on content if not set
            if current_type is None:
                if 'True or False' in line:
                    current_type = 'true_false'
                elif 'Select all' in line or 'all that apply' in line:
                    current_type = 'multiple_answer'
                else:
                    current_type = 'multiple_choice'
            
            if current_type in ['multiple_choice']:
                question, next_i = parse_multiple_choice(lines, i)
                questions.append(question)
                i = next_i
            elif current_type == 'true_false':
                question, next_i = parse_true_false(lines, i)
                questions.append(question)
                i = next_i
            elif current_type == 'multiple_answer':
                question, next_i = parse_multiple_answer(lines, i)
                questions.append(question)
                i = next_i
            else:
                i += 1
        elif current_type == 'matching' and not line.lower().startswith('correct'):
            question, next_i = parse_matching(lines, i)
            questions.append(question)
            i = next_i
        else:
            i += 1
    
    return questions

def parse_multiple_choice(lines, start_idx):
    """Parse a multiple choice question"""
    question_text = clean_text(lines[start_idx]['text'])
    
    choices = []
    i = start_idx + 1
    
    while i < len(lines):
        line_obj = lines[i]
        line = line_obj['text']
        
        if line.lower().startswith('correct answer:'):
            correct_text = clean_text(line.split(':', 1)[1])
            for choice in sorted(choices, key=lambda x: len(x['text']), reverse=True):
                if correct_text.lower() == choice['text'].lower():
                    choice['correct'] = True
                    break
                elif correct_text.lower() in choice['text'].lower() or choice['text'].lower() in correct_text.lower():
                    if len(choice['text']) > len(correct_text) * 0.8:
                        choice['correct'] = True
                        break
            i += 1
            break
        elif '?' in line or line in ['Multiple Choice', 'True/False', 'Multiple Answer', 'Matching']:
            break
        elif line and not line.lower().startswith('correct'):
            is_correct = line_obj.get('bold', False)
            choices.append({
                'text': clean_text(line),
                'correct': is_correct
            })
        
        i += 1
    
    return {
        'type': 'multiple_choice',
        'question': question_text,
        'choices': choices
    }, i

def parse_true_false(lines, start_idx):
    """Parse a true/false question"""
    question_text = clean_text(lines[start_idx]['text'])
    question_text = re.sub(r'^True or False:\s*', '', question_text, flags=re.IGNORECASE)
    
    choices = [
        {'text': 'True', 'correct': False},
        {'text': 'False', 'correct': False}
    ]
    
    i = start_idx + 1
    
    while i < len(lines):
        line_obj = lines[i]
        line = line_obj['text']
        
        if line.lower().startswith('correct answer:'):
            correct_answer = clean_text(line.split(':', 1)[1])
            for choice in choices:
                if choice['text'].lower() == correct_answer.lower():
                    choice['correct'] = True
            i += 1
            break
        elif line in ['True', 'False']:
            for choice in choices:
                if choice['text'] == line and line_obj.get('bold', False):
                    choice['correct'] = True
        elif '?' in line or line in ['Multiple Choice', 'True/False', 'Multiple Answer', 'Matching']:
            break
        
        i += 1
    
    return {
        'type': 'true_false',
        'question': question_text,
        'choices': choices
    }, i

def parse_multiple_answer(lines, start_idx):
    """Parse a multiple answer question"""
    question_text = clean_text(lines[start_idx]['text'])
    
    choices = []
    i = start_idx + 1
    
    while i < len(lines):
        line_obj = lines[i]
        line = line_obj['text']
        
        if line.lower().startswith('correct answer'):
            answers_part = line.split(':', 1)[1] if ':' in line else ''
            correct_texts = [clean_text(a) for a in answers_part.split(',')]
            
            for choice in choices:
                for correct_text in correct_texts:
                    if correct_text.lower() in choice['text'].lower():
                        choice['correct'] = True
            i += 1
            break
        elif '?' in line or line in ['Multiple Choice', 'True/False', 'Multiple Answer', 'Matching']:
            break
        elif line and not line.lower().startswith('correct'):
            is_correct = line_obj.get('bold', False)
            choices.append({
                'text': clean_text(line),
                'correct': is_correct
            })
        
        i += 1
    
    return {
        'type': 'multiple_answer',
        'question': question_text,
        'choices': choices
    }, i

def parse_matching(lines, start_idx):
    """Parse a matching question"""
    question_text = clean_text(lines[start_idx]['text'])
    
    matches = []
    i = start_idx + 1
    
    while i < len(lines):
        line = lines[i]['text']
        
        if '---' in line:
            parts = re.split(r'\s*---\s*', line, 1)
            if len(parts) == 2:
                matches.append({
                    'premise': clean_text(parts[0]),
                    'match': clean_text(parts[1])
                })
        elif line.lower().startswith('correct'):
            i += 1
            while i < len(lines) and ':' in lines[i]['text'] and not '?' in lines[i]['text']:
                i += 1
            break
        elif '?' in line or line in ['Multiple Choice', 'True/False', 'Multiple Answer', 'Matching']:
            break
        
        i += 1
    
    return {
        'type': 'matching',
        'question': question_text,
        'matches': matches
    }, i

def create_qti_xml(questions, title="Quiz Questions"):
    """Create QTI 2.1 XML structure"""
    
    NSMAP = {
        None: 'http://www.imsglobal.org/xsd/imsccv1p1/imscp_v1p1',
        'lom': 'http://ltsc.ieee.org/xsd/imsccv1p1/LOM/resource',
        'lomimscc': 'http://ltsc.ieee.org/xsd/imsccv1p1/LOM/manifest',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
    }
    
    manifest = ET.Element('manifest', {
        'identifier': generate_id('man_'),
        '{http://www.w3.org/2001/XMLSchema-instance}schemaLocation': 'http://www.imsglobal.org/xsd/imsccv1p1/imscp_v1p1 http://www.imsglobal.org/profile/cc/ccv1p1/ccv1p1_imscp_v1p2_v1p0.xsd'
    }, nsmap=NSMAP)
    
    metadata = ET.SubElement(manifest, 'metadata')
    schema = ET.SubElement(metadata, 'schema')
    schema.text = 'IMS Common Cartridge'
    schemaversion = ET.SubElement(metadata, 'schemaversion')
    schemaversion.text = '1.1.0'
    
    organizations = ET.SubElement(manifest, 'organizations')
    resources = ET.SubElement(manifest, 'resources')
    
    qb_resource = ET.SubElement(resources, 'resource', {
        'identifier': generate_id('qdb_'),
        'type': 'imsqti_xmlv1p2/imscc_xmlv1p1/assessment'
    })
    
    qb_file = ET.SubElement(qb_resource, 'file', {
        'href': 'assessment_qti.xml'
    })
    
    assessment = create_assessment_xml(questions, title)
    
    return manifest, assessment

def create_assessment_xml(questions, title):
    """Create the main assessment/question bank XML"""
    
    NSMAP = {
        None: 'http://www.imsglobal.org/xsd/ims_qtiasiv1p2',
        'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
    }
    
    assessment = ET.Element('questestinterop', {
        '{http://www.w3.org/2001/XMLSchema-instance}schemaLocation': 'http://www.imsglobal.org/xsd/ims_qtiasiv1p2 http://www.imsglobal.org/xsd/ims_qtiasiv1p2p1.xsd'
    }, nsmap=NSMAP)
    
    assessment_meta = ET.SubElement(assessment, 'assessment', {
        'ident': generate_id('assessment_'),
        'title': title
    })
    
    section = ET.SubElement(assessment_meta, 'section', {
        'ident': generate_id('section_')
    })
    
    for q in questions:
        if q['type'] == 'multiple_choice':
            add_multiple_choice_question(section, q)
        elif q['type'] == 'true_false':
            add_true_false_question(section, q)
        elif q['type'] == 'multiple_answer':
            add_multiple_answer_question(section, q)
        elif q['type'] == 'matching':
            add_matching_question(section, q)
    
    return assessment

def add_metadata_field(qtimetadata, label, entry):
    """Add a metadata field"""
    qtimetadatafield = ET.SubElement(qtimetadata, 'qtimetadatafield')
    fieldlabel = ET.SubElement(qtimetadatafield, 'fieldlabel')
    fieldlabel.text = label
    fieldentry = ET.SubElement(qtimetadatafield, 'fieldentry')
    fieldentry.text = str(entry)

def add_multiple_choice_question(section, question):
    """Add a multiple choice question to QTI XML"""
    
    item_id = generate_id('item_')
    item = ET.SubElement(section, 'item', {
        'ident': item_id,
        'title': 'Question'
    })
    
    itemmetadata = ET.SubElement(item, 'itemmetadata')
    qtimetadata = ET.SubElement(itemmetadata, 'qtimetadata')
    
    add_metadata_field(qtimetadata, 'question_type', 'multiple_choice_question')
    add_metadata_field(qtimetadata, 'points_possible', '1')
    add_metadata_field(qtimetadata, 'assessment_question_identifierref', item_id)
    
    presentation = ET.SubElement(item, 'presentation')
    material = ET.SubElement(presentation, 'material')
    mattext = ET.SubElement(material, 'mattext', {'texttype': 'text/html'})
    mattext.text = f"<p>{question['question']}</p>"
    
    response = ET.SubElement(presentation, 'response_lid', {
        'ident': 'response1',
        'rcardinality': 'Single'
    })
    render_choice = ET.SubElement(response, 'render_choice')
    
    choice_ids = []
    for i, choice in enumerate(question['choices']):
        choice_id = generate_id(f'choice_{i}_')
        choice_ids.append((choice_id, choice['correct']))
        
        response_label = ET.SubElement(render_choice, 'response_label', {'ident': choice_id})
        mat = ET.SubElement(response_label, 'material')
        mat_text = ET.SubElement(mat, 'mattext', {'texttype': 'text/plain'})
        mat_text.text = choice['text']
    
    resprocessing = ET.SubElement(item, 'resprocessing')
    outcomes = ET.SubElement(resprocessing, 'outcomes')
    decvar = ET.SubElement(outcomes, 'decvar', {
        'maxvalue': '100',
        'minvalue': '0',
        'varname': 'SCORE',
        'vartype': 'Decimal'
    })
    
    for choice_id, is_correct in choice_ids:
        if is_correct:
            respcondition = ET.SubElement(resprocessing, 'respcondition', {'continue': 'No'})
            conditionvar = ET.SubElement(respcondition, 'conditionvar')
            varequal = ET.SubElement(conditionvar, 'varequal', {'respident': 'response1'})
            varequal.text = choice_id
            setvar = ET.SubElement(respcondition, 'setvar', {
                'action': 'Set',
                'varname': 'SCORE'
            })
            setvar.text = '100'

def add_true_false_question(section, question):
    """Add a true/false question to QTI XML"""
    add_multiple_choice_question(section, question)

def add_multiple_answer_question(section, question):
    """Add a multiple answer question to QTI XML"""
    
    item_id = generate_id('item_')
    item = ET.SubElement(section, 'item', {
        'ident': item_id,
        'title': 'Question'
    })
    
    itemmetadata = ET.SubElement(item, 'itemmetadata')
    qtimetadata = ET.SubElement(itemmetadata, 'qtimetadata')
    
    add_metadata_field(qtimetadata, 'question_type', 'multiple_answers_question')
    add_metadata_field(qtimetadata, 'points_possible', '1')
    add_metadata_field(qtimetadata, 'assessment_question_identifierref', item_id)
    
    presentation = ET.SubElement(item, 'presentation')
    material = ET.SubElement(presentation, 'material')
    mattext = ET.SubElement(material, 'mattext', {'texttype': 'text/html'})
    mattext.text = f"<p>{question['question']}</p>"
    
    response = ET.SubElement(presentation, 'response_lid', {
        'ident': 'response1',
        'rcardinality': 'Multiple'
    })
    render_choice = ET.SubElement(response, 'render_choice')
    
    choice_ids = []
    for i, choice in enumerate(question['choices']):
        choice_id = generate_id(f'choice_{i}_')
        choice_ids.append((choice_id, choice['correct']))
        
        response_label = ET.SubElement(render_choice, 'response_label', {'ident': choice_id})
        mat = ET.SubElement(response_label, 'material')
        mat_text = ET.SubElement(mat, 'mattext', {'texttype': 'text/plain'})
        mat_text.text = choice['text']
    
    resprocessing = ET.SubElement(item, 'resprocessing')
    outcomes = ET.SubElement(resprocessing, 'outcomes')
    decvar = ET.SubElement(outcomes, 'decvar', {
        'maxvalue': '100',
        'minvalue': '0',
        'varname': 'SCORE',
        'vartype': 'Decimal'
    })
    
    respcondition = ET.SubElement(resprocessing, 'respcondition', {'continue': 'No'})
    conditionvar = ET.SubElement(respcondition, 'conditionvar')
    and_condition = ET.SubElement(conditionvar, 'and')
    
    for choice_id, is_correct in choice_ids:
        if is_correct:
            varequal = ET.SubElement(and_condition, 'varequal', {'respident': 'response1'})
            varequal.text = choice_id
        else:
            not_elem = ET.SubElement(and_condition, 'not')
            varequal = ET.SubElement(not_elem, 'varequal', {'respident': 'response1'})
            varequal.text = choice_id
    
    setvar = ET.SubElement(respcondition, 'setvar', {
        'action': 'Set',
        'varname': 'SCORE'
    })
    setvar.text = '100'

def add_matching_question(section, question):
    """Add a matching question to QTI XML"""
    
    item_id = generate_id('item_')
    item = ET.SubElement(section, 'item', {
        'ident': item_id,
        'title': 'Question'
    })
    
    itemmetadata = ET.SubElement(item, 'itemmetadata')
    qtimetadata = ET.SubElement(itemmetadata, 'qtimetadata')
    
    add_metadata_field(qtimetadata, 'question_type', 'matching_question')
    add_metadata_field(qtimetadata, 'points_possible', '1')
    add_metadata_field(qtimetadata, 'assessment_question_identifierref', item_id)
    
    presentation = ET.SubElement(item, 'presentation')
    material = ET.SubElement(presentation, 'material')
    mattext = ET.SubElement(material, 'mattext', {'texttype': 'text/html'})
    mattext.text = f"<p>{question['question']}</p>"
    
    match_ids = []
    
    # First pass: generate all premise and match IDs
    for i, match_pair in enumerate(question['matches']):
        premise_id = generate_id(f'premise_{i}_')
        match_id = generate_id(f'match_{i}_')
        match_ids.append((premise_id, match_id))
    
    # Second pass: build the XML structure with consistent IDs
    for i, match_pair in enumerate(question['matches']):
        premise_id, _ = match_ids[i]
        
        response = ET.SubElement(presentation, 'response_lid', {
            'ident': premise_id,
            'rcardinality': 'Single'
        })
        
        mat = ET.SubElement(response, 'material')
        mat_text = ET.SubElement(mat, 'mattext', {'texttype': 'text/plain'})
        mat_text.text = match_pair['premise']
        
        render_choice = ET.SubElement(response, 'render_choice')
        for j, mp in enumerate(question['matches']):
            _, match_id = match_ids[j]  # Use the stored match_id
            response_label = ET.SubElement(render_choice, 'response_label', {
                'ident': match_id
            })
            mat2 = ET.SubElement(response_label, 'material')
            mat_text2 = ET.SubElement(mat2, 'mattext', {'texttype': 'text/plain'})
            mat_text2.text = mp['match']
    
    resprocessing = ET.SubElement(item, 'resprocessing')
    outcomes = ET.SubElement(resprocessing, 'outcomes')
    decvar = ET.SubElement(outcomes, 'decvar', {
        'maxvalue': '100',
        'minvalue': '0',
        'varname': 'SCORE',
        'vartype': 'Decimal'
    })
    
    for i, (premise_id, match_id) in enumerate(match_ids):
        respcondition = ET.SubElement(resprocessing, 'respcondition', {'continue': 'Yes'})
        conditionvar = ET.SubElement(respcondition, 'conditionvar')
        varequal = ET.SubElement(conditionvar, 'varequal', {'respident': premise_id})
        varequal.text = match_id
        setvar = ET.SubElement(respcondition, 'setvar', {
            'action': 'Add',
            'varname': 'SCORE'
        })
        points_per_match = round(100 / len(match_ids), 2)
        setvar.text = str(points_per_match)

def create_qti_package(questions, output_path, title="Quiz Questions"):
    """Create a complete QTI package as a zip file"""
    
    manifest, assessment = create_qti_xml(questions, title)
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        manifest_xml = ET.tostring(manifest, encoding='utf-8', xml_declaration=True, pretty_print=True)
        zf.writestr('imsmanifest.xml', manifest_xml)
        
        assessment_xml = ET.tostring(assessment, encoding='utf-8', xml_declaration=True, pretty_print=True)
        zf.writestr('assessment_qti.xml', assessment_xml)

def convert_one(doc_path: Path, out_path: Path):
    """Convert one Word document to QTI zip"""
    log(f"Converting: {doc_path.name}")
    
    questions = parse_docx(doc_path)
    
    if not questions:
        log(f"  WARNING: No questions found in {doc_path.name}")
        return False
    
    log(f"  Found {len(questions)} questions")
    
    create_qti_package(questions, out_path, doc_path.stem)
    log(f"  Saved: {out_path.name}")
    return True

# ==============================
# Main
# ==============================
def main():
    docs = sorted(INPUT_DIR.glob("*.docx"))
    if not docs:
        log(f"No .docx files found in: {INPUT_DIR}")
        open_folder(INPUT_DIR)
        pause()
        return

    success_count = 0
    for doc_path in docs:
        out_path = OUTPUT_DIR / (doc_path.stem + "_qti.zip")
        try:
            if convert_one(doc_path, out_path):
                os.remove(doc_path)
                log(f"  Removed source: {doc_path.name}")
                success_count += 1
        except Exception as ex:
            log(f"  FAILED: {doc_path.name} — {ex}")

    log(f"Conversion complete: {success_count} file(s) processed")
    log("Opening output folder...")
    open_folder(OUTPUT_DIR)
    pause()

if __name__ == "__main__":
    main()
