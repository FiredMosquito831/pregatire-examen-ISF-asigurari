import json
import re
from docx import Document
import os

# Numele fișierului tău
INPUT_FILENAME = 'întrebări EXAMEN ISF.docx'


def clean_text(text):
    """Curăță textul de spații inutile și caractere invizibile."""
    if not text:
        return ""
    return " ".join(text.split()).strip()


def parse_answer_text(text):
    """
    Extrage textul răspunsului și verifică dacă este corect.
    Returnează (text_curat, is_correct).
    """
    is_correct = False

    # Verificăm marcajul '*'
    if '*' in text:
        is_correct = True
        text = text.replace('*', '')

    # Eliminăm prefixele de tip "a)", "b.", "c "
    # Regex: litera la inceput, urmata de ) sau . sau spatiu
    text = re.sub(r'^[a-zA-Z][\)\.]\s*', '', text.strip())

    return text.strip(), is_correct


def detect_columns(table):
    """
    Încearcă să ghicească indicii coloanelor analizând primele 10 rânduri.
    Returnează (col_id, col_quest, col_ans).
    """
    print("Se detectează structura tabelului...")
    id_scores = {}
    q_scores = {}
    a_scores = {}

    # Analizăm primele 15 rânduri pentru a determina tiparul
    rows_to_check = table.rows[:15]
    num_cols = len(rows_to_check[0].cells)

    for c in range(num_cols):
        id_scores[c] = 0
        q_scores[c] = 0
        a_scores[c] = 0

    for row in rows_to_check:
        cells = row.cells
        for i, cell in enumerate(cells):
            txt = clean_text(cell.text)
            if not txt:
                continue

            # Scor pentru ID: este număr (ex: "1", "25")
            if re.match(r'^\d+$', txt):
                id_scores[i] += 1

            # Scor pentru Întrebare: lungime > 15 caractere și conține '?'
            if len(txt) > 15 and '?' in txt:
                q_scores[i] += 2
            elif len(txt) > 20 and not re.match(r'^[a-zA-Z][\)\.]', txt):
                # Text lung care nu pare răspuns
                q_scores[i] += 1

            # Scor pentru Răspuns: începe cu "a)", "b)", "*"
            if re.match(r'^[a-zA-Z][\)\.]', txt) or '*' in txt:
                a_scores[i] += 2

    # Alegem coloana cu scorul maxim pentru fiecare
    best_id = max(id_scores, key=id_scores.get)
    best_q = max(q_scores, key=q_scores.get)
    best_a = max(a_scores, key=a_scores.get)

    print(f"Coloane detectate -> ID: {best_id}, Întrebare: {best_q}, Răspunsuri: {best_a}")
    return best_id, best_q, best_a


def extract_data(file_path):
    if not os.path.exists(file_path):
        print(f"Fișierul {file_path} nu există.")
        return []

    doc = Document(file_path)
    all_questions = []

    # Procesăm fiecare tabel din document
    for table in doc.tables:
        # Detectăm coloanele pentru acest tabel
        if not table.rows:
            continue

        col_id, col_q, col_a = detect_columns(table)

        current_q = None

        for row in table.rows:
            cells = row.cells

            # Siguranță pentru rânduri scurte
            if len(cells) <= max(col_id, col_q, col_a):
                continue

            txt_id = clean_text(cells[col_id].text)
            txt_q = clean_text(cells[col_q].text)
            txt_a = clean_text(cells[col_a].text)

            # Caz 1: Avem un ID nou -> Începe o nouă întrebare
            if txt_id and txt_id.isdigit():
                # Salvăm întrebarea anterioară
                if current_q:
                    all_questions.append(current_q)

                current_q = {
                    "id": txt_id,
                    "question": txt_q,
                    "answers": []
                }

                # Verificăm dacă există și un răspuns pe același rând
                if txt_a:
                    ans_text, correct = parse_answer_text(txt_a)
                    if ans_text:
                        current_q['answers'].append({
                            "text": ans_text,
                            "is_correct": correct
                        })

            # Caz 2: Nu avem ID, dar avem text la Răspuns -> Continuarea întrebării curente
            elif not txt_id and txt_a and current_q:
                # Verificăm dacă celula de răspuns conține mai multe linii
                lines = txt_a.split('\n')
                for line in lines:
                    line = clean_text(line)
                    if not line: continue

                    # Verificăm să nu duplicăm răspunsuri (python-docx uneori vede celule merge-uite repetat)
                    # Simplu check: dacă textul e identic cu ultimul adăugat
                    is_duplicate = False
                    if current_q['answers']:
                        last_ans = current_q['answers'][-1]['text']
                        if line.replace('*', '').strip().endswith(last_ans[-20:]):  # Check parțial
                            is_duplicate = True

                    if not is_duplicate:
                        ans_text, correct = parse_answer_text(line)
                        if ans_text:
                            current_q['answers'].append({
                                "text": ans_text,
                                "is_correct": correct
                            })

            # Caz 3: Text suplimentar la întrebare (dacă întrebarea se întinde pe mai multe rânduri)
            elif not txt_id and txt_q and current_q:
                # Dacă textul întrebării nu e identic cu ce avem deja
                if txt_q not in current_q['question']:
                    current_q['question'] += " " + txt_q

        # Adăugăm ultima întrebare
        if current_q:
            all_questions.append(current_q)

    return all_questions


def save_output(data):
    # Salvare JSON
    with open('intrebari.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    print("JSON salvat: intrebari.json")

    # Salvare HTML
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body { font-family: sans-serif; padding: 20px; max-width: 800px; margin: 0 auto; }
            .card { border: 1px solid #ddd; padding: 15px; margin-bottom: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
            .header { display: flex; justify-content: space-between; color: #666; font-size: 0.9em; margin-bottom: 10px; }
            .question { font-weight: bold; font-size: 1.1em; margin-bottom: 15px; color: #333; }
            .answer { padding: 8px; margin: 5px 0; border-radius: 4px; background: #f8f9fa; border: 1px solid #eee; }
            .correct { background-color: #d4edda; border-color: #c3e6cb; color: #155724; }
            .correct::before { content: "✓ "; font-weight: bold; }
        </style>
    </head>
    <body>
        <h1>Întrebări Extrase ISF</h1>
    """

    for q in data:
        html += f"""
        <div class="card">
            <div class="header">ID Întrebare: {q['id']}</div>
            <div class="question">{q['question']}</div>
            <div class="answers">
        """
        for ans in q['answers']:
            cls = "answer correct" if ans['is_correct'] else "answer"
            html += f'<div class="{cls}">{ans["text"]}</div>'
        html += "</div></div>"

    html += "</body></html>"

    with open('intrebari.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print("HTML salvat: intrebari.html")


if __name__ == "__main__":
    data = extract_data(INPUT_FILENAME)
    print(f"Total întrebări extrase: {len(data)}")
    if data:
        save_output(data)
    else:
        print("Nu s-au găsit date. Verifică numele fișierului.")