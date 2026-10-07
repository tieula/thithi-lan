import eventlet
eventlet.monkey_patch()
# -*- coding: utf-8 -*-
import os
import socket
import re
import time
import random
import copy
from io import BytesIO
from docx import Document
from flask import Flask, render_template, request, jsonify, send_file
from flask_socketio import SocketIO, emit
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['SECRET_KEY'] = 'lan_exam_secret_key_2026'
socketio = SocketIO(app, cors_allowed_origins="*")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SUBMISSION_DIR = os.path.join(BASE_DIR, 'bailamthisinh')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')

for d in [SUBMISSION_DIR, UPLOAD_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

submitted_ips = set()  # Lưu danh sách IP đã nộp bài để chặn thi lần 2

exam_state = {
    "class_name": "",
    "raw_bank": {"filename": "", "mcq": [], "tf": []},
    "active_exam": {
        "num_mcq": 0,
        "score_mcq": 0.0,
        "score_per_mcq": 0.0,
        "num_tf": 0,
        "score_tf": 0.0,
        "score_per_tf": 0.0,
        "duration": 15,
        "tf_scale": {"1": 0.1, "2": 0.25, "3": 0.5, "4": 1.0},
        "status": "waiting",
        "start_time": None
    },
    "payload": None,
    "students": {}  # sid -> {ip, name, time_left, violations, submitted, score, answers, exam}
}

def get_base_url():
    """Tự động ưu tiên domain Render khi chạy online để tạo QR chuẩn xác"""
    render_url = os.environ.get('RENDER_EXTERNAL_URL')
    if render_url:
        return render_url.rstrip('/')
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return f"http://{ip}:5000"

def clean_question_text(raw_text):
    """Gọt sạch tiền tố 'Câu 1:', 'Câu 15.', 'Câu 48:' cũ trong văn bản ngân hàng đề"""
    if not raw_text:
        return ""
    return re.sub(r'^(câu\s*\d+[\s\:\.\-\)]*)\s*', '', raw_text.strip(), flags=re.IGNORECASE)

def generate_individual_exam():
    """Tạo đề thi ngẫu nhiên riêng cho từng thí sinh"""
    num_mcq = min(exam_state["active_exam"]["num_mcq"], len(exam_state["raw_bank"]["mcq"]))
    num_tf = min(exam_state["active_exam"]["num_tf"], len(exam_state["raw_bank"]["tf"]))

    selected_mcq_raw = random.sample(exam_state["raw_bank"]["mcq"], num_mcq) if num_mcq > 0 else []
    selected_tf_raw = random.sample(exam_state["raw_bank"]["tf"], num_tf) if num_tf > 0 else []

    random.shuffle(selected_mcq_raw)
    random.shuffle(selected_tf_raw)

    mcq_send = []
    mcq_server = []

    labels = ['A', 'B', 'C', 'D']
    for idx, q in enumerate(selected_mcq_raw, 1):
        shuffled_opts = copy.deepcopy(q['options'])
        random.shuffle(shuffled_opts)

        client_opts = []
        server_opts = []

        for o_idx, opt in enumerate(shuffled_opts):
            key = labels[o_idx] if o_idx < len(labels) else opt['key']
            client_opts.append({"key": key, "text": opt["text"]})
            server_opts.append({"key": key, "text": opt["text"], "correct": opt.get("correct", False)})

        clean_q = clean_question_text(q["question"])
        mcq_send.append({
            "id": idx,
            "type": "mcq",
            "question": clean_q,
            "options": client_opts
        })
        mcq_server.append({
            "id": idx,
            "type": "mcq",
            "question": clean_q,
            "options": server_opts
        })

    tf_send = []
    tf_server = []
    for idx, q in enumerate(selected_tf_raw, 1):
        client_subs = [{"key": s["key"], "text": s["text"]} for s in q["sub_items"]]
        server_subs = [{"key": s["key"], "text": s["text"], "correct": s.get("correct", False)} for s in q["sub_items"]]

        clean_q_tf = clean_question_text(q["question"])
        tf_send.append({
            "id": idx,
            "type": "tf",
            "question": clean_q_tf,
            "sub_items": client_subs
        })
        tf_server.append({
            "id": idx,
            "type": "tf",
            "question": clean_q_tf,
            "sub_items": server_subs
        })

    client_payload = {
        "mcq": mcq_send,
        "tf": tf_send,
        "duration": exam_state["active_exam"]["duration"]
    }
    server_exam = {
        "mcq": mcq_server,
        "tf": tf_server
    }
    return client_payload, server_exam

def update_summary_excel():
    """Tự động cập nhật tệp bảng điểm chung kèm bảng thống kê"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bang_Diem_Tong_Hop"

    thin_border = Border(
        left=Side(style='thin', color='B0B0B0'),
        right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'),
        bottom=Side(style='thin', color='B0B0B0')
    )
    navy_fill = PatternFill(start_color="0B3C6D", end_color="0B3C6D", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    bold_font = Font(name="Arial", size=10, bold=True)
    regular_font = Font(name="Arial", size=10)

    headers = ["TT", "Địa chỉ IP", "Họ và tên", "Thời gian còn lại", "Vi phạm", "Trạng thái", "Điểm số"]
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = navy_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    scores_list = []
    idx = 1
    current_row = 2

    for sid, st in exam_state["students"].items():
        mins = st['time_left'] // 60
        secs = st['time_left'] % 60
        score_val = float(st.get("score", 0.0))
        scores_list.append(score_val)

        ws.append([
            idx,
            st["ip"],
            st["name"],
            f"{mins:02d}:{secs:02d}",
            st["violations"],
            st["submitted"],
            score_val
        ])

        for col_idx in range(1, 8):
            c = ws.cell(row=current_row, column=col_idx)
            c.font = regular_font
            c.border = thin_border
            if col_idx == 3:
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif col_idx == 7:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = bold_font
            else:
                c.alignment = Alignment(horizontal="center", vertical="center")

        idx += 1
        current_row += 1

    total_students = len(scores_list)
    cnt_duoi_5 = sum(1 for s in scores_list if s <= 5.0)
    cnt_duoi_65 = sum(1 for s in scores_list if 5.0 < s <= 6.5)
    cnt_duoi_8 = sum(1 for s in scores_list if 6.5 < s <= 8.0)
    cnt_tren_8 = sum(1 for s in scores_list if s > 8.0)

    def calc_rate(count):
        if total_students == 0:
            return "0.0%"
        return f"{round((count / total_students) * 100, 1)}%"

    stat_start_row = current_row + 2
    stat_headers = ["Mức điểm", "Tỉ lệ", "Số lượng"]
    stat_data = [
        ["<= 5.0", calc_rate(cnt_duoi_5), cnt_duoi_5],
        ["<= 6.5", calc_rate(cnt_duoi_65), cnt_duoi_65],
        ["<= 8.0", calc_rate(cnt_duoi_8), cnt_duoi_8],
        ["> 8", calc_rate(cnt_tren_8), cnt_tren_8]
    ]

    for c_idx, h_text in enumerate(stat_headers, 1):
        cell = ws.cell(row=stat_start_row, column=c_idx, value=h_text)
        cell.font = bold_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    for r_offset, row_values in enumerate(stat_data, 1):
        r_idx = stat_start_row + r_offset
        for c_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = regular_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    col_widths = {1: 12, 2: 20, 3: 32, 4: 20, 5: 14, 6: 16, 7: 14}
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    summary_path = os.path.join(SUBMISSION_DIR, "Bang_Diem_Tong_Hop.xlsx")
    wb.save(summary_path)

def save_individual_submission(st, answers, score):
    """Xuất file bài làm chi tiết của từng thí sinh"""
    safe_name = re.sub(r'[^\w\s-]', '', st['name']).strip().replace(" ", "_")
    safe_ip = st["ip"].replace(":", "_")
    filename = f"{safe_name}_{safe_ip}.txt"
    filepath = os.path.join(SUBMISSION_DIR, filename)

    exam_data = st.get("exam", exam_state["raw_bank"])

    with open(filepath, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write(f"BÀI LÀM THÍ SINH - {st['name']}\n")
        f.write(f"Địa chỉ IP : {st['ip']}\n")
        f.write(f"Trạng thái : {st['submitted']}\n")
        f.write(f"Số lần vi phạm: {st['violations']}\n")
        f.write(f"ĐIỂM TỔNG KẾT: {score}\n")
        f.write("=" * 60 + "\n\n")

        f.write("--- PHẦN 1: TRẮC NGHIỆM NHIỀU LỰA CHỌN ---\n")
        mcq_list = exam_data.get("mcq", [])
        for q in mcq_list:
            ans = answers.get(f"mcq_{q['id']}", "Chưa làm")
            correct_opt = next((opt['key'] for opt in q['options'] if opt.get('correct')), "")
            f.write(f"Câu {q['id']}: {q['question']}\n")
            for opt in q['options']:
                mark = "(#)" if opt.get('correct') else "    "
                f.write(f"   {mark} {opt['key']}. {opt['text']}\n")
            f.write(f"   => Thí sinh chọn: {ans} | Đáp án đúng: {correct_opt}\n\n")

        f.write("\n--- PHẦN 2: TRẮC NGHIỆM ĐÚNG / SAI ---\n")
        tf_list = exam_data.get("tf", [])
        for q in tf_list:
            f.write(f"Câu {q['id']}: {q['question']}\n")
            for sub in q.get('sub_items', []):
                user_val = answers.get(f"tf_{q['id']}_{sub['key']}", "Chưa làm")
                expected = "Đ" if sub.get('correct') else "S"
                f.write(f"   {sub['key']}) {sub['text']}\n")
                f.write(f"      => Thí sinh chọn: {user_val} | Đáp án đúng: {expected}\n")
            f.write("\n")

def process_grading(sid, answers=None, student_name=None):
    """Hàm nội bộ thực hiện tính điểm, chống mất dấu học sinh khi bị đổi SID lúc nộp bài"""
    st = None
    target_sid = sid

    # 1. Tìm thí sinh theo sid hiện tại
    if sid in exam_state["students"]:
        st = exam_state["students"][sid]
    # 2. Nếu không thấy do rớt mạng đổi sid, tìm theo Tên thí sinh
    elif student_name:
        for s_id, s_data in exam_state["students"].items():
            if s_data.get("name") == student_name and s_data.get("submitted") != "Đã nộp":
                st = s_data
                target_sid = s_id
                break

    if not st:
        return

    # Nếu đã được chấm rồi thì gửi lại điểm về cho client
    if st.get("submitted") == "Đã nộp":
        emit('exam_finished_ack', {"score": st.get("score", 0.0)}, room=sid)
        return

    # Sử dụng câu trả lời gửi lên; nếu rỗng thì lấy câu trả lời đã đồng bộ thời gian thực
    if not answers and st.get("answers"):
        answers = st["answers"]
    elif not answers:
        answers = {}

    st["submitted"] = "Đã nộp"
    st["answers"] = answers
    submitted_ips.add(st["ip"])

    student_exam = st.get("exam")
    if not student_exam:
        student_exam = {
            "mcq": exam_state["raw_bank"]["mcq"][:exam_state["active_exam"]["num_mcq"]],
            "tf": exam_state["raw_bank"]["tf"][:exam_state["active_exam"]["num_tf"]]
        }

    score = 0.0
    mcq_pts = exam_state["active_exam"].get("score_per_mcq", 0.25)
    for q in student_exam.get("mcq", []):
        ans = answers.get(f"mcq_{q['id']}")
        correct_opt = next((opt['key'] for opt in q['options'] if opt.get('correct')), None)
        if ans and str(ans).strip().upper() == str(correct_opt).strip().upper():
            score += mcq_pts

    tf_pts = exam_state["active_exam"].get("score_per_tf", 1.0)
    scale = exam_state["active_exam"].get("tf_scale", {"1": 0.1, "2": 0.25, "3": 0.5, "4": 1.0})
    for q in student_exam.get("tf", []):
        correct_cnt = 0
        for sub in q.get('sub_items', []):
            user_val = answers.get(f"tf_{q['id']}_{sub['key']}")
            expected = "Đ" if sub.get('correct') else "S"
            if user_val and str(user_val).strip() == expected:
                correct_cnt += 1
        ratio = float(scale.get(str(correct_cnt), 0.0))
        score += tf_pts * ratio

    st["score"] = round(score, 2)
    st["time_left"] = 0

    save_individual_submission(st, answers, st["score"])
    update_summary_excel()

    # Cập nhật danh sách bảng giám sát giáo viên ngay lập tức
    emit('update_teacher_list', list(exam_state["students"].values()), broadcast=True)
    # Phản hồi báo điểm về cho học sinh
    emit('exam_finished_ack', {"score": st["score"]}, room=sid)
    if target_sid != sid:
        emit('exam_finished_ack', {"score": st["score"]}, room=target_sid)

@app.route('/')
def teacher_dashboard():
    return render_template('index.html')

@app.route('/student')
def student_view():
    return render_template('student.html')

@app.route('/api/get_ip', methods=['GET'])
def api_get_ip():
    base_url = get_base_url()
    return jsonify({
        "ip": base_url.replace("http://", "").replace("https://", "").split(":")[0],
        "port": 5000,
        "url": f"{base_url}/student"
    })

@app.route('/api/upload_bank', methods=['POST'])
def upload_bank():
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "Không tìm thấy file tải lên!"})
    
    file = request.files['file']
    filename = file.filename
    save_path = os.path.join(UPLOAD_DIR, filename)
    file.save(save_path)

    try:
        doc = Document(save_path)
    except Exception as e:
        return jsonify({"success": False, "error": f"Lỗi đọc file: {str(e)}"})

    paragraphs = []
    for p in doc.paragraphs:
        for sub in p.text.split('\n'):
            s = sub.strip()
            if s:
                paragraphs.append(s)

    mcq_questions = []
    tf_questions = []
    current_mode = None
    curr_q = None
    q_counter_mcq = 0
    q_counter_tf = 0

    for line in paragraphs:
        lower_line = line.lower()
        if lower_line.startswith('g1') or 'phần 1' in lower_line or 'phan 1' in lower_line or 'part 1' in lower_line:
            if curr_q:
                if current_mode == 'g1': mcq_questions.append(curr_q)
                elif current_mode == 'g2': tf_questions.append(curr_q)
                curr_q = None
            current_mode = 'g1'
            continue
        elif lower_line.startswith('g2') or 'phần 2' in lower_line or 'phan 2' in lower_line or 'part 2' in lower_line:
            if curr_q:
                if current_mode == 'g1': mcq_questions.append(curr_q)
                elif current_mode == 'g2': tf_questions.append(curr_q)
                curr_q = None
            current_mode = 'g2'
            continue

        if current_mode == 'g1':
            match_opt = re.match(r'^(#?)([A-Da-d])[\.\)]\s*(.*)', line)
            if match_opt:
                if not curr_q:
                    return jsonify({"success": False, "error": f"Lỗi tại phương án: '{line}'. Chưa có câu hỏi!"})
                is_correct = bool(match_opt.group(1) == '#')
                curr_q['options'].append({"key": match_opt.group(2).upper(), "text": match_opt.group(3), "correct": is_correct})
            else:
                if curr_q: mcq_questions.append(curr_q)
                q_counter_mcq += 1
                curr_q = {"id": q_counter_mcq, "type": "mcq", "question": line, "options": []}

        elif current_mode == 'g2':
            match_tf_opt = re.match(r'^(\*?)([a-d])[\.\)]\s*(.*)', line)
            if match_tf_opt:
                if not curr_q:
                    return jsonify({"success": False, "error": f"Lỗi tại ý Đ/S: '{line}'. Chưa có câu hỏi!"})
                is_correct = bool(match_tf_opt.group(1) == '*')
                curr_q['sub_items'].append({"key": match_tf_opt.group(2).lower(), "text": match_tf_opt.group(3), "correct": is_correct})
            else:
                if curr_q: tf_questions.append(curr_q)
                q_counter_tf += 1
                curr_q = {"id": q_counter_tf, "type": "tf", "question": line, "sub_items": []}

    if curr_q:
        if current_mode == 'g1': mcq_questions.append(curr_q)
        elif current_mode == 'g2': tf_questions.append(curr_q)

    for q in mcq_questions:
        num_c = sum(1 for opt in q['options'] if opt['correct'])
        if len(q['options']) < 2 or num_c != 1:
            return jsonify({"success": False, "error": f"Câu {q['id']} (Phần 1) phải có các phương án và đúng 1 dấu (#)!"})

    for q in tf_questions:
        if len(q['sub_items']) == 0:
            return jsonify({"success": False, "error": f"Câu {q['id']} (Phần 2) chưa có các ý hỏi con!"})

    exam_state["raw_bank"] = {"filename": filename, "mcq": mcq_questions, "tf": tf_questions}
    return jsonify({"success": True, "filename": filename, "total_mcq": len(mcq_questions), "total_tf": len(tf_questions)})

@app.route('/api/configure_exam', methods=['POST'])
def configure_exam():
    global submitted_ips
    submitted_ips.clear()
    exam_state["students"].clear()

    data = request.json
    n_mcq = data.get("num_mcq", 0)
    s_mcq = data.get("score_mcq", 0.0)
    n_tf = data.get("num_tf", 0)
    s_tf = data.get("score_tf", 0.0)
    dur = data.get("duration", 15)

    exam_state["class_name"] = data.get("class_name", "")
    exam_state["active_exam"].update({
        "num_mcq": n_mcq,
        "score_mcq": s_mcq,
        "score_per_mcq": round(s_mcq / max(1, n_mcq), 3),
        "num_tf": n_tf,
        "score_tf": s_tf,
        "score_per_tf": round(s_tf / max(1, n_tf), 3),
        "duration": dur,
        "tf_scale": data.get("tf_scale", {"1": 0.1, "2": 0.25, "3": 0.5, "4": 1.0}),
        "status": "ready"
    })
    socketio.emit('update_teacher_list', [])
    return jsonify({"success": True})

@app.route('/api/reset_exam', methods=['POST'])
def api_reset_exam():
    """API dọn dẹp sạch toàn bộ dữ liệu ca thi phục vụ nút 'Tạo mới'"""
    global submitted_ips
    submitted_ips.clear()
    exam_state["students"].clear()
    exam_state["class_name"] = ""
    exam_state["active_exam"] = {
        "num_mcq": 0, "score_mcq": 0.0, "score_per_mcq": 0.0,
        "num_tf": 0, "score_tf": 0.0, "score_per_tf": 0.0,
        "duration": 15, "tf_scale": {"1": 0.1, "2": 0.25, "3": 0.5, "4": 1.0},
        "status": "waiting", "start_time": None
    }
    socketio.emit('update_teacher_list', [])
    return jsonify({"success": True})

@app.route('/api/export_excel', methods=['GET'])
def export_excel():
    summary_path = os.path.join(SUBMISSION_DIR, "Bang_Diem_Tong_Hop.xlsx")
    if os.path.exists(summary_path):
        return send_file(summary_path, as_attachment=True, download_name="Bang_Diem_Tong_Hop.xlsx")
    update_summary_excel()
    return send_file(summary_path, as_attachment=True, download_name="Bang_Diem_Tong_Hop.xlsx")

# --- SOCKET.IO REALTIME EVENTS ---
@socketio.on('join_student')
def handle_student_join(data):
    sid = request.sid
    client_ip = request.remote_addr

    if client_ip in submitted_ips:
        emit('exam_blocked', {'message': 'Bạn đã thi rồi!'}, room=sid)
        return

    name = data.get('name', 'Thí sinh').strip() or 'Thí sinh'

    # Kiểm tra nếu học sinh này đã từng kết nối trước đó (bị rớt mạng vào lại)
    existing_sid = None
    for s_id, s_data in exam_state["students"].items():
        if s_data["name"] == name and s_data["ip"] == client_ip:
            existing_sid = s_id
            break

    time_left = exam_state["active_exam"]["duration"] * 60
    if exam_state["active_exam"]["status"] == "running" and exam_state["active_exam"]["start_time"]:
        elapsed = int(time.time() - exam_state["active_exam"]["start_time"])
        time_left = max(0, time_left - elapsed)

    if existing_sid:
        # Giữ lại bài thi và các câu đã làm của học sinh
        student_record = exam_state["students"].pop(existing_sid)
        student_record["time_left"] = time_left
        exam_state["students"][sid] = student_record
    else:
        exam_state["students"][sid] = {
            "ip": client_ip,
            "name": name,
            "time_left": time_left,
            "violations": 0,
            "submitted": "Đang thi",
            "score": 0.0,
            "answers": {},
            "exam": None
        }

    emit('update_teacher_list', list(exam_state["students"].values()), broadcast=True)

    # Nếu ca thi đang chạy, gửi đề ngay cho học sinh
    if exam_state["active_exam"]["status"] == "running":
        if not exam_state["students"][sid].get("exam"):
            client_payload, server_exam = generate_individual_exam()
            exam_state["students"][sid]["exam"] = server_exam
            client_payload["time_left"] = time_left
            emit('start_exam_now', client_payload, room=sid)
        else:
            # Gửi lại đề cũ cho em đó làm tiếp
            client_payload = {
                "mcq": [{"id": q["id"], "type": "mcq", "question": q["question"], "options": [{"key": o["key"], "text": o["text"]} for o in q["options"]]} for q in exam_state["students"][sid]["exam"]["mcq"]],
                "tf": [{"id": q["id"], "type": "tf", "question": q["question"], "sub_items": [{"key": s["key"], "text": s["text"]} for s in q["sub_items"]]} for q in exam_state["students"][sid]["exam"]["tf"]],
                "duration": exam_state["active_exam"]["duration"],
                "time_left": time_left
            }
            emit('start_exam_now', client_payload, room=sid)

@socketio.on('sync_answer')
def handle_sync_answer(data):
    sid = request.sid
    if sid in exam_state["students"]:
        key = data.get("key")
        val = data.get("val")
        if key:
            exam_state["students"][sid]["answers"][key] = val

@socketio.on('sync_timer')
def handle_timer_sync(data):
    sid = request.sid
    if sid in exam_state["students"]:
        st = exam_state["students"][sid]
        t_left = data.get("time_left", 0)
        st["time_left"] = t_left

        if t_left <= 0 and st["submitted"] != "Đã nộp":
            process_grading(sid, st.get("answers", {}), st.get("name"))
            return

        emit('update_teacher_list', list(exam_state["students"].values()), broadcast=True)

@socketio.on('report_violation')
def handle_violation():
    sid = request.sid
    if sid in exam_state["students"]:
        exam_state["students"][sid]["violations"] += 1
        emit('update_teacher_list', list(exam_state["students"].values()), broadcast=True)

@socketio.on('submit_exam')
def handle_submit(data):
    sid = request.sid
    if isinstance(data, dict) and 'answers' in data:
        process_grading(sid, data.get('answers', {}), data.get('name'))
    else:
        process_grading(sid, data if isinstance(data, dict) else {})

@socketio.on('teacher_start_exam')
def teacher_start():
    global submitted_ips
    submitted_ips.clear()

    exam_state["active_exam"]["status"] = "running"
    exam_state["active_exam"]["start_time"] = time.time()
    
    duration_secs = exam_state["active_exam"]["duration"] * 60
    for sid, st in exam_state["students"].items():
        client_payload, server_exam = generate_individual_exam()
        st["exam"] = server_exam
        client_payload["time_left"] = duration_secs
        emit('start_exam_now', client_payload, room=sid)

@socketio.on('reset_exam_session')
def handle_reset_session():
    global submitted_ips
    submitted_ips.clear()
    exam_state["students"].clear()
    exam_state["active_exam"]["status"] = "waiting"
    emit('update_teacher_list', [], broadcast=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    socketio.run(app, host='0.0.0.0', port=port, debug=False)
