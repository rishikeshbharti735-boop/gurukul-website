from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = "gurukul_secret_key"

# ── FOLDERS ──
UPLOAD_FOLDER = os.path.join('static', 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# ── EXCEL FILES ──
EXCEL_MARKS         = "marks.xlsx"
EXCEL_ADMISSION     = "admissions.xlsx"
EXCEL_REG           = "registrations.xlsx"
EXCEL_FEES_LEDGER   = "fees_ledger.xlsx"
EXCEL_FEE_STRUCTURE = "fee_structure.xlsx"
EXCEL_MARKSHEET     = "marksheet_data.xlsx"      # ← NEW: bulk marksheet data

# ── CREDENTIALS ──
ADMIN_USER, ADMIN_PASS = "admin", "gurukul@123"
FEE_USER,   FEE_PASS   = "accounts", "fee@786"

# ── SUBJECT MAP ──
SUBJECT_MAP = {
    "Nursery": ["English", "Hindi", "Maths", "Drawing"],
    "L.KG":    ["English", "Hindi", "Maths", "Drawing"],
    "U.KG":    ["English", "Hindi", "Maths", "Drawing"],
    "1st":  ["English", "Maths", "Hindi", "E.V.S.", "G.K.", "Computer"],
    "2nd":  ["Hindi", "English", "Maths", "E.V.S.", "G.K.", "Sanskrit", "Computer"],
    "3rd":  ["Hindi", "English", "Maths", "E.V.S.", "G.K.", "Sanskrit", "Computer"],
    "4th":  ["Hindi", "English", "Maths", "E.V.S.", "G.K.", "Sanskrit", "Computer"],
    "5th":  ["Hindi", "English", "Maths", "E.V.S.", "G.K.", "Sanskrit", "Computer"],
    "6th":  ["Hindi", "English", "Maths", "Science", "S.St", "Sanskrit", "Computer"],
    "7th":  ["Hindi", "English", "Maths", "Science", "S.St", "Sanskrit", "Computer"],
    "8th":  ["Hindi", "English", "Maths", "Science", "S.St", "Sanskrit", "Computer"],
    "9th":  ["Hindi", "English", "Maths", "Science", "S.St"],
    "10th": ["Hindi", "English", "Maths", "Science", "S.St"],
}

# ── HELPERS ──
def save_to_excel(data, filename):
    if os.path.exists(filename):
        df = pd.read_excel(filename)
    else:
        df = pd.DataFrame()
    df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
    df.to_excel(filename, index=False)

def is_admin():     return session.get('logged_in')
def is_fee_user():  return session.get('fee_logged_in') or session.get('logged_in')

# ══════════════════════════════
#  HOME
# ══════════════════════════════
@app.route('/')
def home():
    return render_template('home.html')

# ══════════════════════════════
#  ADMIN LOGIN & PORTAL
# ══════════════════════════════
@app.route('/admin_login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        if request.form.get('username') == ADMIN_USER and request.form.get('password') == ADMIN_PASS:
            session['logged_in'] = True
            return redirect(url_for('admin_portal'))
        flash("Admin Login Failed! Username ya Password galat hai.")
    return render_template('login.html')

@app.route('/admin_portal')
def admin_portal():
    if not is_admin():
        return redirect(url_for('admin_login'))
    m = pd.read_excel(EXCEL_MARKS).fillna('').to_dict('records')          if os.path.exists(EXCEL_MARKS)         else []
    a = pd.read_excel(EXCEL_ADMISSION).fillna('').to_dict('records')      if os.path.exists(EXCEL_ADMISSION)     else []
    r = pd.read_excel(EXCEL_REG).fillna('').to_dict('records')            if os.path.exists(EXCEL_REG)           else []
    f = pd.read_excel(EXCEL_FEES_LEDGER).fillna('').to_dict('records')    if os.path.exists(EXCEL_FEES_LEDGER)   else []
    ms = pd.read_excel(EXCEL_MARKSHEET).fillna('').to_dict('records')     if os.path.exists(EXCEL_MARKSHEET)     else []
    total_collection = sum(float(x.get('Paid Amount', 0) or 0) for x in f)
    return render_template('admin.html', marks=m, admissions=a, registrations=r,
                           fees=f, total_collection=total_collection, marksheet_count=len(ms))

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('home'))

# ══════════════════════════════
#  FEE LOGIN & DASHBOARD
# ══════════════════════════════
@app.route('/fee_login', methods=['GET', 'POST'])
def fee_login():
    if request.method == 'POST':
        u = request.form.get('username', '').strip()
        p = request.form.get('password', '').strip()
        if (u == FEE_USER and p == FEE_PASS) or (u == ADMIN_USER and p == ADMIN_PASS):
            session['fee_logged_in'] = True
            return redirect(url_for('fee_management'))
        flash("Galat ID ya Password! Dobara try karein.")
    return render_template('fee_login.html')

@app.route('/fee_logout')
def fee_logout():
    session.pop('fee_logged_in', None)
    return redirect(url_for('fee_login'))

@app.route('/fee_management', methods=['GET', 'POST'])
def fee_management():
    if not is_fee_user():
        return redirect(url_for('fee_login'))

    s_data, history, paid, f_struct, dues = None, [], 0.0, {}, 0.0
    reg_no = request.args.get('reg_no') or request.form.get('reg_no')

    if reg_no:
        reg_no = str(reg_no).strip()
        if os.path.exists(EXCEL_REG):
            df_r = pd.read_excel(EXCEL_REG).fillna('')
            match = df_r[df_r['Reg No'].astype(str).str.strip() == reg_no]
            if not match.empty:
                s_data = {k: str(v) for k, v in match.iloc[0].to_dict().items()}
                if os.path.exists(EXCEL_FEE_STRUCTURE):
                    df_fs = pd.read_excel(EXCEL_FEE_STRUCTURE).fillna(0)
                    sm = df_fs[df_fs['Reg No'].astype(str).str.strip() == reg_no]
                    if not sm.empty:
                        f_struct = sm.iloc[0].to_dict()
                if os.path.exists(EXCEL_FEES_LEDGER):
                    df_f = pd.read_excel(EXCEL_FEES_LEDGER).fillna(0)
                    hm = df_f[df_f['Reg No'].astype(str).str.strip() == reg_no]
                    history = hm.to_dict('records')
                    paid = float(pd.to_numeric(hm['Paid Amount'], errors='coerce').sum() or 0)
                total_yearly = float(pd.to_numeric(f_struct.get('Total Yearly', 0), errors='coerce') or 0)
                dues = max(0.0, total_yearly - paid)
            else:
                flash("Student record nahi mila! Reg No check karein.")

    return render_template('fee_page.html', s=s_data, history=history,
                           total=paid, f=f_struct, dues=dues, datetime=datetime)

@app.route('/update_fees', methods=['POST'])
def update_fees():
    if not is_fee_user():
        return redirect(url_for('fee_login'))
    reg_no = request.form.get('reg_no', '').strip()
    month  = request.form.get('month', '')
    amount = request.form.get('amount', 0)
    p_date = request.form.get('payment_date', '')
    try:
        amount = float(amount)
    except:
        amount = 0.0
    if reg_no and month and amount > 0:
        save_to_excel({"Reg No": reg_no, "Month": month, "Paid Amount": amount, "Date": p_date}, EXCEL_FEES_LEDGER)
        flash(f"✅ {month} ki ₹{int(amount)} fees save ho gayi!")
    else:
        flash("❌ Koi field khali hai ya amount galat hai!")
    return redirect(url_for('fee_management', reg_no=reg_no))

@app.route('/delete_fee_entry/<int:index>')
def delete_fee_entry(index):
    if not is_fee_user():
        return redirect(url_for('fee_login'))
    if os.path.exists(EXCEL_FEES_LEDGER):
        df = pd.read_excel(EXCEL_FEES_LEDGER)
        reg_no = str(df.iloc[index].get('Reg No', '')) if index < len(df) else ''
        df = df.drop(df.index[index])
        df.to_excel(EXCEL_FEES_LEDGER, index=False)
        flash("✅ Entry delete ho gayi!")
        if reg_no:
            return redirect(url_for('fee_management', reg_no=reg_no))
    return redirect(url_for('fee_management'))

@app.route('/save_fee_structure', methods=['POST'])
def save_fee_structure():
    if not is_admin():
        return redirect(url_for('admin_login'))
    reg_no = request.form.get('reg_no', '').strip()
    heads  = ["Tuition Fee", "Transport Fee", "Computer Fee", "Annual Charge", "Exam Fee"]
    data   = {"Reg No": reg_no}
    total  = 0.0
    for h in heads:
        key = h.lower().replace(' ', '_')
        val = float(request.form.get(key, 0) or 0)
        data[h] = val
        total += val
    data["Total Yearly"] = total
    if os.path.exists(EXCEL_FEE_STRUCTURE):
        df = pd.read_excel(EXCEL_FEE_STRUCTURE)
        df = df[df['Reg No'].astype(str).str.strip() != reg_no]
    else:
        df = pd.DataFrame()
    pd.concat([df, pd.DataFrame([data])], ignore_index=True).to_excel(EXCEL_FEE_STRUCTURE, index=False)
    flash(f"✅ Fee structure save ho gaya!")
    return redirect(url_for('admin_portal'))

# ══════════════════════════════
#  BILL / RECEIPT
# ══════════════════════════════
@app.route('/view_bill/<path:reg_no>')
def view_bill(reg_no):
    if not is_fee_user():
        return redirect(url_for('fee_login'))
    if not os.path.exists(EXCEL_REG):
        return "Registration file not found"
    df_reg = pd.read_excel(EXCEL_REG).fillna('')
    m = df_reg[df_reg['Reg No'].astype(str).str.strip() == str(reg_no).strip()]
    if m.empty:
        return "Student Not Found!"
    s = {k: str(v) for k, v in m.iloc[0].to_dict().items()}
    f_h, hist, paid = {}, [], 0.0
    if os.path.exists(EXCEL_FEE_STRUCTURE):
        df_fs = pd.read_excel(EXCEL_FEE_STRUCTURE).fillna(0)
        sm = df_fs[df_fs['Reg No'].astype(str).str.strip() == str(reg_no).strip()]
        if not sm.empty:
            f_h = sm.iloc[0].to_dict()
    if os.path.exists(EXCEL_FEES_LEDGER):
        df_l = pd.read_excel(EXCEL_FEES_LEDGER).fillna(0)
        lm = df_l[df_l['Reg No'].astype(str).str.strip() == str(reg_no).strip()]
        hist = lm.to_dict('records')
        paid = float(pd.to_numeric(lm['Paid Amount'], errors='coerce').sum() or 0)
    total_yearly = float(pd.to_numeric(f_h.get('Total Yearly', 0), errors='coerce') or 0)
    return render_template('bill_receipt.html', b={
        "name": s.get('Student Name', ''), "reg_no": reg_no,
        "class": s.get('Class', ''), "f": f_h, "history": hist,
        "total_paid": paid, "dues": max(0.0, total_yearly - paid),
        "date": datetime.now().strftime("%d-%m-%Y")
    })

# ══════════════════════════════
#  ADMISSION
# ══════════════════════════════
@app.route('/admission', methods=['GET', 'POST'])
def admission():
    if request.method == 'POST':
        df_a = pd.read_excel(EXCEL_ADMISSION) if os.path.exists(EXCEL_ADMISSION) else pd.DataFrame()
        next_id = len(df_a) + 1
        sid = f"GAA/ADM/2026/{next_id:03d}"
        photo    = request.files.get('student_photo')
        photo_fn = f"{sid.replace('/', '_')}.jpg" if (photo and photo.filename) else "default.jpg"
        if photo and photo.filename:
            photo.save(os.path.join(UPLOAD_FOLDER, photo_fn))
        data = {
            "Student ID": sid, "Date": request.form.get('admission_date'),
            "Student Name": request.form.get('student_name'), "DOB": request.form.get('dob'),
            "Father": request.form.get('father_name'), "Mother": request.form.get('mother_name'),
            "Mobile": request.form.get('whatsapp_no'), "Class": request.form.get('admission_class'),
            "Paid": request.form.get('paid_amount'), "Photo Path": photo_fn
        }
        save_to_excel(data, EXCEL_ADMISSION)
        return redirect(url_for('view_admission_receipt', index=len(df_a)))
    return render_template('admission.html')

@app.route('/view_admission_receipt/<int:index>')
def view_admission_receipt(index):
    df = pd.read_excel(EXCEL_ADMISSION).fillna('')
    s  = {k: str(v) for k, v in df.iloc[index].to_dict().items()}
    return render_template('admission_receipt.html', s=s)

# ══════════════════════════════
#  REGISTRATION
# ══════════════════════════════
@app.route('/get_admission_details/<path:adm_no>')
def get_admission_details(adm_no):
    if os.path.exists(EXCEL_ADMISSION):
        df    = pd.read_excel(EXCEL_ADMISSION).fillna('')
        match = df[df['Student ID'].astype(str).str.strip() == str(adm_no).strip()]
        if not match.empty:
            return jsonify({"success": True, "data": match.iloc[0].to_dict()})
    return jsonify({"success": False, "message": "ID Not Found"})

@app.route('/register', methods=['GET', 'POST'])
def registration():
    if request.method == 'POST':
        name   = request.form.get('student_name', '').strip()
        father = request.form.get('father', '').strip()
        dob    = request.form.get('dob', '')
        df_r = pd.read_excel(EXCEL_REG) if os.path.exists(EXCEL_REG) else pd.DataFrame()
        if not df_r.empty and 'Student Name' in df_r.columns:
            dup = df_r[
                (df_r['Student Name'].astype(str).str.lower().str.strip() == name.lower()) &
                (df_r['Father Name'].astype(str).str.lower().str.strip()  == father.lower()) &
                (df_r['DOB'].astype(str).str.strip() == str(dob))
            ]
            if not dup.empty:
                flash(f"⚠️ '{name}' pehle se registered hai! Duplicate entry block ki gayi.")
                return render_template('registration.html')
        next_num = 101 + len(df_r)
        data = {
            "Reg No": f"GAA/REG/2026/{next_num}", "Roll No": next_num,
            "Student Name": name, "Class": request.form.get('class', ''),
            "Section": request.form.get('section', ''), "Father Name": father,
            "Mother Name": request.form.get('mother_name', ''), "DOB": dob,
            "Gender": request.form.get('gender', ''), "Mobile No": request.form.get('mobile', ''),
            "Blood Group": request.form.get('blood_group', ''), "Address": request.form.get('address', ''),
        }
        save_to_excel(data, EXCEL_REG)
        return redirect(url_for('view_reg_receipt', index=len(df_r)))
    return render_template('registration.html')

@app.route('/view_reg_receipt/<int:index>')
def view_reg_receipt(index):
    df = pd.read_excel(EXCEL_REG).fillna('')
    s  = {k: str(v) for k, v in df.iloc[index].to_dict().items()}
    return render_template('reg_receipt.html', s=s)

# ══════════════════════════════
#  RESULT / MARKSHEET
# ══════════════════════════════
@app.route('/result_search')
def result_search():
    return render_template('result.html')

@app.route('/identify_class', methods=['POST'])
def identify_class():
    return render_template('verify_student.html', selected_class=request.form.get('class_name'))

@app.route('/view_result', methods=['POST'])
def view_result():
    u_class = request.form.get('class_name', '')
    u_name  = request.form.get('student_name', '').lower().strip()
    u_roll  = str(request.form.get('roll_no', '')).strip()

    # ── STEP 1: Pehle purane marks.xlsx mein search ──
    if os.path.exists(EXCEL_MARKS):
        df = pd.read_excel(EXCEL_MARKS).fillna(0)
        df.columns = df.columns.str.strip()
        match = df[
            (df['Class'].astype(str).str.strip() == u_class) &
            (df['Student Name'].astype(str).str.lower().str.strip() == u_name) &
            (df['Roll No'].astype(str).str.strip() == u_roll)
        ]
        if not match.empty:
            s    = match.iloc[0].to_dict()
            subs = SUBJECT_MAP.get(u_class, ["English", "Hindi", "Maths", "Science", "S.St"])
            marks_list, total_ob = [], 0
            for sub in subs:
                t = int(pd.to_numeric(s.get(sub, 0), errors='coerce') or 0)
                p = int(pd.to_numeric(s.get(f"{sub} Practical", 0), errors='coerce') or 0)
                marks_list.append({"name": sub, "t_ob": t, "p_ob": p, "tot_ob": t+p, "t_max": 80, "p_max": 20})
                total_ob += (t + p)
            max_total = len(subs) * 100
            s.update({
                "total_marks": total_ob, "max_total": max_total,
                "percent": f"{(total_ob/max_total*100) if max_total>0 else 0:.2f}%",
                "marks_list": marks_list
            })
            return render_template('marksheet.html', s=s)

    # ── STEP 2: Naye marksheet_data.xlsx mein search ──
    if os.path.exists(EXCEL_MARKSHEET):
        df2 = pd.read_excel(EXCEL_MARKSHEET).fillna('')
        df2.columns = df2.columns.str.strip()

        # Class mapping — naye system mein I,II,III... purane mein 1st,2nd,3rd...
        CLASS_MAP = {
            '1st':'I','2nd':'II','3rd':'III','4th':'IV','5th':'V',
            '6th':'VI','7th':'VII','8th':'VIII','9th':'IX','10th':'X',
            'I':'I','II':'II','III':'III','IV':'IV','V':'V',
            'VI':'VI','VII':'VII','VIII':'VIII','IX':'IX','X':'X'
        }
        mapped_class = CLASS_MAP.get(u_class, u_class)

        match2 = df2[
            (df2['class'].astype(str).str.strip().isin([u_class, mapped_class])) &
            (df2['student_name'].astype(str).str.lower().str.strip() == u_name) &
            (df2['roll_no'].astype(str).str.strip() == u_roll)
        ]

        if not match2.empty:
            s2 = match2.iloc[0].to_dict()

            # Naye format se marks_list banao
            NEW_SUBJECTS = [
                ('Maths',          'maths',    80, 20),
                ('Hindi',          'hindi',    80, 20),
                ('English',        'english',  80, 20),
                ('E.V.S./Science', 'science',  80, 20),
                ('Sanskrit',       'sanskrit', 80, 20),
                ('Computer',       'computer', 40, 10),
                ('G.K.',           'gk',       40, 10),
            ]
            marks_list, total_ob = [], 0
            for label, kt, maxT, maxP in NEW_SUBJECTS:
                t = int(pd.to_numeric(s2.get(kt+'_theory', 0), errors='coerce') or 0)
                p = int(pd.to_numeric(s2.get(kt+'_prac',   0), errors='coerce') or 0)
                marks_list.append({
                    "name": label, "t_ob": t, "p_ob": p,
                    "tot_ob": t+p, "t_max": maxT, "p_max": maxP
                })
                total_ob += (t + p)

            max_total = 600
            result = {
                "Student Name": s2.get('student_name', ''),
                "Father Name":  s2.get('father_name', ''),
                "Class":        s2.get('class', ''),
                "Roll No":      s2.get('roll_no', ''),
                "total_marks":  total_ob,
                "max_total":    max_total,
                "percent":      f"{(total_ob/max_total*100) if max_total>0 else 0:.2f}%",
                "marks_list":   marks_list,
                "drawing_grade": s2.get('drawing_grade', ''),
                "position":     s2.get('position', ''),
            }
            return render_template('marksheet.html', s=result)

    return "Result Not Found! Naam, Roll No ya Class check karein."

# ══════════════════════════════
#  ID CARD
# ══════════════════════════════
@app.route('/generate_id_card/<path:reg_no>')
def generate_id_card(reg_no):
    if not is_admin():
        return redirect(url_for('admin_login'))
    if not os.path.exists(EXCEL_REG):
        return "Registration file not found"
    df    = pd.read_excel(EXCEL_REG).fillna('')
    match = df[df['Reg No'].astype(str).str.strip() == str(reg_no).strip()]
    if not match.empty:
        s = {k: str(v) for k, v in match.iloc[0].to_dict().items()}
        photo = "default.jpg"
        if os.path.exists(EXCEL_ADMISSION):
            df_a    = pd.read_excel(EXCEL_ADMISSION).fillna('')
            a_match = df_a[
                df_a['Student Name'].astype(str).str.strip().str.lower() ==
                str(s.get('Student Name', '')).strip().lower()
            ]
            if not a_match.empty:
                fp = str(a_match.iloc[0].get('Photo Path', '') or '')
                if fp and fp.lower() not in ('nan', 'none', ''):
                    photo = fp
        s['Photo'] = photo
        return render_template('id_card.html', s=s)
    return "Student Not Found!"

# ══════════════════════════════
#  DELETE
# ══════════════════════════════
@app.route('/delete_entry/<string:category>/<int:index>')
def delete_entry(category, index):
    if not is_admin():
        return redirect(url_for('admin_login'))
    f_map = {"marks": EXCEL_MARKS, "admission": EXCEL_ADMISSION, "reg": EXCEL_REG}
    fname = f_map.get(category)
    if fname and os.path.exists(fname):
        df = pd.read_excel(fname)
        if index < len(df):
            df.drop(df.index[index]).to_excel(fname, index=False)
            flash("✅ Entry delete ho gayi!")
    return redirect(url_for('admin_portal'))

# ══════════════════════════════
#  STUDENT FEE CHECK (Public)
# ══════════════════════════════
@app.route('/student_fee_check')
@app.route('/studentfeecheck')
def student_fee_check():
    return render_template('studentfeecheck.html')

@app.route('/fee-details', methods=['POST'])
def fee_details():
    student_id = request.form.get('student_id', '').strip()
    dob        = request.form.get('dob', '').strip()
    result     = None
    error      = None
    if not os.path.exists(EXCEL_REG):
        error = "Records abhi available nahi hain. Office se contact karein."
        return render_template('studentfeecheck.html', error=error)
    df_r = pd.read_excel(EXCEL_REG).fillna('')
    match = df_r[df_r['Reg No'].astype(str).str.strip() == student_id]
    if match.empty:
        match = df_r[df_r['Roll No'].astype(str).str.strip() == student_id]
    if match.empty:
        error = "Student nahi mila! Admission No ya Roll No sahi se likhein."
    else:
        s = {k: str(v) for k, v in match.iloc[0].to_dict().items()}
        stored_dob = str(s.get('DOB', '')).strip()
        dob_match  = False
        try:
            input_dob_norm = datetime.strptime(dob, "%Y-%m-%d").strftime("%Y-%m-%d")
            for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"):
                try:
                    stored_dob_norm = datetime.strptime(stored_dob, fmt).strftime("%Y-%m-%d")
                    if input_dob_norm == stored_dob_norm:
                        dob_match = True
                    break
                except:
                    continue
        except:
            pass
        if not dob_match:
            error = "Date of Birth match nahi hua! Dobara check karein."
        else:
            f_struct, history, paid = {}, [], 0.0
            reg_no = s.get('Reg No', '')
            if os.path.exists(EXCEL_FEE_STRUCTURE):
                df_fs = pd.read_excel(EXCEL_FEE_STRUCTURE).fillna(0)
                sm = df_fs[df_fs['Reg No'].astype(str).str.strip() == reg_no]
                if not sm.empty:
                    f_struct = sm.iloc[0].to_dict()
            if os.path.exists(EXCEL_FEES_LEDGER):
                df_l = pd.read_excel(EXCEL_FEES_LEDGER).fillna(0)
                lm   = df_l[df_l['Reg No'].astype(str).str.strip() == reg_no]
                history = lm.to_dict('records')
                paid = float(pd.to_numeric(lm['Paid Amount'], errors='coerce').sum() or 0)
            total_yearly = float(pd.to_numeric(f_struct.get('Total Yearly', 0), errors='coerce') or 0)
            result = {
                "student": s, "f_struct": f_struct, "history": history,
                "paid": paid, "dues": max(0.0, total_yearly - paid), "total_yearly": total_yearly
            }
    return render_template('studentfeecheck.html', result=result, error=error, searched_id=student_id)

# ══════════════════════════════════════════════════════
#  MARKSHEET UPLOAD — Bulk Excel upload & Manual entry
# ══════════════════════════════════════════════════════
@app.route('/admin/marksheet-upload', methods=['GET', 'POST'])
def marksheet_upload():
    if not is_admin():
        return redirect(url_for('admin_login'))

    if request.method == 'POST':
        action = request.form.get('action', '')

        # ── MANUAL SAVE (one student) ──
        if action == 'manual_save':
            data = {
                "student_name":    request.form.get('student_name', '').strip(),
                "father_name":     request.form.get('father_name', '').strip(),
                "class":           request.form.get('class', '').strip(),
                "section":         request.form.get('section', '').strip(),
                "roll_no":         request.form.get('roll_no', '').strip(),
                "position":        request.form.get('position', '').strip(),
                "drawing_grade":   request.form.get('drawing_grade', '').strip(),
                "maths_theory":    request.form.get('maths_theory', ''),
                "maths_prac":      request.form.get('maths_prac', ''),
                "hindi_theory":    request.form.get('hindi_theory', ''),
                "hindi_prac":      request.form.get('hindi_prac', ''),
                "english_theory":  request.form.get('english_theory', ''),
                "english_prac":    request.form.get('english_prac', ''),
                "science_theory":  request.form.get('science_theory', ''),
                "science_prac":    request.form.get('science_prac', ''),
                "sanskrit_theory": request.form.get('sanskrit_theory', ''),
                "sanskrit_prac":   request.form.get('sanskrit_prac', ''),
                "computer_theory": request.form.get('computer_theory', ''),
                "computer_prac":   request.form.get('computer_prac', ''),
                "gk_theory":       request.form.get('gk_theory', ''),
                "gk_prac":         request.form.get('gk_prac', ''),
            }
            if not data['student_name'] or not data['class']:
                flash("❌ Student Name aur Class required hai!")
            else:
                save_to_excel(data, EXCEL_MARKSHEET)
                flash(f"✅ {data['student_name']} ka marksheet data save ho gaya!")
            return redirect(url_for('marksheet_upload'))

        # ── BULK EXCEL UPLOAD ──
        if action == 'bulk_upload':
            file = request.files.get('excel_file')
            if not file or not file.filename:
                flash("❌ Koi file select nahi ki!")
                return redirect(url_for('marksheet_upload'))
            try:
                df_new = pd.read_excel(file)
                if os.path.exists(EXCEL_MARKSHEET):
                    df_existing = pd.read_excel(EXCEL_MARKSHEET)
                    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                else:
                    df_combined = df_new
                df_combined.to_excel(EXCEL_MARKSHEET, index=False)
                flash(f"✅ {len(df_new)} students ka data successfully upload ho gaya!")
            except Exception as e:
                flash(f"❌ File upload error: {str(e)}")
            return redirect(url_for('marksheet_upload'))

        # ── DELETE RECORD ──
        if action == 'delete_record':
            idx = int(request.form.get('index', -1))
            if os.path.exists(EXCEL_MARKSHEET) and idx >= 0:
                df = pd.read_excel(EXCEL_MARKSHEET)
                if idx < len(df):
                    df.drop(df.index[idx]).to_excel(EXCEL_MARKSHEET, index=False)
                    flash("✅ Record delete ho gaya!")
            return redirect(url_for('marksheet_upload'))

        # ── CLEAR ALL ──
        if action == 'clear_all':
            if os.path.exists(EXCEL_MARKSHEET):
                os.remove(EXCEL_MARKSHEET)
                flash("✅ Saara marksheet data clear ho gaya!")
            return redirect(url_for('marksheet_upload'))

    # GET — load existing records
    records = []
    if os.path.exists(EXCEL_MARKSHEET):
        records = pd.read_excel(EXCEL_MARKSHEET).fillna('').to_dict('records')

    return render_template('marksheet_upload.html', records=records)


# ══════════════════════════════════════════════════
#  MARKSHEET PRINT — Class-wise bulk PDF print
# ══════════════════════════════════════════════════
@app.route('/admin/marksheet-print')
def marksheet_print():
    if not is_admin():
        return redirect(url_for('admin_login'))

    selected_class   = request.args.get('class', '')
    selected_section = request.args.get('section', '')

    records = []
    all_classes = []

    if os.path.exists(EXCEL_MARKSHEET):
        df = pd.read_excel(EXCEL_MARKSHEET).fillna('')
        # All distinct classes that have data
        all_classes = sorted(
            df['class'].astype(str).unique().tolist(),
            key=lambda c: ['I','II','III','IV','V','VI','VII','VIII','IX','X'].index(c)
                          if c in ['I','II','III','IV','V','VI','VII','VIII','IX','X'] else 99
        )
        if selected_class:
            filtered = df[df['class'].astype(str) == selected_class]
            if selected_section:
                filtered = filtered[filtered['section'].astype(str) == selected_section]
            records = filtered.sort_values('roll_no').fillna('').to_dict('records')

    return render_template('marksheet_print.html',
                           records=records,
                           all_classes=all_classes,
                           selected_class=selected_class,
                           selected_section=selected_section)




# ══════════════════════════════════════════════════
#  MARKSHEET PRINT ONLY — Clean PDF print page
# ══════════════════════════════════════════════════
@app.route('/admin/marksheet-printonly')
def marksheet_printonly():
    if not is_admin():
        return redirect(url_for('admin_login'))

    selected_class   = request.args.get('class', '')
    selected_section = request.args.get('section', '')
    records = []

    if os.path.exists(EXCEL_MARKSHEET):
        df = pd.read_excel(EXCEL_MARKSHEET).fillna('')
        if selected_class:
            filtered = df[df['class'].astype(str) == selected_class]
            if selected_section:
                filtered = filtered[filtered['section'].astype(str) == selected_section]
            records = filtered.sort_values('roll_no').fillna('').to_dict('records')

    return render_template('marksheet_printonly.html',
                           records=records,
                           selected_class=selected_class,
                           selected_section=selected_section)

# ══════════════════════════════
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=9000)