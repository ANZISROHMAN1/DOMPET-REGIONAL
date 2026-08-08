import os
import json
import base64
import urllib.request
import urllib.parse
import uuid
from werkzeug.utils import secure_filename
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_from_directory
import sheets_helper

app = Flask(__name__)
app.secret_key = 'super_secret_key_klaim_id' # For flash messages

app.config['UPLOAD_FOLDER'] = os.path.join('static', 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'pdf'} # Support local files

@app.route('/manifest.json')
def manifest():
    return send_from_directory('static', 'manifest.json')

@app.route('/sw.js')
def service_worker():
    return send_from_directory('static', 'sw.js', mimetype='application/javascript')

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        nama = request.form.get('nama')
        kegiatan = request.form.get('kegiatan')
        nominal = request.form.get('nominal')
        bank = request.form.get('bank')
        rekening = request.form.get('rekening')
        
        if 'bukti' not in request.files:
            flash('Tidak ada file bukti yang diupload', 'error')
            return redirect(request.url)
            
        file = request.files['bukti']
        if file.filename == '':
            flash('Tidak ada file yang dipilih', 'error')
            return redirect(request.url)
            
        if file and allowed_file(file.filename):
            # Save locally
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4().hex}_{filename}"
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], unique_filename))
            bukti_path = unique_filename
            
            # Save to SQLite database
            import database
            conn = database.get_db_connection()
            cursor = conn.cursor()
            cursor.execute('INSERT INTO reimbursements (nama, kegiatan, nominal, bank, rekening, bukti_path) VALUES (?, ?, ?, ?, ?, ?)',
                         (nama, kegiatan, float(nominal), bank, rekening, bukti_path))
            new_id = cursor.lastrowid
            conn.commit()
            conn.close()
            
            # Save to Google Sheets
            try:
                sheets_helper.append_row(nama, kegiatan, float(nominal), bank, rekening, 'Pending', bukti_path)
            except Exception as e:
                print(f"Failed to save to Google Sheets: {e}")
            
            flash(f'Pengajuan berhasil! Mohon simpan ID Anda.', 'success')
            return redirect(url_for('sukses', id=new_id))
        else:
            flash('Tipe file tidak diizinkan. Mohon gunakan JPG/PNG/PDF.', 'error')
            return redirect(request.url)

    return render_template('index.html')

@app.route('/sukses/<int:id>')
def sukses(id):
    return render_template('sukses.html', id=id)

@app.route('/cek-status', methods=['GET', 'POST'])
def cek_status():
    pengajuan = None
    if request.method == 'POST':
        try:
            req_id = int(request.form.get('id_pengajuan').replace('#', '').strip())
            import database
            conn = database.get_db_connection()
            pengajuan = conn.execute('SELECT * FROM reimbursements WHERE id = ?', (req_id,)).fetchone()
            conn.close()
            
            if not pengajuan:
                flash('ID Pengajuan tidak ditemukan.', 'error')
        except ValueError:
            flash('Format ID Pengajuan tidak valid.', 'error')
            
    return render_template('cek_status.html', pengajuan=pengajuan)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        
        # Hardcoded credentials for MVP
        if username == 'admin' and password == 'admin123':
            session['admin_logged_in'] = True
            flash('Login berhasil!', 'success')
            return redirect(url_for('admin'))
        else:
            flash('Username atau password salah', 'error')
            
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('admin_logged_in', None)
    flash('Anda telah logout.', 'success')
    return redirect(url_for('login'))

@app.route('/admin')
def admin():
    if not session.get('admin_logged_in'):
        return redirect(url_for('login'))
        
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    
    import database
    conn = database.get_db_connection()
    
    if start_date and end_date:
        reimbursements = conn.execute('SELECT * FROM reimbursements WHERE date(tanggal) >= date(?) AND date(tanggal) <= date(?) ORDER BY id DESC', (start_date, end_date)).fetchall()
    else:
        reimbursements = conn.execute('SELECT * FROM reimbursements ORDER BY id DESC').fetchall()
    
    # Calculate stats
    total_pengajuan = len(reimbursements)
    total_pending = sum(1 for r in reimbursements if r['status'] == 'Pending')
    total_approved = sum(1 for r in reimbursements if r['status'] == 'Approved')
    
    conn.close()
    
    return render_template('admin.html', 
                           reimbursements=reimbursements,
                           total_pengajuan=total_pengajuan,
                           total_pending=total_pending,
                           total_approved=total_approved,
                           start_date=start_date,
                           end_date=end_date)

@app.route('/admin/update/<int:id>', methods=['POST'])
def update_status(id):
    if not session.get('admin_logged_in'):
        return redirect(url_for('login'))
        
    import database
    new_status = request.form.get('status')
    
    bukti_tf_path = None
    if new_status == 'Approved' and 'bukti_tf' in request.files:
        file = request.files['bukti_tf']
        if file and file.filename != '' and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4().hex}_tf_{filename}"
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], unique_filename))
            bukti_tf_path = unique_filename
            
    if new_status in ['Approved', 'Rejected']:
        conn = database.get_db_connection()
        if bukti_tf_path:
            conn.execute('UPDATE reimbursements SET status = ?, bukti_tf_path = ? WHERE id = ?', (new_status, bukti_tf_path, id))
        else:
            conn.execute('UPDATE reimbursements SET status = ? WHERE id = ?', (new_status, id))
        conn.commit()
        conn.close()
        flash(f'Status pengajuan #{id} berhasil diupdate menjadi {new_status}.', 'success')
    return redirect(url_for('admin'))


# Allow external connections by setting host to 0.0.0.0
if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=5050)
