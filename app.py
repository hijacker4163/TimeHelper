import json
import os
from flask import Flask, render_template, request, jsonify, send_from_directory

app = Flask(__name__)
DATA_FILE = os.path.join(os.path.dirname(__file__), 'records.json')

def load_records():
    """讀取紀錄檔案 records.json，若檔案不存在或損壞則返回空陣列"""
    if not os.path.exists(DATA_FILE):
        return []
    try:
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading records: {e}")
        return []

def save_records(records):
    """將歷史紀錄寫入 records.json 本機檔案"""
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Error saving records: {e}")
        return False

@app.route('/')
def index():
    """渲染主頁面"""
    return render_template('index.html')

@app.route('/api/records', methods=['GET'])
def get_records():
    """獲取所有歷史紀錄"""
    records = load_records()
    return jsonify({'status': 'success', 'data': records})

@app.route('/api/records', methods=['POST'])
def add_record():
    """新增一筆歷史紀錄並儲存到 JSON 檔案"""
    data = request.json
    if not data:
        return jsonify({'status': 'error', 'message': '無效的資料'}), 400
    
    records = load_records()
    records.append(data)
    save_records(records)
    return jsonify({'status': 'success', 'data': data})

@app.route('/api/records/<int:record_id>', methods=['DELETE'])
def delete_record(record_id):
    """刪除單筆指定 ID 的歷史紀錄"""
    records = load_records()
    new_records = [r for r in records if r.get('id') != record_id]
    save_records(new_records)
    return jsonify({'status': 'success', 'deleted_id': record_id})

@app.route('/api/records/<int:record_id>/task', methods=['PUT'])
def update_record_task(record_id):
    """更新指定 ID 歷史紀錄的備忘事項（其他資料欄位維持不變）"""
    data = request.json or {}
    task_text = data.get('task', '').strip()
    records = load_records()
    updated = False
    for r in records:
        if r.get('id') == record_id:
            r['task'] = task_text
            updated = True
            break
    if updated:
        save_records(records)
        return jsonify({'status': 'success', 'task': task_text})
    return jsonify({'status': 'error', 'message': '找不到該筆紀錄'}), 404

@app.route('/api/records', methods=['DELETE'])
def clear_all_records():
    """清空歷史紀錄 (依使用者需求調整為只刪除已到期紀錄，保留進行中倒數)"""
    import time
    now_ms = time.time() * 1000
    records = load_records()
    new_records = [r for r in records if r.get('targetTimestamp', 0) > now_ms]
    save_records(new_records)
    return jsonify({'status': 'success'})

@app.route('/api/records/<int:record_id>/notify', methods=['PUT'])
def mark_notified(record_id):
    """將紀錄標記為已推播"""
    records = load_records()
    updated = False
    for r in records:
        if r.get('id') == record_id:
            r['notified'] = True
            updated = True
            break
    if updated:
        save_records(records)
    return jsonify({'status': 'success', 'updated': updated})

@app.route('/sw.js')
def serve_sw():
    """提供 Service Worker 檔案供離線與行動端推播支援"""
    return send_from_directory(os.path.dirname(__file__), 'sw.js', mimetype='application/javascript')

@app.route('/manifest.json')
def serve_manifest():
    """提供 PWA Manifest 供手機加入主畫面"""
    return send_from_directory(os.path.dirname(__file__), 'manifest.json', mimetype='application/json')

if __name__ == '__main__':
    # 按照需求指定埠號 5001
    print("啟動 剩餘時間計算器 服務，Port: 5001...")
    app.run(host='0.0.0.0', port=5001, debug=True)
