from ultralytics import YOLO
import cv2
import time
from datetime import datetime
import os
import sys 
from google.cloud import storage 

# =======================================================
# ⚙️ KONFIGURASI FIREBASE & KAMERA MULTI
# =======================================================
SERVICE_ACCOUNT_FILE = 'credensial.json' 

# ✅ PERBAIKAN BUCKET FINAL: DITAMBAH '.appspot.com'
BUCKET_NAME = 'deteksigulma-ab7b3.appspot.com' 

# 🚨 DEFINISI KAMERA IP ANDA
CAMERA_SOURCES = {
    'CAM_1': 'rtsp://camera1:camera1@192.168.1.102:554/stream1',
    'CAM_2': 'rtsp://camera2:camera2@192.168.1.101:554/stream1'
}

SAVE_DIR = os.path.join(os.path.dirname(__file__), "gulma_capture")
os.makedirs(SAVE_DIR, exist_ok=True)

# 💡 THRESHOLD: Nilai ini bisa Anda sesuaikan antara 0.01 (longgar) hingga 0.99 (ketat)
YOLO_CONFIDENCE_THRESHOLD = 0.5 
# =======================================================


# =======================================================
# FUNGSI FIREBASE (TIDAK BERUBAH)
# =======================================================
def get_storage_client():
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        svc_path = os.path.join(script_dir, SERVICE_ACCOUNT_FILE)
        if not os.path.isfile(svc_path):
            raise FileNotFoundError(f"File service account tidak ditemukan: {svc_path}")
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = svc_path
        return storage.Client()
    except Exception as e:
        raise Exception(f"Gagal inisialisasi Storage Client: {str(e)}")

def test_storage_connection():
    try:
        client = get_storage_client()
        bucket = client.bucket(BUCKET_NAME)
        bucket.get_blob('nonexistent-test-file-999') 
        return True
    except Exception as e:
        if '404' not in str(e):
             print(f"❌ ERROR: Koneksi ke Firebase Storage gagal: {str(e)}")
             return False
        return True

def upload_file_to_storage(local_file_path: str, destination_blob_name: str) -> tuple[bool, str]:
    try:
        client = get_storage_client()
        bucket = client.bucket(BUCKET_NAME)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_filename = os.path.basename(local_file_path)
        
        destination_blob_name = f"{destination_blob_name.split('/')[0]}/{timestamp}_{base_filename}"
            
        blob = bucket.blob(destination_blob_name)
        blob.upload_from_filename(local_file_path)
        
        url = f"https://storage.googleapis.com/{BUCKET_NAME}/{destination_blob_name}"
        
        os.remove(local_file_path) 
        
        print(f"✅ Upload berhasil!\n   File: {destination_blob_name}\n   URL: {url}")
        return True, url
        
    except Exception as e:
        print(f"❌ ERROR: Gagal mengupload file: {str(e)}")
        return False, ""


# =======================================================
# 🎯 FUNGSI CAPTURE & DETEKSI MULTI-KAMERA (DENGAN THRESHOLD)
# =======================================================
def capture_and_detect(cap, folder_name: str, model_yolo):
    """Ambil gambar, deteksi, dan unggah untuk satu kamera."""
    ret, frame = cap.read()
    if not ret:
        print(f"❌ Gagal ambil gambar dari kamera ({folder_name}). Mungkin stream mati atau koneksi putus.")
        return

    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    image_to_save = frame 
    local_filename = f"{folder_name}_raw_{now}.jpg" 

    # Deteksi hanya jika model dimuat
    if model_yolo:
        try:
            # 🎯 MENERAPKAN CONFIDENCE THRESHOLD
            results = model_yolo(frame, conf=YOLO_CONFIDENCE_THRESHOLD) 
            
            image_to_save = results[0].plot() 
            
            # Cek apakah ada deteksi yang LOLOS THRESHOLD (boxes tidak kosong)
            if len(results[0].boxes) > 0:
                local_filename = f"{folder_name}_detected_{now}.jpg" 
                print(f"   [YOLO] Deteksi berhasil: {len(results[0].boxes)} objek (Conf >= {YOLO_CONFIDENCE_THRESHOLD}).")
            else:
                # Jika tidak ada deteksi yang lolos threshold, kita simpan sebagai raw.
                 local_filename = f"{folder_name}_raw_{now}.jpg" 
                 
        except Exception as e:
             print(f"⚠️ Peringatan: Deteksi model gagal untuk {folder_name}: {e}. Menggunakan gambar mentah.")

    detect_path = os.path.join(SAVE_DIR, local_filename)
    cv2.imwrite(detect_path, image_to_save)
    print(f"\n[{folder_name}] 🔎 Gambar disimpan lokal: {detect_path}")
    
    DESTINATION_BLOB = f"{folder_name}/{local_filename}"
    
    upload_file_to_storage(detect_path, DESTINATION_BLOB)


# =======================================================
# 🔁 LOOP UTAMA 
# =======================================================

# 1. Test koneksi Firebase Storage (WAJIB)
print("\n🔄 Mengetes koneksi Firebase Storage...")
if not test_storage_connection():
    print("❌ Tidak bisa melanjutkan tanpa koneksi ke Firebase Storage.")
    sys.exit(1)
print(f"✅ Koneksi Firebase Storage berhasil! (Target: {BUCKET_NAME})\n")


# 2. Load model YOLO
model = None 
try:
    model_path = os.path.join(os.path.dirname(__file__), "best.pt")
    model = YOLO(model_path)
    print(f"✅ Model YOLO '{os.path.basename(model_path)}' berhasil dimuat.")
except Exception as e:
    print(f"⚠️ Peringatan: Gagal memuat model YOLO. Unggahan akan menggunakan gambar mentah. Error: {e}")


# 3. Inisialisasi Semua Kamera
caps = {}
print("\n🔄 Inisialisasi Kamera...")
for name, url in CAMERA_SOURCES.items():
    # Menggunakan cv2.CAP_FFMPEG untuk stabilitas RTSP
    cap = cv2.VideoCapture(url, cv2.CAP_FFMPEG) 
    
    if not cap.isOpened():
        print(f"❌ Gagal inisialisasi Kamera {name} dari sumber: {url}")
        continue
        
    # ✅ KODE STABILISASI STREAM: Membersihkan buffer RTSP
    print(f"   [Stabilisasi] Menunggu {name} stabil (membersihkan buffer 5 frame)...")
    for _ in range(5):
        cap.read()

    caps[name] = cap
    print(f"✅ Kamera {name} berhasil diinisialisasi.")

if not caps:
    print("❌ Tidak ada kamera yang berhasil diinisialisasi. Program berhenti.")
    sys.exit(1)


# 4. Loop Utama
INTERVAL_SECONDS = 5
print(f"\n--- PROGRAM DETEKSI/CAPTURE GULMA AKTIF (HEADLESS) ---")
print(f"Mengambil dari kamera: {list(caps.keys())}. Unggah setiap {INTERVAL_SECONDS} detik.")

try:
    while True:
        start_time = time.time()
        
        for folder_name, cap in caps.items():
            capture_and_detect(cap, folder_name, model)
        
        elapsed_time = time.time() - start_time
        
        time_to_sleep = INTERVAL_SECONDS - elapsed_time
        if time_to_sleep > 0:
            print(f"Menunggu {time_to_sleep:.2f} detik untuk loop berikutnya...")
            time.sleep(time_to_sleep)
            
except KeyboardInterrupt:
    pass 
finally:
    for cap in caps.values():
        cap.release()
    print("\nProgram Selesai. Semua kamera dilepaskan.")