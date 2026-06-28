import cv2
import numpy as np
from PIL import Image
import os
from pathlib import Path

master_dir = Path(r"G:\Grad\College")
model_path = "face_detection_yunet_2023mar.onnx"

print(f"OpenCV version: {cv2.__version__}")
print(f"Model exists: {os.path.exists(model_path)}, size: {os.path.getsize(model_path):,} bytes\n")

# Find first 5 images
found = 0
for program_dir in master_dir.iterdir():
    if not program_dir.is_dir():
        continue
    for student_folder in program_dir.iterdir():
        if not student_folder.is_dir():
            continue
        for img_file in student_folder.glob("*.[jp][pn]g"):
            found += 1
            print(f"--- Testing {found}: {img_file.name} ---")
            
            img = Image.open(img_file).convert('RGB')
            img_w, img_h = img.size
            print(f"  Size: {img_w}x{img_h}")
            
            bgr = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            detector = cv2.FaceDetectorYN_create(model_path, "", (img_w, img_h), 0.3, 0.3, 5000)
            success, faces = detector.detect(bgr)
            print(f"  Success: {success}, Faces: {faces is not None and len(faces) or 0}")
            
            if faces is not None and len(faces) > 0:
                for face in faces:
                    x, y, w, h, score = face[0], face[1], face[2], face[3], face[14]
                    print(f"  Face: score={score:.3f}, pos=({x:.0f},{y:.0f}), size={w:.0f}x{h:.0f}")
            
            print()
            if found >= 5:
                break
        if found >= 5:
            break
    if found >= 5:
        break

if found == 0:
    print("No images found! Check your folder structure.")