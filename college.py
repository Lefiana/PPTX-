import copy
import os
import io
from pathlib import Path
from pptx import Presentation
from pptx.util import Cm
from PIL import Image, ImageOps

# ==========================================
# --- CONFIGURATION & PATHS ----------------
# ==========================================
INPUT_PPTX = "Commencement Exercises - PPT Template.pptx.pptx"
OUTPUT_PPTX = "final_graduation_slides_college.pptx"
MASTER_STUDENT_DIR = Path(r"G:\Grad\College")

TEMPLATE_SLIDE_INDEX = 0 

COURSE_MAP = {
    "BACOMM": "Bachelor of Arts in Communication",
    "BSCPE": "Bachelor of Science in Computer Engineering",
    "BSAIS": "Bachelor of Science in Accounting Information Systems",
    "BSIT": "Bachelor of Science in Information Technology",
    "BMMA": "Bachelor of Multimedia Arts",
    "BSHM": "Bachelor of Science in Hospitality Management",
    "BSTM": "Bachelor of Science in Tourism Management",
}

# --- PORTRAIT POSITIONING (Centimeters) ---
IMG_LEFT = Cm(29.79)   
IMG_TOP = Cm(2.58)     

# Absolute frame constraints
IMG_WIDTH = Cm(17.27)  
IMG_HEIGHT = Cm(23.43) 

# ==========================================
# --- TARGET SHAPE NAMES (From inspect_slide.py) ---
# ==========================================
# Replace these with the exact shape names from your diagnostic script!
TARGET_SURNAME_SHAPE = "TextBox 6"
TARGET_FIRSTNAME_SHAPE = "TextBox 7"
TARGET_COURSE_SHAPE = "TextBox 8"


def get_sanitized_image(image_path):
    with Image.open(image_path) as img:
        # This fixes the orientation based on EXIF data
        img = ImageOps.exif_transpose(img) 
        rgb_img = img.convert('RGB')
        img_stream = io.BytesIO()
        rgb_img.save(img_stream, format="JPEG")
        img_stream.seek(0)
        return img_stream


def duplicate_slide(prs, index):
    """
    Standard-compliant duplication: Clones the slide and 
    automatically handles relationship mapping.
    """
    source_slide = prs.slides[index]
    # Create the new slide based on the same layout as the source
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    
    # 1. Clear out the default placeholder content from the new slide
    for shp in list(new_slide.shapes):
        shp.element.getparent().remove(shp.element)
        
    # 2. Deep copy the elements from the source to the new slide
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.append(new_el)
            
    # 3. Automatically map relationships (images, etc.) 
    # By iterating through the source slide's relationships, 
    # we link the new slide to the same images properly.
    for rel in source_slide.part.rels.values():
        if "notesSlide" in rel.reltype:
            continue
        new_slide.part.relate_to(rel.target_part, rel.reltype)
            
    return new_slide


def replace_text_preserve_format(shape, new_text):
    if not shape.has_text_frame:
        return
    text_frame = shape.text_frame
    for paragraph in text_frame.paragraphs:
        if not paragraph.runs:
            continue
        p_font = paragraph.runs[0].font
        paragraph.text = "" 
        new_run = paragraph.add_run()
        new_run.text = new_text
        if p_font:
            new_run.font.name = p_font.name
            new_run.font.size = p_font.size
            new_run.font.color.rgb = p_font.color.rgb
            new_run.font.bold = p_font.bold
            new_run.font.italic = p_font.italic


def get_sorted_student_data(master_dir: Path):
    student_data = []
    for course_folder in master_dir.iterdir():
        if not course_folder.is_dir():
            continue
        for student_folder in course_folder.iterdir():
            if not student_folder.is_dir():
                continue
            parts = student_folder.name.split('_')
            if len(parts) >= 4:
                try:
                    seq_number = int(parts[0])
                    surname = parts[1].upper()
                    first_name = " ".join(parts[2:-1]).title() 
                    course_acronym = parts[-1].upper()
                    full_course_title = COURSE_MAP.get(course_acronym, course_acronym)
                    
                    image_path = None
                    for ext in ['*.jpg', '*.jpeg', '*.png']:
                        imgs = list(student_folder.glob(ext))
                        if imgs:
                            image_path = imgs[0]
                            break
                    if image_path:
                        student_data.append({
                            "sequence": seq_number,
                            "surname": surname,
                            "firstname": first_name,
                            "course": full_course_title,
                            "image_path": image_path
                        })
                except ValueError:
                    continue
    return sorted(student_data, key=lambda x: (x["course"], x["surname"]))


def update_slide(slide, student):
    print(f"Assembling Slide {student['sequence']}: {student['surname']}")
    
    # 1. Update text fields (No cleanup needed, as the background is locked in Master)
    for shape in slide.shapes:
        if shape.name == TARGET_SURNAME_SHAPE:
            replace_text_preserve_format(shape, student["surname"])
        elif shape.name == TARGET_FIRSTNAME_SHAPE:
            replace_text_preserve_format(shape, student["firstname"])
        elif shape.name == TARGET_COURSE_SHAPE:
            replace_text_preserve_format(shape, student["course"])

    # 2. Add the portrait on the "clean" master-backed slide
    clean_image_stream = get_sanitized_image(student["image_path"])
    slide.shapes.add_picture(
        clean_image_stream, 
        IMG_LEFT, 
        IMG_TOP, 
        width=IMG_WIDTH, 
        height=IMG_HEIGHT
    )


def main():
    print("--- PHASE 0: Pre-flight check ---")
    prs = Presentation(INPUT_PPTX)
    students = get_sorted_student_data(MASTER_STUDENT_DIR)
    
    student_count = len(students)
    slide_count = len(prs.slides)
    
    print(f"Total Student Records Found: {student_count}")
    print(f"Initial Presentation Slides: {slide_count}")
    
    gap = student_count - slide_count
    
    print("\n--- PHASE 1: Slide Duplication Engine ---")
    if gap > 0:
        print(f"Manufacturing {gap} layout frames safely...")
        for _ in range(gap):
            duplicate_slide(prs, TEMPLATE_SLIDE_INDEX)
        print("Clones generated successfully without XML corruption.")
    
    all_slides = list(prs.slides)

    print("\n--- PHASE 2: Production Update Engine ---")
    for slide, student in zip(all_slides, students):
        update_slide(slide, student)
        
    prs.save(OUTPUT_PPTX)
    print(f"\nSuccess! Perfected slides saved to: {OUTPUT_PPTX}")

if __name__ == "__main__":
    main()