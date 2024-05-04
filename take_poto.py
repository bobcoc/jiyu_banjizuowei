import cv2
import openpyxl
import imageio
import os
os.environ["PYTHONIOENCODING"] = "utf-8"

def load_students_info(excel_path):
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active
    students_info = []
    for row in sheet.iter_rows(min_row=2):
        exam_id = row[0].value
        name = row[1].value
        if exam_id and name:
            students_info.append((exam_id, name))
    return students_info

def capture_and_save_photos(students_info):
    cap = cv2.VideoCapture(0)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
    if not cap.isOpened():
        print("Error: Camera could not be opened.")
        return

    print("Press 'Space' to capture a photo, 'N' to retake the last photo, 'Q' to quit.")
    idx = 0
    while idx < len(students_info):
        exam_id, name = students_info[idx]
        cv2.namedWindow(f'Capture Photo for: {exam_id} {name}', cv2.WINDOW_NORMAL)

        while True:
            ret, frame = cap.read()
            if not ret:
                print("Failed to capture image; retrying...")
                continue

            cv2.imshow(f'Capture Photo for: {exam_id} {name}', frame)
            key = cv2.waitKey(100)  # Reduce delay to improve responsiveness

            if key == ord(' '):
                save_photo(exam_id, name, frame)
                idx += 1
                break
            elif key == ord('n') and idx > 0:
                save_photo(students_info[idx - 1][0], students_info[idx - 1][1], frame)
                # Do not increment idx, just retake the last photo
                break
            elif key == ord('q'):
                break

        if key == ord('q'):
            break
        cv2.destroyWindow(f'Capture Photo for: {exam_id} {name}')

    cap.release()
    cv2.destroyAllWindows()

def save_photo(exam_id, name, frame):
    photo_name = f"{exam_id}_{name}.png"
    encoded_photo_name = photo_name.encode('GBK', errors='ignore').decode('GBK')
    imageio.imwrite(encoded_photo_name, frame)
    print(f"Photo saved as {photo_name}")

excel_path = 'mt.xlsx'
students_info = load_students_info(excel_path)
capture_and_save_photos(students_info)