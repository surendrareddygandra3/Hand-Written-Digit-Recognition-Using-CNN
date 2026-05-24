import cv2
import os
import win32gui
import win32con
import pywintypes
from HandTracker import HandDetector
from dottedline import drawrect, drawline
import numpy as np
import win32com.client as win32  # For PowerPoint integration

# Set working directory
os.chdir(r"C:\Users\SURENDRA REDDY\OneDrive\Desktop\Surendra_profile\my_projects\Gesture-Driven-Presentation-Control-main\Gesture-Driven-Presentation-Control-main\code\final ppt.pptx")
# Function to convert PowerPoint slides to JPEG images
def convert_pptx_to_jpeg(input_file, output_folder):
    """
    Converts all slides in a PowerPoint presentation (pptx) to JPEG images.

    Args:
        input_file (str): Path to the input PowerPoint presentation file.
        output_folder (str): Path to the folder where the output JPEG images will be saved.
    """
    # Remove existing images from the output folder
    for filename in os.listdir(output_folder):
        if filename.endswith('.jpg'):
            os.remove(os.path.join(output_folder, filename))
    
    ppt = win32.Dispatch('PowerPoint.Application')
    presentation = ppt.Presentations.Open(input_file)

    for slide_number in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides(slide_number)
        output_filename = os.path.join(output_folder, f"slide_{slide_number}.jpg")
        slide.Export(output_filename, "JPG")

    presentation.Close()
    ppt.Quit()

# Example usage for converting PowerPoint to JPEG
input_file = r"C:\Users\SURENDRA REDDY\OneDrive\Desktop\Surendra_profile\my_projects\Gesture-Driven-Presentation-Control-main\Gesture-Driven-Presentation-Control-main\code\final ppt.pptx"
output_folder = r"C:\Users\SURENDRA REDDY\OneDrive\Desktop\Surendra_profile\my_projects\Gesture-Driven-Presentation-Control-main\Gesture-Driven-Presentation-Control-main\code\Images"
convert_pptx_to_jpeg(input_file, output_folder)

# Define variables for slide control and gesture settings
width, height = 1280, 720
frames_folder = "Images"
slide_num = 0
hs, ws = int(120 * 1.2), int(213 * 1.2)
ge_thresh_y = 400
ge_thresh_x = 750
gest_done = False
gest_counter = 0
delay = 15
annotations = [[]]
annot_num = 0
annot_start = False

# Get list of presentation images
path_imgs = sorted(os.listdir(frames_folder), key=len)
print(path_imgs)

# Camera Setup (Webcam)
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

# Initialize HandDetector
detector = HandDetector(detectionCon=0.8, maxHands=1)

# Main loop for detecting hand gestures and controlling slides
while True:
    # Get image frame
    success, frame = cap.read()
    frame = cv2.flip(frame, 1)
    pathFullImage = os.path.join(frames_folder, path_imgs[slide_num])
    slide_current = cv2.imread(pathFullImage)
    slide_current = cv2.resize(slide_current, (1280, 720)) 

    # Find the hand and its landmarks
    hands, frame = detector.findHands(frame)

    # Draw Gesture Threshold line
    drawrect(frame, (width, 0), (ge_thresh_x, ge_thresh_y), (0, 255, 0), 5, 'dotted')

    if hands and gest_done is False:  # If hand is detected
        hand = hands[0]
        cx, cy = hand["center"]
        lm_list = hand["lmList"]  # List of 21 Landmark points
        fingers = detector.fingersUp(hand)  

        # Constrain values for easier drawing
        x_val = int(np.interp(lm_list[8][0], [width//2, width], [0, width]))
        y_val = int(np.interp(lm_list[8][1], [150, height - 150], [0, height]))
        index_fing = x_val, y_val

        if cy < ge_thresh_y and cx > ge_thresh_x :  
            annot_start = False

            # gest_1 (previous)
            if fingers == [1, 0, 0, 0, 0]:
                annot_start = False
                if slide_num > 0:
                    gest_done = True
                    slide_num -= 1
                    annotations = [[]]
                    annot_num = 0

            # gest_2 (next)
            if fingers == [0, 0, 0, 0, 1]:
                annot_start = False
                if slide_num < len(path_imgs) - 1:
                    gest_done = True
                    slide_num += 1
                    annotations = [[]]
                    annot_num = 0
            
            # gest_3 (clear screen)
            if fingers == [1, 1, 1, 1, 1]:
                if annotations:
                    annot_start = False
                    if annot_num >= 0:
                        annotations.clear()
                        annot_num = 0
                        gest_done = True 
                        annotations = [[]]

        # gest_4 (show pointer)
        if fingers == [0, 1, 0, 0, 0]:
            cv2.circle(slide_current, index_fing, 4, (0, 0, 255), cv2.FILLED)
            annot_start = False

        # gest_5 (draw)
        if fingers == [0, 1, 1, 0, 0]:
            if annot_start is False:
                annot_start = True
                annot_num += 1
                annotations.append([])

            annotations[annot_num].append(index_fing)
            cv2.circle(slide_current, index_fing, 4, (0, 0, 255), cv2.FILLED)

        else:
            annot_start = False

        # gest_6 (erase)
        if fingers == [0, 1, 1, 1, 0]:
            if annotations:
                annot_start = False
                if annot_num >= 0:
                    annotations.pop(-1)
                    annot_num -= 1
                    gest_done = True 

    else:
        annot_start = False

    # Gesture Performed Iterations
    if gest_done:
        gest_counter += 1
        if gest_counter > delay:
            gest_counter = 0
            gest_done = False

    for i, annotation in enumerate(annotations):
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(slide_current, annotation[j - 1], annotation[j], (0, 0, 255), 6)

    # Adding webcam feed to the slide
    img_small = cv2.resize(frame, (ws, hs))
    h, w, _ = slide_current.shape   
    slide_current[h-hs:h, w-ws:w] = img_small

    # Create a named window and set it to fullscreen
    cv2.namedWindow("Slides", cv2.WINDOW_NORMAL)
    cv2.setWindowProperty("Slides", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

    # Show the slide with annotations
    hwnd = win32gui.FindWindow(None, "Slides")

    # Check if the window handle is valid
    if hwnd:
    # Bring the window to the foreground
        win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)  # Ensure the window is not minimized
        try:
            win32gui.SetForegroundWindow(hwnd)  # Bring the window to the foreground
        except pywintypes.error as e:
            print(f"Error setting the window to foreground: {e}")
    else:
        print("Failed to find the 'Slides' window. Ensure the name matches.")

    # Show the slide with annotations
    cv2.imshow("Slides", slide_current)

    # Check for key press to exit
    key = cv2.waitKey(1)
    if key == ord('q'):  # Press 'q' to exit
        break

# Release the camera and close all OpenCV windows   
cap.release()
cv2.destroyAllWindows()
