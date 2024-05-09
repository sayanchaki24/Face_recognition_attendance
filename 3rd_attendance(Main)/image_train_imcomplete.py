import cv2
import face_recognition
import os
import datetime
import pandas as pd

# Function to detect faces in a frame
def detect_faces(frame):
    # Convert frame to RGB format (expects BGR)
    rgb_frame = frame[:, :, ::-1]

    # Find all faces in the frame using Haar cascade classifier
    face_locations = face_recognition.face_locations(rgb_frame)
    return face_locations

# Function to encode known faces and create a dictionary
def encode_known_faces(known_faces_dir):
    known_faces_encodings = {}

    # Loop through image files in the known faces directory
    for filename in os.listdir("C:\\GCU\\Pyhton programs\\Project\\random_attendance\\known_faces"):
        if filename.endswith((".jpg", ".jpeg", ".png")):
            # Load image
            image = cv2.imread(os.path.join(known_faces_dir, filename))
            rgb_image = image[:, :, ::-1]  # Convert BGR to RGB

            # Get face encoding
            face_encoding = face_recognition.face_encodings(rgb_image)[0]

            # Extract name from filename (assuming format: name.jpg)
            name = os.path.splitext(filename)[0]

            # Add name and encoding to dictionary
            known_faces_encodings[name] = face_encoding

    return known_faces_encodings

# Function to mark attendance and create/update an Excel file
def mark_attendance(name, timestamp):
    # Create or load existing attendance sheet
    attendance_sheet_path = "attendance.xlsx"
    if not os.path.exists(attendance_sheet_path):
        data = {"Name": [], "Timestamp": []}
        df = pd.DataFrame(data)
        df.to_excel(attendance_sheet_path, index=False)
    else:
        df = pd.read_excel(attendance_sheet_path)

    # Append new attendance record
    new_row = {"Name": name, "Timestamp": timestamp}
    df = df.append(new_row, ignore_index=True)

    # Save updated attendance sheet
    df.to_excel(attendance_sheet_path, index=False)

# Main program
KNOWN_FACES_DIR = "known_faces"  # Path to directory containing known faces
VIDEO_INPUT = 0  # Change if using a video file instead of webcam

# Load known face encodings
known_face_encodings = encode_known_faces("C:\\GCU\\Pyhton programs\\Project\\random_attendance\\known_faces")

# Initialize video capture
video_capture = cv2.VideoCapture(VIDEO_INPUT)

while True:
    # Capture frame-by-frame
    ret, frame = video_capture.read()

    # Detect faces
    face_locations = detect_faces(frame)

    # Process each detected face
    for (top, right, bottom, left) in face_locations:
        # Extract the facial ROI
        face_image = frame[top:bottom, left:right]
        rgb_face_image = face_image[:, :, ::-1]

        # Try face recognition
        matches = face_recognition.compare_faces(
            list(known_face_encodings.values()), face_recognition.face_encodings(rgb_face_image)[0]
        )
        name = "Unknown"

        # If a match is found, get the name
        if True in matches:
            first_match_index = matches.index(True)
            name = list(known_face_encodings.keys())[first_match_index]

            # Mark attendance (if name is known)
            if name != "Unknown":
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                mark_attendance(name, timestamp)

        # Draw facial recognition box and label
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)
