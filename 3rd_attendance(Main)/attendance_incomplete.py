import cv2
import face_recognition
import numpy as np
import os

# Function to capture faces and create training data
def capture_faces_and_train(known_face_encodings, known_face_names):
    while True:
        ret, frame = cap.read()

        # Detect faces in the frame
        face_locations = face_recognition.face_locations(frame)
        face_encodings = face_recognition.face_encodings(frame, face_locations)

        for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
            # Draw a rectangle around the face
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)

            # Get a name for the person if it's known
            name = "Unknown"

            # Check if the face encoding is in the known face encodings
            if face_encoding in known_face_encodings:
                first_match_index = known_face_encodings.index(face_encoding)
                name = known_face_names[first_match_index]

                # Optional: Update attendance list
                if name not in attendance_list:
                    attendance_list.append(name)
                    print(f"{name} marked present at {datetime.datetime.now().strftime('%H:%M:%S')}")

            # Display the name under the box (optional)
            cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
            font = cv2.FONT_HERSHEY_DUPLEX
            cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

        cv2.imshow('Facial Recognition Attendance System', frame)

        # Press 'q' to quit capturing faces and training
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # Release the capture and close all windows
    cap.release()
    cv2.destroyAllWindows()

    # Save attendance list (optional)
    if attendance_list:
        with open("attendance.txt", "a") as f:
            f.write(f"{datetime.datetime.now().strftime('%Y-%m-%d')}: {', '.join(attendance_list)}\n")

# Function to mark attendance and save to file
def mark_attendance(attendance_list):
    if attendance_list:
        with open("attendance.txt", "a") as f:
            f.write(f"{datetime.datetime.now().strftime('%Y-%m-%d')}: {', '.join(attendance_list)}\n")

# Import libraries
import cv2
import face_recognition
import numpy as np
import os
import datetime  # Added for date and time in attendance list

# Create variables
known_face_encodings = []
known_face_names = []
attendance_list = []  # Stores names of attendees for current session
cap = cv2.VideoCapture(0)  # Use 0 for default webcam

# Path to the folder containing known face images
known_faces_dir = "C:\\GCU\\Pyhton programs\\Project\\Face_recognition_attendance\\3rd_attendance(Main)\\known_faces"

# Loop through all the images in the known faces folder
for filename in os.listdir(known_faces_dir):
    # Load the image
    image = face_recognition.load_image_file(os.path.join(known_faces_dir, filename))

    # Encode the face
    face_encoding = face_recognition.face_encodings(image)[0]

    # Get the name of the person from the filename
    name = os.path.splitext(filename)[0]

    # Add the encoding and name to the known face arrays
    known_face_encodings.append(face_encoding)
    known_face_names.append(name)

# Option 1: Capture faces and train (uncomment to create initial training data)
capture_faces_and_train(known_face_encodings, known_face_names)

# Option 2: Use existing training data (replace with your actual path)
# known_face_encodings = np.load("known_face_encodings.npy")
