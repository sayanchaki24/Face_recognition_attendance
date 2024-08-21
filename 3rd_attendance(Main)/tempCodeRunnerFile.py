import face_recognition
import cv2
import pandas as pd
import openpyxl

# Load the face recognition model
face_recognition_model = face_recognition.FaceRecognition()

# Load the known faces (students)
known_faces = []
known_face_names = []
with open("known_faces.txt", "r") as f:
    for line in f:
        name, face_encoding = line.strip().split(",")
        known_faces.append(face_encoding)
        known_face_names.append(name)

# Create an Excel file to store the attendance data
attendance_file = "attendance.xlsx"
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Attendance"
sheet["A1"] = "Date"
sheet["B1"] = "Name"
sheet["C1"] = "Attendance"

# Set up the camera
cap = cv2.VideoCapture(0)

while True:
    # Capture a frame from the camera
    ret, frame = cap.read()
    if not ret:
        break

    # Convert the frame to grayscale
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    # Detect faces in the frame
    face_locations = face_recognition.face_locations(gray)
    face_encodings = face_recognition.face_encodings(gray, face_locations)

    # Loop through each face
    for face_encoding, face_location in zip(face_encodings, face_locations):
        # Compare the face encoding to the known faces
        matches = face_recognition.compare_faces(known_faces, face_encoding)
        name = "Unknown"
        for i, match in enumerate(matches):
            if match:
                name = known_face_names[i]
                break

        # Mark attendance if the face is recognized
        if name != "Unknown":
            date = datetime.date.today()
            attendance = "Present"
            sheet.append([date, name, attendance])
            wb.save(attendance_file)

    # Display the output
    cv2.imshow("Attendance", frame)
    if cv2.waitKey(1) & 0xFF == ord("q"):
        break

# Release the camera and close the Excel file
cap.release()
wb.close()
cv2.destroyAllWindows()