from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier('haarcascade_frontalface_default.xml')

# Load face data, names, and IDs
with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open('data/ids.pkl', 'rb') as i:
    IDS = pickle.load(i)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print('Shape of Faces matrix --> ', FACES.shape)

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

imgBackground = cv2.imread("bg.png")

COL_NAMES = ['ID', 'NAME', 'TIME', 'ATTENDANCE']
absent_ids = set()
while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    
    for (x, y, w, h) in faces:
        # Crop the detected face
        crop_img = frame[y:y+h, x:x+w, :]
        
        # Resize the cropped image to match the expected input size for KNN
        resized_img = cv2.resize(crop_img, (75, 75)).flatten().reshape(1, -1)
        
        # Perform KNN prediction (output is the predicted name)
        output = knn.predict(resized_img)
        
        # Find the ID associated with the predicted name
        name = output[0]
        user_id = IDS[LABELS.index(name)]
        
        # Get the timestamp
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")
        
        # Draw rectangles and text on the frame
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x + w, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        # cv2.putText(frame, "Name:" + str(name), (x, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        cv2.putText(frame, "Name:" + str(name), (x, y + h + 20), cv2.FONT_HERSHEY_COMPLEX, 1, (0, 0, 255), 2)
        cv2.putText(frame, "Roll no:" + str(user_id), (x, y + h + 45), cv2.FONT_HERSHEY_COMPLEX, 1, (0, 0, 255), 2)

        
        # Create the attendance list (add 'Present' dynamically)
        attendance = [str(user_id), str(name), str(timestamp), 'Present']
    
    # Display the frame on the background image
    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame", imgBackground)
    
    
    
    # Capture key press
    k = cv2.waitKey(1)
    if k == ord('o'):
        speak("Attendance Taken..")
        time.sleep(5)
        if exist:
            with open("Attendance/Attendance_" + date + ".csv", "+a", newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance)
        else:
            with open("Attendance/Attendance_" + date + ".csv", "+a", newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance)
    if k == ord('q'): 
        # End the attendance-taking loop
        date = datetime.now().strftime("%d-%m-%Y")
        attendance_file = "Attendance/Attendance_" + date + ".csv"

        # Read the CSV file to find which users are already marked present
        present_ids = set()  # Use a set for fast lookup of IDs
        if os.path.isfile(attendance_file):
            with open(attendance_file, 'r') as csvfile:
                reader = csv.reader(csvfile)
                next(reader)  # Skip the header row
                for row in reader:
                    present_ids.add(row[0])  # Add the ID to the set

        # Now iterate over all registered users
        with open(attendance_file, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            for i, user_id in enumerate(IDS):
                if user_id not in present_ids and user_id not in absent_ids:
                    # If user ID not in present_ids and not marked absent, mark as absent 
                    name = LABELS[i]  # Get the name corresponding to the ID
                    timestamp = datetime.now().strftime("%H:%M:%S")  # Mark the current time
                    writer.writerow([user_id, name, timestamp, "Absent"])
                    absent_ids.add(user_id)  # Add to absent_ids set to avoid duplicates
                    print(f"{name} (ID: {user_id}) marked as Absent.")

        break
video.release()
cv2.destroyAllWindows()






