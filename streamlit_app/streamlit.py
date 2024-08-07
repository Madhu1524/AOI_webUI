import cv2

cap = cv2.VideoCapture(0)

if not cap.isOpened():
    print("Error: Couldn't open webcam.")
else:
    print("Webcam opened successfully.")
    ret, frame = cap.read()
    if ret:
        cv2.imshow('Webcam Test', frame)
        cv2.waitKey(0)
    cap.release()
    cv2.destroyAllWindows()
