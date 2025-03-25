from ultralytics import YOLO

# Load YOLOv8 model (pre-trained or custom trained)
model = YOLO("yolov8n.pt")  

# Run detection on the whiteboard image
def detect_whiteboard_objects(image_path):
    results = model(image_path)  # Detect elements in the image
    return results
