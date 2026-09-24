"""Local face match against enrolled roster photos. Frames stay on this computer."""
from __future__ import annotations

import urllib.request
from pathlib import Path

import cv2
import numpy as np

ROOT = Path(__file__).resolve().parent
MODEL_DIR = ROOT / "models"
DETECTOR_NAME = "face_detection_yunet_2023mar.onnx"
RECOGNIZER_NAME = "face_recognition_sface_2021dec.onnx"
MODEL_URLS = {
    DETECTOR_NAME: "https://github.com/opencv/opencv_zoo/raw/main/models/face_detection_yunet/face_detection_yunet_2023mar.onnx",
    RECOGNIZER_NAME: "https://github.com/opencv/opencv_zoo/raw/main/models/face_recognition_sface/face_recognition_sface_2021dec.onnx",
}
# OpenCV SFace cosine threshold from the model card.
MATCH_THRESHOLD = 0.363


def ensure_models() -> tuple[Path, Path]:
    MODEL_DIR.mkdir(parents=True, exist_ok=True)
    paths = []
    for name in (DETECTOR_NAME, RECOGNIZER_NAME):
        path = MODEL_DIR / name
        if not path.exists() or path.stat().st_size < 1000:
            urllib.request.urlretrieve(MODEL_URLS[name], path)
        paths.append(path)
    return paths[0], paths[1]


def decode_image(blob: bytes) -> np.ndarray:
    array = np.frombuffer(blob, dtype=np.uint8)
    image = cv2.imdecode(array, cv2.IMREAD_COLOR)
    if image is None:
        raise ValueError("Could not read the image")
    return image


def _engines(width: int, height: int):
    detector_path, recognizer_path = ensure_models()
    detector = cv2.FaceDetectorYN.create(str(detector_path), "", (width, height), 0.5, 0.3, 5000)
    detector.setInputSize((width, height))
    recognizer = cv2.FaceRecognizerSF.create(str(recognizer_path), "")
    return detector, recognizer


def feature_for_image(image: np.ndarray) -> np.ndarray | None:
    height, width = image.shape[:2]
    detector, recognizer = _engines(width, height)
    _, faces = detector.detect(image)
    if faces is None or len(faces) == 0:
        return None
    face = faces[0]
    aligned = recognizer.alignCrop(image, face)
    return recognizer.feature(aligned)


def match_frame(frame: np.ndarray, enrolled: list[tuple[int, str, str, np.ndarray]]) -> dict | None:
    """Return the best enrolled match above the SFace threshold, or None."""
    height, width = frame.shape[:2]
    detector, recognizer = _engines(width, height)
    _, faces = detector.detect(frame)
    if faces is None or len(faces) == 0:
        return None
    best = None
    for face in faces:
        aligned = recognizer.alignCrop(frame, face)
        feature = recognizer.feature(aligned)
        for person_id, name, category, reference in enrolled:
            score = float(recognizer.match(feature, reference, cv2.FaceRecognizerSF_FR_COSINE))
            if best is None or score > best["score"]:
                best = {"person_id": person_id, "name": name, "category": category, "score": score}
    if best is None or best["score"] < MATCH_THRESHOLD:
        return None
    return best
