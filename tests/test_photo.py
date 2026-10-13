import cv2
import numpy as np

from resume import ROOT
from resume.photo import SIDE
from resume.photo import square
from resume.photo import portrait
from resume.photo import face_center

PHOTO = ROOT / "assets" / "photo.jpg"


def read() -> np.ndarray:
    img = cv2.imread(str(PHOTO))
    assert img is not None
    return img


def test_finds_the_face() -> None:
    cx, cy, w = face_center(read()) or (0, 0, 0)
    assert 400 < cx < 800
    assert 300 < cy < 700
    assert w > 300


def test_crop_follows_the_face_in_an_off_centre_picture() -> None:
    img = cv2.resize(read(), (500, 500))
    wide = cv2.copyMakeBorder(img, 0, 0, 900, 200, cv2.BORDER_CONSTANT, value=(128, 128, 128))
    crop = square(wide)
    assert crop.shape[0] == crop.shape[1]
    assert crop.shape[0] <= 500
    assert face_center(crop) is not None  # the face survived the crop, though it sat far from the middle


def test_no_face_falls_back_to_the_middle_square() -> None:
    blank = np.full((300, 500, 3), 90, dtype=np.uint8)
    assert square(blank).shape[:2] == (300, 300)


def test_portrait_is_a_small_square_jpeg() -> None:
    jpeg = portrait(PHOTO)
    img = cv2.imdecode(np.frombuffer(jpeg, dtype=np.uint8), cv2.IMREAD_COLOR)
    assert img is not None
    assert jpeg[:2] == b"\xff\xd8"
    assert img.shape[0] == img.shape[1] <= SIDE


def test_square_is_centered_on_the_face_when_it_fits() -> None:
    img = read()
    cx, cy, _ = face_center(img) or (0, 0, 0)
    height, width = img.shape[:2]
    crop = square(img)
    # find the crop's top-left by matching it back into the picture
    hit = cv2.minMaxLoc(cv2.matchTemplate(img, crop, cv2.TM_SQDIFF))[2]
    assert abs(hit[0] + crop.shape[1] / 2 - cx) <= 2
    assert abs(hit[1] + crop.shape[0] / 2 - cy) <= 2
    assert crop.shape[0] <= min(width, height)
