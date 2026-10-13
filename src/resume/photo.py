"""The CV's photo, cropped to a square around the face so every output frames it the same way."""

from pathlib import Path
from functools import cache

import cv2
import cv2.data

SIDE = 480  # px: sharp at 2x in the page and at 300 dpi in print
ZOOM = 1.8  # the square's side as a multiple of the face's width: head and shoulders, not the whole picture
MIN_ZOOM = 1.3  # the tightest crop worth making to keep the face centered


def face_center(img: cv2.typing.MatLike) -> tuple[float, float, float] | None:
    """The middle and width of the largest face, or None when there is none."""
    cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
    gray = cv2.equalizeHist(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY))
    faces = cascade.detectMultiScale(
        gray, scaleFactor=1.1, minNeighbors=6, minSize=(max(40, min(gray.shape) // 12),) * 2
    )
    if not len(faces):
        return None
    x, y, w, h = max(faces, key=lambda f: f[2] * f[3])
    return x + w / 2, y + h / 2, float(w)


def square(img: cv2.typing.MatLike) -> cv2.typing.MatLike:
    """A square centered exactly on the face (about ZOOM face widths wide), as large as fits around that center.

    A face near an edge gets a tighter square rather than being shifted off-center; the square never gets tighter
    than MIN_ZOOM face widths, below which the face is moved off-center instead. Without a face: the middle square.
    """
    height, width = img.shape[:2]
    face = face_center(img)
    cx, cy, w = face if face else (width / 2, height / 2, float(max(width, height)))
    fit = 2 * min(cx, width - cx, cy, height - cy)  # the biggest square centered on the face that stays in the picture
    side = round(min(width, height, w * ZOOM, max(fit, w * MIN_ZOOM)))
    left = round(min(max(cx - side / 2, 0), width - side))
    top = round(min(max(cy - side / 2, 0), height - side))
    return img[top : top + side, left : left + side]


@cache
def portrait(path: Path) -> bytes:
    """A JPEG of the photo at `path`, square around the face and no bigger than SIDE px."""
    img = cv2.imread(str(path))
    if img is None:
        raise ValueError(f"cannot read the photo {path}")
    crop = square(img)
    if crop.shape[0] > SIDE:
        crop = cv2.resize(crop, (SIDE, SIDE), interpolation=cv2.INTER_AREA)
    ok, jpeg = cv2.imencode(".jpg", crop, [cv2.IMWRITE_JPEG_QUALITY, 90])
    if not ok:
        raise ValueError(f"cannot encode the photo {path}")
    return jpeg.tobytes()
