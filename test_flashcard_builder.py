from pathlib import Path
import flashcard_builder as fb


IMG_DIR = Path(__file__).parent.parent / "test_headshots"


def test_normalize_filename():
    assert fb.normalize_filename("John Doe") == "john_doe"
    assert fb.normalize_filename("Jane-Smith.JPG") == "jane_smith_jpg"
    assert fb.normalize_filename(None) == ""


def test_build_image_index():
    idx = fb.build_image_index(IMG_DIR)

    # Ensure dictionary was created
    assert isinstance(idx, dict)

    # Ensure at least one file was indexed
    assert len(idx) > 0

