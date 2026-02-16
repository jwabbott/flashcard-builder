from pathlib import Path
import flashcard_builder as fb


IMG_DIR = Path(__file__).parent.parent / "test_headshots"


def test_normalize_filename_non_characters():							# checks whether spaces and dashes get replaced by underscores
	assert fb.normalize_filename("John Doe") == "john_doe"
	assert fb.normalize_filename("Jane-Smith.JPG") == "jane_smith_jpg"
	assert fb.normalize_filename(None) == ""

def test_normalize_filename_collapses_and_strips():						# tests that sequences of non-alphanumerics become single underscores,
	assert fb.normalize_filename("  .A--B__C!! ") == "a_b_c"			# and leading/trailing underscores are stripped	

		
def test_normalize_filename_numbers_and_hash():							# tests that numbers get kept, but that punctuation gets replaced
	assert fb.normalize_filename("Room#101-A") == "room_101_a"			# with underscores and gets collapsed
    
   


