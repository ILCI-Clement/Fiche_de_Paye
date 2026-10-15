"""Regression coverage for landscape PDF coordinates and ambiguous anchors."""
import unittest
from unittest.mock import patch
from pypdf.generic import RectangleObject
from fastapi import HTTPException
from backend.signature_flow import attendance_signature_fields
from backend.signature_placement import place_signature_fields


class Page:
    def __init__(self, *, duplicate=False, missing=False, rotation=0):
        self.mediabox=RectangleObject([0,0,842,595])
        self.cropbox=RectangleObject([0,0,842,595])
        self.rotation=rotation
        self.duplicate,self.missing=duplicate,missing
    def extract_text(self,visitor_text):
        for f,x in zip(attendance_signature_fields(),[290,16]):
            if self.missing and f['party']==1:continue
            visitor_text(f['anchor'],[1,0,0,1,0,0],[1,0,0,1,x,244],None,6.5)
            if self.duplicate:visitor_text(f['anchor'],[1,0,0,1,0,0],[1,0,0,1,x,244],None,6.5)


class SignaturePlacementTests(unittest.TestCase):
    def place(self,page):
        with patch('backend.signature_placement.PdfReader') as reader:
            reader.return_value.pages=[page]
            return place_signature_fields(b'pdf-test',attendance_signature_fields())
    def test_landscape_coordinates_reconstruct_original_blank_area(self):
        fields=self.place(Page())
        for field,x in zip(fields,[290,16]):
            self.assertAlmostEqual(field['x']*842/595,x,places=2)
            self.assertAlmostEqual(field['y']*595/842,361.8,places=2)
            self.assertAlmostEqual(field['width']*842/595,150,places=2)
            self.assertAlmostEqual(field['height']*595/842,70,places=2)
            self.assertNotIn('anchor',field)
    def test_duplicate_labels_are_rejected(self):
        with self.assertRaises(HTTPException):self.place(Page(duplicate=True))
    def test_missing_label_is_rejected(self):
        with self.assertRaises(HTTPException):self.place(Page(missing=True))
    def test_rotated_page_is_rejected_instead_of_guessing(self):
        with self.assertRaises(HTTPException):self.place(Page(rotation=90))


if __name__=='__main__':unittest.main()
