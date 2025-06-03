from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import os

# Filter image files and sort
def get_sorted_images(folder):
  return sorted([
    os.path.join(folder, f)
    for f in os.listdir(folder)
    if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif'))
  ])

def main():
  folder_left = "../501 Photos Initiales"
  folder_right = "../501"

  left_images = get_sorted_images(folder_left)
  right_images = get_sorted_images(folder_right)
  pairs = zip(left_images, right_images)
  
  prs = Presentation()
  prs.slide_width = Inches(12.01)
  prs.slide_height = Inches(8.47)

  for left_image_path, right_image_path in pairs:
    # Get image names (without extensions) for title
    name = os.path.splitext(os.path.basename(right_image_path))[0]
    slide_title = f"{name}"

    # Add a new slide (Title and Content layout)
    slide_layout = prs.slide_layouts[6]  # Use blank layout
    slide = prs.slides.add_slide(slide_layout)

    # Add background
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)  # Black

    # Add title
    title_shape = slide.shapes.add_textbox(Inches(0), Inches(0.1), Inches(12.01), Inches(0.59))
    text_frame = title_shape.text_frame
    p = text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = slide_title
    font = run.font
    font.name = 'Comic Sans MS'
    font.size = Pt(60)
    font.color.rgb = RGBColor(0xB9, 0xB4, 0x53)  # Hex #B9B453

    # Add left image
    top = Inches(1.31)
    height = Inches(6.4)
    slide.shapes.add_picture(left_image_path, Inches(0.8), top, height=height)

    text_top = Inches(7.71)
    # Add 2020
    left_shape = slide.shapes.add_textbox(Inches(0.8), text_top, Inches(5.12), Inches(0.59))
    text_frame = left_shape.text_frame
    p = text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "2020"
    font = run.font
    font.name = 'Comic Sans MS'
    font.size = Pt(32)
    font.color.rgb = RGBColor(0xB9, 0xB4, 0x53)  # Hex #B9B453

    # Add right image
    slide.shapes.add_picture(right_image_path, Inches(6.4), top, height=height)

    # Add 2025
    right_shape = slide.shapes.add_textbox(Inches(6.4), text_top, Inches(4.8), Inches(0.59))
    text_frame = right_shape.text_frame
    p = text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "2025"
    font = run.font
    font.name = 'Comic Sans MS'
    font.size = Pt(32)
    font.color.rgb = RGBColor(0xB9, 0xB4, 0x53)  # Hex #B9B453

  # Save the presentation
  prs.save("Presentation.pptx")

if __name__ == "__main__":
  main()