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

# Normalize: removes prefix like "1-" and extension
def normalize(name):
    name = os.path.splitext(name)[0]
    if '-' in name and name.split('-')[0].isdigit():
        name = name.split('-', 1)[1]
    return name.strip().lower()

def main():
  folder_left = "../501 Photos Initiales"
  folder_right = "../501"

  left_images = get_sorted_images(folder_left)
  right_images = get_sorted_images(folder_right)
  left_map = {normalize(os.path.basename(f)): f for f in left_images}
  right_map = {normalize(os.path.basename(f)): f for f in right_images}

  # Keys to generate slides for: all in right (even if left is missing)
  all_keys = sorted(right_map.keys())
  
  prs = Presentation()
  prs.slide_width = Inches(12.01)
  prs.slide_height = Inches(8.47)

  for key in all_keys:
    left_path = left_map.get(key)
    right_path = right_map.get(key)

    slide_title = os.path.splitext(os.path.basename(right_path))[0]

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

    if left_path and right_path:
      # Add left image
      top = Inches(1.31)
      height = Inches(6.4)
      slide.shapes.add_picture(left_path, Inches(0.8), top, height=height)

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
      slide.shapes.add_picture(right_path, Inches(6.4), top, height=height)

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
    elif right_path:
       # Add right image
      slide.shapes.add_picture(right_path, Inches(3.605), top, height=height)

      # Add 2025
      right_shape = slide.shapes.add_textbox(Inches(0), text_top, Inches(12.01), Inches(0.59))
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