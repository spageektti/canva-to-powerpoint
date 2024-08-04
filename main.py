from pptx import Presentation
from pptx.util import Inches
import os

# path to the directory containing the GIF files (current working directory)
gif_directory = os.getcwd()

# create a new PowerPoint presentation
presentation = Presentation()

# set the slide dimensions to 1920x1080
presentation.slide_width = Inches(20)
presentation.slide_height = Inches(11.25)

print(f"Creating PowerPoint presentation with dimensions 1920x1080.")
print(f"Reading GIF files from: {gif_directory}")

# loop through the files in the directory
gif_files = [f for f in os.listdir(gif_directory) if f.endswith('.gif')]
gif_files.sort(key=lambda f: int(f.split('.')[0]))  # Ensure correct order

for gif_file in gif_files:
    gif_path = os.path.join(gif_directory, gif_file)
    if os.path.isfile(gif_path):
        print(f"Adding {gif_file} to slide.")
        slide_layout = presentation.slide_layouts[5]  # 5 corresponds to a blank slide layout
        slide = presentation.slides.add_slide(slide_layout)
        
        # add the GIF to the slide
        left = Inches(0)
        top = Inches(0)
        width = presentation.slide_width
        height = presentation.slide_height
        slide.shapes.add_picture(gif_path, left, top, width, height)

# save the presentation
output_path = os.path.join(gif_directory, 'output.pptx')
presentation.save(output_path)

print(f"Presentation saved as {output_path}")
