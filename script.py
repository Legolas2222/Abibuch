import pandas as pd
from pptx import Presentation

# Load Excel data
df = pd.read_excel('antworten.xlsx', sheet_name='data')

# Load PowerPoint template
prs = Presentation('buchtemplate_single.pptx')
template_slide = prs.slides[0]  # Assuming the first slide is the template

def duplicate_and_fill_slide(data):
    # Duplicate template slide
    slide = prs.slides.add_slide(template_slide.slide_layout)
    
    # Copy shapes from the template
    for shape in template_slide.shapes:
        if shape.has_text_frame:
            new_shape = slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
            new_shape.text_frame.text = shape.text_frame.text

    # Replace placeholders
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text
            for key, value in data.items():
                placeholder = f"{{{key}}}"
                text = text.replace(placeholder, str(value))
            shape.text_frame.text = text

# For each row in Excel, create a new slide
for index, row in df.iterrows():
    duplicate_and_fill_slide(row)

# Remove the original template slide (optional)
# del prs.slides[0]  # Remove the first slide if it was a template
# Save the final presentation
prs.save('filled_presentation.pptx')
print("Presentation generated!")
