from pptx import Presentation
import os
import tempfile
import streamlit as st

def generate_file():
    #Hello World! example
    #../_images/hello-world.png
    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "Hello, World!"
    subtitle.text = "python-pptx was here!"

    #Bullet slide example
    #../_images/bullet-slide.png
    bullet_slide_layout = prs.slide_layouts[1]

    slide = prs.slides.add_slide(bullet_slide_layout)
    shapes = slide.shapes

    title_shape = shapes.title
    body_shape = shapes.placeholders[1]

    title_shape.text = 'Adding a Bullet Slide'

    tf = body_shape.text_frame
    tf.text = 'Find the bullet slide layout'

    p = tf.add_paragraph()
    p.text = 'Use _TextFrame.text for first bullet'
    p.level = 1

    p = tf.add_paragraph()
    p.text = 'Use _TextFrame.add_paragraph() for subsequent bullets'
    p.level = 2

    return prs

# save locally
# prs = generate_file()
# prs.save('test.pptx')

# save on streamlit cloud
@st.cache_resource(ttl="1h")
def save_file():
    temp_dir = tempfile.TemporaryDirectory()
    temp_filepath = os.path.join(temp_dir.name, 'test.pptx')
    prs = generate_file()
    prs.save(temp_filepath)
    return temp_filepath

temp_filepath = save_file()
with open(temp_filepath, "rb") as template_file:
        template_byte = template_file.read()

st.download_button(label="Click to Download Template File",
                    data=template_byte,
                    file_name="test.pptx",
                    mime='application/octet-stream')