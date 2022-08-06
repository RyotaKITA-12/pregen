from pptx import Presentation
from pptx.util import Cm

presentation = Presentation()
slide_layout = presentation.slide_layouts[0]
slide = presentation.slides.add_slide(slide_layout)

pic = slide.shapes.add_picture('../data/unzipped/ppt/media/image1.png', Cm(5), Cm(5))
presentation.save('test_output.pptx')
