import pptx


def main(path):
    ppt = pptx.Presentation(path)

    for i, slide in enumerate(ppt.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                print("===")
                print(shape.text)


if __name__ == '__main__':
    path = '../data/test.pptx'
    main(path)
