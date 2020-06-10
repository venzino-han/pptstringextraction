# string extraction
from pptx import Presentation
import pandas as pd
import re
import os
from pptx.enum.shapes import MSO_SHAPE_TYPE


def getGroupText(groupShape,strings):
    for shape in groupShape.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            getGroupText(shape,strings)
        else:
            if hasattr(shape, "text"):
                str = re.findall(r'“(.*?)”', shape.text)
                strings += str
    return strings


def extraction(filename):
    prs = Presentation(filename)
    result = []
    i = 0
    for slide in prs.slides:
        i += 1
        if len(slide.shapes) == 0:
            continue
        else:
            title = slide.shapes.title.text
            for shape in slide.shapes:
                if shape.has_table:
                    table = shape.table
                    for r in table.rows:
                        for c in r.cells:
                            txt = ''
                            for t in c.text_frame.paragraphs:
                                txt += t.text
                            strings = re.findall(r'“(.*?)”', txt)
                            for s in strings:
                                result.append([i, title, s])

                if hasattr(shape, "text"):
                    strings = re.findall(r'“(.*?)”', shape.text)
                    for s in strings:
                        result.append([i, title, s])

                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    strings = []
                    strings = getGroupText(shape,strings)
                    for s in strings:
                        result.append([i, title, s])

    df = pd.DataFrame(result)
    newFileName = filename[:-4]+'csv'
    os.remove(filename)
    df.to_csv(newFileName, encoding="euc-kr")
    return newFileName