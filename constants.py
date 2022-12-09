from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT, MSO_VERTICAL_ANCHOR
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
import os.path

#os paths
jy_icon_path_dark = os.path.join("resources", "JY-Icon-Dark.png")
database_path = os.path.join("resources", "database.csv")
presentation_path_dark = os.path.join("resources", "MusicSlidesTemplateDark.pptx")

#TitleSlide title constants
title_fontName = 'Calibri Light (Headings)'
title_fontSize = 72
title_lineSpacing = 0.9
title_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
title_VAlign = MSO_VERTICAL_ANCHOR.BOTTOM
title_left = Cm(4.23)
title_top = Cm(3.12)
title_width = Cm(25.4)
title_height = Cm(6.63)
title_limit = 25

#TitleSlide JY Icon constants
title_JYIcon_left = Cm(15.61)
title_JYIcon_top = Cm(11.96)
title_JYIcon_height = Cm(2.64)

#TitleSlide subtitle constants
subTitle_fontName = 'Calibri (Body)'
subTitle_fontSize = 24
subTitle_lineSpacing = 0.9
subTitle_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
subTitle_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
subTitle_left = Cm(4.23)
subTitle_top = Cm(14.97)
subTitle_width = Cm(25.4)
subTitle_height = Cm(1.25)
subTitle_text = "Jesus Youth Australia"

#normalSlide JY Icon constants
norm_JYIcon_left = Cm(16.23)
norm_JYIcon_top = Cm(17.37)
norm_JYIcon_height = Cm(1.4)

#indexSlide title constants
indexHeading_fontName = 'Calibri Light (Headings)'
indexHeading_fontSize = 48
indexHeading_lineSpacing = 0.9
indexHeading_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
indexHeading_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
indexHeading_text = "Index"
indexHeading_left = Cm(13.75)
indexHeading_top = Cm(0.54)
indexHeading_width = Cm(6.22)
indexHeading_height = Cm(2.13)

#indexSlide songName constants
indexList_fontName = 'Calibri (Body)'
indexList_fontSize = 24
indexList_colorDark = RGBColor(255, 255, 255)
indexList_colorLight = RGBColor(0, 0, 0)
indexList_HAlign = PP_PARAGRAPH_ALIGNMENT.LEFT
indexList_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
indexList_Shape = MSO_SHAPE_TYPE.AUTO_SHAPE
indexList_left = [Cm(2.48), Cm(17.08)]
indexList_top = [Cm(3.55), Cm(4.77), Cm(6), Cm(7.23), Cm(8.45), Cm(9.68), Cm(10.9), Cm(12.13), Cm(13.36), Cm(14.58)]
indexList_width = Cm(14.3)
indexList_height = Cm(1.04)
indexList_limit = 28

#songNameSlide constants
songTitles_fontName = 'Calibri Light (Headings)'
songTitles_fontSize = 80
songTitles_lineSpacing = 0.9
songTitles_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
songTitles_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
songTitles_left = Cm(2.27)
songTitles_top = Cm(4.48)
songTitles_width = Cm(29.21)
songTitles_height = Cm(9.68)
songTitles_limit = 25

#songVerseSlide constants
verse_fontName = 'Calibri (Body)'
verse_HAlign = PP_PARAGRAPH_ALIGNMENT.CENTER
verse_VAlign = MSO_VERTICAL_ANCHOR.MIDDLE
verse_left = Cm(2.42)
verse_top = Cm(2)
verse_width = Cm(29.01)
verse_height = Cm(14.5)
#fontSize based on line count
verse_fontSize_lines5 = 60
verse_fontSize_lines6 = 56
verse_fontSize_lines7 = 51
#characterLimit based on line count
verse_charLimit_lines5 = 33
verse_charLimit_lines6 = 33
verse_charLimit_lines7 = 38