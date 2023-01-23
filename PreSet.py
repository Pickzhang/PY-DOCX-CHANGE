from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn


def Fontset(document):
    style_family = document.styles.add_style('SimHei', WD_STYLE_TYPE.CHARACTER)
    style_family.font.name = '黑体'
    document.styles['SimHei']._element.rPr.rFonts.set(qn('w:eastAsia'), u'黑体')
    style_family = document.styles.add_style('JianBiaoSong', WD_STYLE_TYPE.CHARACTER)
    style_family.font.name = '创艺简标宋'
    document.styles['JianBiaoSong']._element.rPr.rFonts.set(qn('w:eastAsia'), u'创艺简标宋')
    style_family = document.styles.add_style('FangSong_GB2312', WD_STYLE_TYPE.CHARACTER)
    style_family.font.name = '仿宋_GB2312'
    document.styles['FangSong_GB2312']._element.rPr.rFonts.set(qn('w:eastAsia'), u'仿宋_GB2312')
    style_family = document.styles.add_style('KaiTi_GB2312', WD_STYLE_TYPE.CHARACTER)
    style_family.font.name = '楷体_GB2312'
    document.styles['KaiTi_GB2312']._element.rPr.rFonts.set(qn('w:eastAsia'), u'楷体_GB2312')
