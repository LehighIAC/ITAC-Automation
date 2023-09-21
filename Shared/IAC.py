"""
(Purpose) IAC.py is a module that contains functions used in the IAC report
"""

import math, os, locale
#from easydict import EasyDict
from lxml import etree
import latex2mathml.converter

def grouping_num(dic):
    """
    Add thousand separator to numbers in a dictionary and format it to string
    :param dic: EasyDict
    :return: Dictionary with keys in thousand separator
    """
    # set locale to US
    locale.setlocale(locale.LC_ALL, 'en_US')
    for key in dic.keys():
        if type(dic[key]) == int:
            dic[key] = locale.format_string('%d', dic[key], grouping=True)
        elif type(dic[key]) == float:
            dic[key] = locale.format_string('%g', dic[key], grouping=True)
        else:
            pass
    return dic

def dollar(varlist, dic, digits):
    """
    Format numbers in a dictionary and to currency string
    :param varlist: List of keys in the dictionary
    :param dic: EasyDict
    :param digits: Number of digits
    :return: Dictionary with keys in formatted currency string
    """
    locale.setlocale(locale.LC_ALL, 'en_US')
    locale._override_localeconv={'frac_digits':digits}
    for var in varlist:
        dic[var] = locale.currency(dic[var], grouping=True)
    return dic

def combine_words(words):
    """
    :param words: list of strings
    :return: string of words separated by "," and "and"
    """
    combined = ""
    for i in range(len(words)):  
        combined = combined + words[i]    
        if i < len(words) - 2:
            combined = combined + ', '
        if i == len(words) - 2:
            combined = combined + ' and ' 
        else:
            pass
    return combined

def add_image(doc, tag, image_path, wd):
    """
    Add image to Word document, search for tag in doc and replace with the image
    :param doc: Document
    :param tag: Image tag string
    :param image_path: Path to the image
    :param wd: Image width
    :return: None
    """
    found_tag = False
    for p in doc.paragraphs:
        if tag in p.text:
            p.text = p.text.replace(tag, '')
            r = p.add_run()
            r.add_picture(image_path, width=wd)
            found_tag = True
            break
    if found_tag == False:
        print("Tag "+ tag +" not found")

def add_eqn(doc, tag, eqn_input):
    """
    Add equation to Word document, search for eqn in doc and replace with eqn_input
    :param doc: Document
    :param tag: Equation tag string
    :param eqn_input: Equation input
    :return: None
    """
    found_tag = False
    for p in doc.paragraphs:
        if tag in p.text:
            p.text = p.text.replace(tag, '')
            word_math = latex2word(eqn_input)
            p._element.append(word_math)
            found_tag = True
            break
    if found_tag == False:
        print("Tag "+ tag +" not found")

def latex2word(latex_input):
    """
    Convert LaTeX equation to Word equation
    :param latex_input: LaTeX equation
    :return: Word equation
    """
    mathml = latex2mathml.converter.convert(latex_input)
    tree = etree.fromstring(mathml)
    script_path = os.path.dirname(os.path.abspath(__file__))
    xslt = etree.parse(os.path.join(script_path,'..','Shared','MML2OMML.XSL'))
    transform = etree.XSLT(xslt)
    new_dom = transform(tree)
    return new_dom.getroot()

def payback(ACS,IC):
    """
    Formet payback period by year and month
    :param ACS: Annual Cost Savings ($/yr)
    :param IC: Implementation Cost ($)
    :return: formatted Payback Period (str)
    """
    PB = IC / ACS
    if PB <= 11.0 / 12.0:
        PB = math.ceil(PB * 12.0)
        PBstr = str(PB) + " month"
    else:
        PB = math.ceil(PB * 10.0) / 10.0
        PBstr = str(PB).rstrip("0").rstrip(".") + " year"
    if PB > 1.0:
        PBstr = PBstr + "s"
    return PBstr